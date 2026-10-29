"""Chromium integration tests, isolated from Render and GitHub."""
import copy
import io
import os
from pathlib import Path
import tempfile
import threading
import unittest

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright
from werkzeug.serving import make_server
import app as poker


class BrowserTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.original = copy.deepcopy(poker.STATE)
        cls.tmp = tempfile.TemporaryDirectory()
        cls.exports = poker.EXPORTS_DIR
        poker.EXPORTS_DIR = Path(cls.tmp.name)
        cls.server = make_server("127.0.0.1", 0, poker.app, threaded=True)
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()
        cls.url = "http://127.0.0.1:" + str(cls.server.server_port)
        cls.pw = sync_playwright().start()
        options = {"headless": True}
        if os.environ.get("POKER_BROWSER"):
            options["executable_path"] = os.environ["POKER_BROWSER"]
        cls.browser = cls.pw.chromium.launch(**options)

    @classmethod
    def tearDownClass(cls):
        cls.browser.close()
        cls.pw.stop()
        cls.server.shutdown()
        cls.thread.join()
        poker.EXPORTS_DIR = cls.exports
        cls.tmp.cleanup()
        poker.STATE.clear()
        poker.STATE.update(cls.original)

    def setUp(self):
        self.context = self.browser.new_context()
        self.addCleanup(self.context.close)
        self.page = self.context.new_page()
        self.page.set_default_timeout(5000)
        # Install before navigation so the application creates only controlled timers.
        self.page.clock.install()
        self.page.clock.pause_at(self.page.evaluate("Date.now()") / 1000 + 1)
        self.errors = []
        self.page.on("pageerror", lambda error: self.errors.append(str(error)))

    def tearDown(self):
        self.assertEqual(self.errors, [])

    def prepare(self, language="fr", winner="Alice"):
        self.context.request.get(self.url + "/set-lang/" + language)
        p = self.page
        p.goto(self.url + "/setup")
        p.locator("#players").fill("3")
        p.locator("#players").press("Tab")
        for i, name in enumerate([winner, "Bob", "Chloé"], 1):
            p.locator('[name="player_'+str(i)+'"]').fill(name)
        p.locator('[name="duration"]').fill("60")
        p.locator('[name="buyin_money"]').fill("50000")
        p.locator('[name="buyin_chips"]').fill("1000")
        p.locator("#paid_places").fill("2")
        p.locator("#paid_places").press("Tab")
        p.locator('[name="winner_pct"]').fill("70")
        p.locator('button[type="submit"]').click()
        p.locator('a[href="/clock"]').click()
        self.assertEqual(p.locator("#timer").inner_text(), "07:00")
        self.assertEqual(p.locator("html").get_attribute("lang"), language)
        self.assertEqual(p.locator("h1").first.inner_text(), poker.TEXTS[language]["clock"])

    def kill(self, name, outcome, killer="Alice"):
        p = self.page
        p.locator("button.kill").click()
        p.select_option("#killedPlayer", name)
        p.select_option("#killerPlayer", killer)
        p.select_option("#outcome", outcome)
        with p.expect_response("**/api/kill") as response:
            p.locator('[onclick="submitKill()"]').click()
        self.assertTrue(response.value.json()["ok"])
        p.wait_for_function("!killPending && !killModalOpen")

    def test_full_game_in_fr_es_en(self):
        for language in ("fr", "es", "en"):
            with self.subTest(language=language):
                self.prepare(language)
                p = self.page
                p.clock.run_for(1000)
                p.locator("button.kill").click()
                saved = p.locator("#timer").inner_text()
                p.clock.run_for(5000)
                self.assertEqual(p.locator("#timer").inner_text(), saved)
                p.locator('[onclick="closeKillModal()"]').click()
                p.clock.run_for(1000)
                self.assertNotEqual(p.locator("#timer").inner_text(), saved)
                p.locator("#pauseBtn").click()
                self.kill("Bob", "rebuy")
                self.assertFalse(p.evaluate("running"))
                p.locator("#pauseBtn").click()
                self.kill("Chloé", "out")
                self.assertTrue(p.evaluate("running"))
                saved = p.evaluate("computeRemainingSeconds()")
                p.clock.run_for(1000)
                self.assertEqual(p.evaluate("computeRemainingSeconds()"), saved - 1)
                self.kill("Bob", "out")
                p.wait_for_selector("#finalModal", state="visible")
                self.assertFalse(p.evaluate("running"))
                self.assertIsNone(p.evaluate("timerHandle"))
                self.assertIn(poker.TEXTS[language]["winner"], p.locator("#winnerText").inner_text())
                self.assertIn("Alice", p.locator("#winnerText").inner_text())
                self.assertEqual(p.locator(".payout-line").count(), 2)
                self.assertTrue(p.locator("#fireworks").is_visible())
                self.assertEqual(p.evaluate("getComputedStyle(document.getElementById('fireworks'), '::before').animationName"), "firework-burst")
                with p.expect_download() as download:
                    p.locator('#finalModal a[href="/export"]').click()
                wb = load_workbook(io.BytesIO(Path(download.value.path()).read_bytes()))
                self.assertEqual(wb["Gains"]["B1"].value, 200000)
                self.assertEqual(wb["Gains"]["D8"].value, 140000)
                wb.close()
                if os.environ.get("POKER_SCREENSHOTS"):
                    dest = Path(os.environ["POKER_SCREENSHOTS"])
                    p.screenshot(path=str(dest / ("final-" + language + ".png")))

    def test_five_second_alarm_levels_and_live_language(self):
        self.prepare()
        p = self.page
        p.evaluate("() => {window.alarmPlays=0; alarmUnlocked=true; document.getElementById('alarmSound').play=()=>{window.alarmPlays++;return Promise.resolve();};}")
        p.clock.run_for(414000)
        self.assertEqual(p.evaluate("alarmPlays"), 0)
        p.clock.run_for(1000)
        self.assertEqual(p.evaluate("alarmPlays"), 1)
        p.clock.run_for(5000)
        p.wait_for_function("state.current_level === 2 && !changingLevel")
        self.assertEqual(p.evaluate("alarmPlays"), 1)
        self.assertEqual(p.locator("#blinds").inner_text(), poker.STATE["schedule"][1])
        p.locator('[onclick="previousLevel()"]').click()
        p.wait_for_function("state.current_level === 1 && !changingLevel")
        p.locator("#pauseBtn").click()
        remaining = p.evaluate("computeRemainingSeconds()")
        for code in ("es", "en", "fr"):
            p.locator(".lang-bar a", has_text=code.upper()).click()
            p.wait_for_function("!languagePending && document.documentElement.lang === '"+code+"'")
            self.assertFalse(p.evaluate("running"))
            self.assertEqual(p.evaluate("computeRemainingSeconds()"), remaining)
        p.locator("#pauseBtn").click()
        deadline = p.evaluate("levelEndTimestamp")
        p.locator(".lang-bar a", has_text="EN").click()
        p.wait_for_function("!languagePending && document.documentElement.lang === 'en'")
        self.assertEqual(p.evaluate("levelEndTimestamp"), deadline)

    def test_network_error_cancel_and_level_recovery(self):
        self.prepare()
        p = self.page
        p.route("**/api/kill", lambda route: route.abort())
        p.locator("button.kill").click()
        p.locator('[onclick="submitKill()"]').click()
        p.wait_for_function("!killPending && document.getElementById('killError').textContent.length > 0")
        p.locator('[onclick="closeKillModal()"]').click()
        self.assertTrue(p.evaluate("running"))
        p.route("**/api/level", lambda route: route.fulfill(status=503, body="<html>offline</html>"))
        p.locator('[onclick="nextLevel(false)"]').click()
        p.wait_for_function("!changingLevel && document.getElementById('clockError').textContent.length > 0")
        self.assertTrue(p.evaluate("running"))
        self.assertEqual(p.evaluate("state.current_level"), 1)
        p.unroute("**/api/level")
        p.locator('[onclick="nextLevel(false)"]').click()
        p.wait_for_function("!changingLevel && state.current_level === 2")

    def test_early_finish_and_final_request_retry(self):
        self.prepare()
        self.kill("Chloé", "out")
        p = self.page
        p.route("**/api/final", lambda route: route.abort())
        p.locator("button.export").click()
        p.wait_for_function("document.getElementById('clockError').textContent.length > 0")
        self.assertFalse(p.evaluate("running"))
        p.unroute("**/api/final")
        p.locator("button.export").click()
        p.wait_for_selector("#finalModal", state="visible")
        self.assertNotIn("Chloé", p.locator("#winnerText").inner_text())
        self.assertEqual(p.locator(".payout-line").count(), 0)
        self.assertFalse(p.locator("#fireworks").is_visible())

    def test_winner_name_is_text_and_mobile_final_is_accessible(self):
        name = '<img src=x onerror="window.injected=true">'
        self.prepare(winner=name)
        self.kill("Chloé", "out", killer=name)
        self.kill("Bob", "out", killer=name)
        p = self.page
        p.wait_for_selector("#finalModal", state="visible")
        self.assertEqual(p.locator("#payoutsBox img").count(), 0)
        self.assertIsNone(p.evaluate("window.injected"))
        p.set_viewport_size({"width":390,"height":844})
        self.assertTrue(p.locator('#finalModal a[href="/export"]').is_visible())
        self.assertTrue(p.locator(".final-box").evaluate("el => el.scrollWidth <= el.clientWidth"))
        if os.environ.get("POKER_SCREENSHOTS"):
            p.screenshot(path=str(Path(os.environ["POKER_SCREENSHOTS"]) / "final-mobile.png"))

    def test_setup_errors_preserve_input_and_allow_launch_in_all_languages(self):
        for language in ('fr', 'es', 'en'):
            with self.subTest(language=language):
                self.context.request.get(self.url + '/set-lang/' + language)
                p = self.page
                p.goto(self.url + '/setup')
                self.assertEqual(p.locator('[name="buyin_chips"]').input_value(), '1200')
                self.assertTrue(p.get_by_text(poker.TEXTS[language]['money_help']).is_visible())
                p.locator('[name="buyin_money"]').fill('50')
                name = 'David "D" <test>'
                p.locator('[name="player_1"]').fill(name)
                p.locator('#players').fill('7')
                p.locator('#players').press('Tab')
                self.assertEqual(p.locator('[name="player_1"]').input_value(), name)
                p.locator('[name="winner_pct"]').fill('80')
                with p.expect_response('**/setup') as response:
                    p.locator('button[type="submit"]').click()
                self.assertEqual(response.value.status, 400)
                self.assertEqual(p.locator('[role="alert"]').inner_text(), poker.TEXTS[language]['setup_payout'])
                self.assertEqual(p.locator('[name="player_1"]').input_value(), name)
                self.assertEqual(p.locator('[name="buyin_money"]').input_value(), '50')
                p.locator('[name="winner_pct"]').fill('60')
                p.locator('[name="player_2"]').fill(name)
                p.locator('button[type="submit"]').click()
                self.assertEqual(p.locator('[role="alert"]').inner_text(), poker.TEXTS[language]['setup_duplicate'])
                p.locator('[name="player_2"]').fill('Bob')
                p.locator('button[type="submit"]').click()
                self.assertEqual(poker.STATE['buyin_chips'], 1200)
                self.assertEqual(poker.total_pot(), 350)
                p.locator('a[href="/clock"]').click()
                before = p.locator('#timer').inner_text()
                p.clock.run_for(1000)
                self.assertNotEqual(p.locator('#timer').inner_text(), before)

    def test_cleared_second_percentage_then_two_paid_places_can_launch(self):
        p = self.page
        p.goto(self.url + '/setup')
        p.locator('[name="second_pct"]').fill('')
        p.locator('#paid_places').fill('2')
        p.locator('#paid_places').press('Tab')
        p.locator('button[type="submit"]').click()
        p.locator('a[href="/clock"]').click()
        self.assertTrue(p.locator('#timer').is_visible())

if __name__ == "__main__":
    unittest.main()

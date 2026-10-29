# Poker-international

Application Flask de gestion d'une soirée de poker en français, espagnol et anglais.
Caves en argent et jetons séparées, blinds, alarme, pause, kills, recaves, classement,
gains, feu d'artifice et export Excel.

## Lancer localement

Python 3.12 recommandé.

```sh
python -m venv .venv
# Activer .venv selon votre système
python -m pip install -r requirements.txt
flask --app app run
```

Le serveur de développement n'est pas destiné à une exposition publique.
Aucune procédure de déploiement n'est exécutée par les tests.

## Tests

```sh
python -m pip install -r requirements-dev.txt
python -m playwright install chromium
python -m unittest discover -s tests -v
node --test tests/clock.test.cjs
```

Node.js 24 est utilisé en CI ; aucun paquet npm n'est requis.
Tests Flask seuls : `python -m unittest discover -s tests -p test_app.py -v`.

Les tests navigateur démarrent un serveur sur 127.0.0.1 et un port libre.
La variable facultative `POKER_BROWSER` peut désigner un exécutable Chrome installé.
`POKER_SCREENSHOTS` peut désigner un dossier existant pour enregistrer les captures.
La CI incluse exécute les tests sous Linux/Chromium ; elle ne déploie pas.

## Comportement de l'horloge

Ouvrir Kill suspend temporairement l'horloge. Annuler ou valider reprend seulement
si elle tournait auparavant, en conservant le temps restant. Une pause volontaire
reste active. La dernière élimination arrête l'horloge et affiche le vainqueur.

Une requête échouée affiche une erreur ; la fenêtre Kill reste annulable après
l'échec ou le délai de 15 secondes. Ne pas répéter aveuglément une recave si la
connexion a été interrompue après son enregistrement : le serveur ne fournit pas
encore d'idempotence. Le changement de langue dans l'horloge ne recharge plus la page.

## Provenance et limites

Cette livraison est basée sur l'archive fournie par le propriétaire, dont le
commentaire GitHub indique `c997114851d6375417b3625339681fdb7a863146`.
La branche locale `fix/elimination-timer` part d'un instantané importé, pas d'un
historique GitHub cloné. Pour une PR, appliquer le patch dans un clone authentifié.

L'application utilise encore un état global non persistant partagé par les
visiteurs d'un processus. Ne pas considérer ce correctif comme une migration
multi-utilisateur ou une sécurisation complète. Voir [l'audit](docs/AUDIT.md).

## Configuration

Le prix d'une cave (argent) sert au calcul du pot et des gains, recaves comprises.
Les jetons reçus par cave (1200 par défaut, modifiable) servent aux blinds.
Une saisie incorrecte affiche un message dans le formulaire et conserve les
valeurs afin de permettre leur correction. Les pourcentages des places non
utilisées ne sont pas requis.

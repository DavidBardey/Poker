# Poker Salomon Web - Mode 3 amélioré

Nouveautés :

- un joueur qui a utilisé toutes ses recaves ne peut plus recaver ;
- dans Kill / Pause, le joueur killé n'apparaît plus comme killeur possible ;
- au clic sur Fin de partie + Excel, l'Excel est téléchargé puis l'application revient automatiquement au menu principal ;
- l'Excel contient le classement complet : vainqueur 1er, dernier éliminé 2e, etc.

Installation :

```bash
pip install flask openpyxl
python app.py
```

Puis ouvrir :

```text
http://127.0.0.1:5000
```

Place `dealer.png` et `Takefive.mp3` dans le dossier `static/`.


## Correction stricte du killeur

Dans la fenêtre Kill / Pause, la liste "Joueur qui a fait le kill" est maintenant recalculée depuis la liste visible des joueurs actifs.

Test attendu avec David, Regis, Elvire :
- si "Regis" est le joueur killé, les killeurs proposés doivent être "David" et "Elvire".
- "Regis" ne doit jamais apparaître dans sa propre liste de killeurs.

Sur la page d'horloge, vous devez voir la mention :
"Version corrigée : filtre strict du killeur".

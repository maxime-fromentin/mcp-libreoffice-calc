# MCP LibreOffice Calc

Serveur MCP pour piloter LibreOffice Calc en direct avec UNO.

Ce projet est une base de travail separee de `mcp-libre`. L'objectif est simple : avoir des outils fiables pour manipuler un classeur ouvert ou ouvert par le serveur MCP, sans devoir modifier le fichier `.xlsx` hors de LibreOffice.

## Outils disponibles

- `libreoffice_status` : verifier la connexion UNO et les classeurs accessibles.
- `open_calc` : ouvrir un fichier Calc/Excel via LibreOffice.
- `ensure_calc_document` : utiliser le classeur deja ouvert, ou l'ouvrir explicitement s'il manque.
- `list_calc_documents` : lister les classeurs Calc ouverts via UNO.
- `list_sheets` : lister les feuilles.
- `add_sheet` : ajouter une feuille.
- `remove_sheet` : supprimer une feuille.
- `read_cell` : lire une cellule.
- `write_cell` : ecrire une cellule en direct.
- `write_range` : ecrire un tableau 2D.
- `set_hyperlink` : ajouter un lien Ctrl+clic dans une cellule Calc.
- `link_url_range` : transformer les URLs texte d'une plage en liens Ctrl+clic.
- `apply_calc_template` : appliquer un template visuel, dont `pc_nostalgia`.
- `auto_fit_sheets` : ajuster largeur, hauteur et retour a la ligne.
- `save_document` : sauvegarder le classeur.

## Templates Calc

### `pc_nostalgia`

Template recommande pour les feuilles de suivi du datacenter :

- fond noir sur la zone visible ;
- texte vert terminal en police Consolas ;
- titres/en-tetes bleus ;
- quadrillage masque ;
- bordures nettoyees par defaut, donc pas de lignes verticales parasites.

Exemple :

```text
apply_calc_template(
  sheet="achat_Vendeur_spe",
  template="pc_nostalgia",
  used_range="A1:N45",
  title_rows=1,
  header_rows=[7]
)
```

Pour rendre les URLs cliquables avec Ctrl+clic :

```text
link_url_range(sheet="achat_Vendeur_spe", cell_range="K8:K20")
```

## Lancement manuel

```bash
cd /home/maxime/.local/share/mcp-libreoffice-calc
uv run python src/main.py
```

Le serveur demarre ou rejoint un listener LibreOffice UNO sur `127.0.0.1:2002`.

## Configuration Code

Exemple pour `~/.code/config.toml` :

```toml
[mcp_servers.libreoffice_calc]
command = "uv"
args = ["run", "python", "/home/maxime/.local/share/mcp-libreoffice-calc/src/main.py"]
cwd = "/home/maxime/.local/share/mcp-libreoffice-calc"

[mcp_servers.libreoffice_calc.env]
PYTHONPATH = "/home/maxime/.local/share/mcp-libreoffice-calc/src:/usr/lib/python3.14/site-packages"
```

## Notes

Pour piloter un classeur deja ouvert, LibreOffice doit etre accessible par UNO. Le chemin recommande est `ensure_calc_document` une seule fois, puis `add_sheet`, `write_cell` ou `write_range`. Les outils de lecture/ecriture ne rouvrent plus un fichier absent par surprise.

Les fichiers copies depuis le projet d'origine restent couverts par leur licence. Le serveur principal dans `src/` a ete reecrit pour cette base personnelle.

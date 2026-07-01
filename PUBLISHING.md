# Renommer et publier le MCP

Nom recommande pour GitHub : `mcp-libreoffice-calc`.

Pourquoi ce nom : il est clair, public, descriptif, et il evite `perso` qui fait projet prive. Il dit directement que le MCP pilote LibreOffice Calc.

## Renommage local

Depuis `/home/maxime/.local/share` :

```bash
mv mcp-libre-perso mcp-libreoffice-calc
cd mcp-libreoffice-calc
```

Modifier ensuite `pyproject.toml` :

```toml
[project]
name = "mcp-libreoffice-calc"
description = "MCP server for controlling LibreOffice Calc via UNO"

[project.scripts]
mcp-libreoffice-calc = "libremcp:main"
```

Modifier le titre du `README.md` :

```markdown
# MCP LibreOffice Calc
```

## Configuration Code

Dans `~/.code/config.toml`, remplacer l'ancien bloc par :

```toml
[mcp_servers.libreoffice_calc]
command = "/home/maxime/.local/bin/uv"
args = ["run", "python", "/home/maxime/.local/share/mcp-libreoffice-calc/src/main.py"]
cwd = "/home/maxime/.local/share/mcp-libreoffice-calc"

[mcp_servers.libreoffice_calc.env]
PYTHONPATH = "/home/maxime/.local/share/mcp-libreoffice-calc:/home/maxime/.local/share/mcp-libreoffice-calc/src:/usr/lib/python3.14/site-packages"
```

Puis redemarrer/reload le client MCP. Les outils MCP sont charges au demarrage du process, donc un serveur deja lance ne voit pas les nouveaux noms ou nouveaux tools.

## Initialiser GitHub

```bash
cd /home/maxime/.local/share/mcp-libreoffice-calc
git init
git add .
git commit -m "Initial LibreOffice Calc MCP server"
gh repo create mcp-libreoffice-calc --public --source=. --remote=origin --push
```

## Nettoyage avant publication

Avant de push :

```bash
find . -type d -name '__pycache__' -prune -exec rm -rf {} +
find . -type d -name '.venv' -prune -exec rm -rf {} +
```

Verifier que `.gitignore` exclut au minimum :

```gitignore
.venv/
__pycache__/
*.pyc
.env
```

## Outils publics importants

- `open_calc` : ouvrir un classeur Calc/Excel.
- `ensure_calc_document` : reutiliser un classeur ouvert ou l'ouvrir.
- `list_sheets` : lister les feuilles.
- `add_sheet` / `remove_sheet` : gerer les feuilles.
- `read_cell` / `write_cell` / `write_range` : lire/ecrire des cellules.
- `set_hyperlink` : ajouter un lien Ctrl+clic.
- `link_url_range` : transformer une plage d'URLs en liens Ctrl+clic.
- `apply_calc_template` : appliquer le template `pc_nostalgia`.
- `auto_fit_sheets` : ajuster largeur/hauteur.
- `save_document` : sauvegarder.

## Note importante

`pc_nostalgia` est un template visuel opinionated. Pour un projet public, le garder comme exemple/demo est bien, mais les outils doivent rester generiques afin que les utilisateurs puissent appliquer leurs propres styles.

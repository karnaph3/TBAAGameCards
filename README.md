# TBAA Game Cards

This tool converts the data in a CSV/XLSX table into a PDF of a tournament's game cards.

`source .venv/Scripts/activate`
`deactivate`

pyinstaller --onefile  --name TBAA_Card_Generator  --add-data "src/game_card_template.html;." --add-data "src/icon.ico;."  --icon=src/icon.ico  src/generate_card.py

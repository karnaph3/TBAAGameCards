# TBAA Game Cards

This tool converts the data in a CSV/XLSX table into a PDF of a tournament's game cards.

`source .venv/Scripts/activate`
`deactivate`

pyinstaller --onefile --icon=src/icon.ico --add-data "src/game_card_template.html;src" src/generate_card.py"

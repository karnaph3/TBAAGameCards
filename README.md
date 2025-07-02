# TBAA Game Cards

This tool converts the data in a CSV/XLSX table into a PDF of a tournament's game cards.

`source .venv/Scripts/activate`
`deactivate`

`pyinstaller --onefile --icon=src/icon.ico src/generate_card.py --name "Tournament Game Card Generator"`
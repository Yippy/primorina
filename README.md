# Teyvat Central Banking Systems
Using Google Sheets to save logs for Genshin Impact, such as:

- Primogems - earning and spending
- Genesis Crystals - earning and spending
- Resins - spending
- Mora earning

## Google Add-on
Coming soon

## Preview
<img src="https://raw.github.com/Yippy/primorina/master/images/teyvat_central_banking_systems_preview.png?sanitize=true">

## Template Document
If you prefer to use the Teyvat Central Banking Systems document with embedded script, you can make a copy here:
spreadsheet link: https://docs.google.com/spreadsheets/d/1VrNIbZGj2XhVZv7eRMgoSO3xUhOm0EVFGFqW0y1BLpI/edit#gid=2087370562

## How to compile script
This project uses https://github.com/google/clasp to help compile code to Google Script.

1. Run ```npm install -g @google/clasp```
2. Edit the file .clasp.json with your Google Script
3. Run ```clasp login```
4. Run ```clasp push -w```
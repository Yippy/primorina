# Teyvat Central Banking Systems
Using Google Sheets to save logs for Genshin Impact, such as:

- Primogems - earning and spending
- Genesis Crystals - earning and spending
- Resins - spending
- Mora earning

## Project Website
Visit the Genshin Impact collection of Google Sheets tools:

https://gensheets.co.uk

## Google Add-on
Teyvat Central Banking Systems is available on Google Workspace Marketplace from 19th November 2021.
https://workspace.google.com/marketplace/app/wish_tally/791037722195

## Preview
<img src="https://raw.github.com/Yippy/primorina/master/images/teyvat_central_banking_systems_preview.png?sanitize=true">

## Tutorial

[Change Language](docs/CHANGE_LANGUAGE.md)

[Get README](docs/GET_README.md)

[Use Auto Import from miHoYo](docs/USE_AUTO_IMPORT.md)

## Template Document
If you prefer to use the Teyvat Central Banking Systems document with embedded script, you can make a copy here:
spreadsheet link: https://docs.google.com/spreadsheets/d/1VrNIbZGj2XhVZv7eRMgoSO3xUhOm0EVFGFqW0y1BLpI/edit

## How to compile script
This project uses https://github.com/google/clasp to help compile code to Google Script.

1. Run ```npm install -g @google/clasp```
2. Edit the file .clasp.json with your Google Script
3. Run ```clasp login```
4. Run ```clasp push -w``` 
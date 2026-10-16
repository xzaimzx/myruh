# MYRUH Padelers Tour

Weekly padel rankings for the MYRUH Padelers Tour in Riyadh. Season 2026/2027.

- `index.html` is the website, served by GitHub Pages.
- `Code.gs` is the Google Apps Script that connects the website to the Google Sheet (tabs: Players, Stops, Matches, ScoreLog, Settings). It runs inside the sheet's Apps Script editor; the copy here is for reference.

Anyone can enter scores with their name and the weekly Score PIN, and every change is recorded in the ScoreLog tab. Organiser tools need the Admin PIN. Both PINs live in the sheet's Settings tab and are never sent to the website.

Ratings are Elo-style: everyone starts at 1000, K is 32 for a player's first 10 games and 20 after, and each game counts the win and the share of points won against the average strength of the opposing pair.

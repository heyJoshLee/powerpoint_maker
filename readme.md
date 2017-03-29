#Automatically Create Powerpoint Presentations from Google Docs + Image Scraping 


Create config.yaml file to add Google Sheet IDs. Ex
GOOGLE_SHEET_VOCABULARY: xxxxxxxx-xxxxxxxxxx-xxxxxxxxxxxxx-xxxxxxxx
GOOGLE_SHEET_QUIZ: xxxxxxxx-xxxxxxxxxx-xxxxxxxxxxxxx-xxxxxxxx


Create config.json file. Get this from Google Ex

{
  "client_id": "xxxxxxxxx.apps.googleusercontent.com",
  "client_secret": "xxxxxxxxx",
  "scope": [
    "https://www.googleapis.com/auth/drive",
    "https://spreadsheets.google.com/feeds/"
  ],
  "refresh_token": "x/xxxxxx-xxxx"
}

TODO:
* Add bundler support
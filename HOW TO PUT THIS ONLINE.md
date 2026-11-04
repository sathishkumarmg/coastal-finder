# Putting the Coastal Company Finder online (free, ~15 minutes)

You need: a free GitHub account and a free Streamlit account (sign in with GitHub).

1. Go to github.com -> New repository -> name it e.g. `coastal-finder`,
   set it to **Private** -> Create.
2. Click "uploading an existing file" and drag in ALL files from this folder,
   including the `.streamlit` folder (app.py, engine.py, coastal_companies.py,
   requirements.txt, .streamlit/config.toml). Commit.
   (Hidden folder tip: if `.streamlit` doesn't drag in, use "Add file > Create new file",
   name it `.streamlit/config.toml` and paste its contents.)
3. Go to share.streamlit.io -> Create app -> pick the repo, branch `main`,
   main file `app.py`.
4. Open "Advanced settings" -> Secrets, and paste:
       APP_PASSWORD = "choose-a-password"
5. Deploy. In a minute you get a link like https://coastal-finder.streamlit.app
6. Share the link + password with the user.

## Test it before sharing
Upload one month's file yourself. If the progress shows lookups working and
companies being found, you're good. If it stops with "site started blocking
requests", Zaubacorp is blocking the cloud server - see options below.

## If the cloud version gets blocked
- Run the same portal from your own PC (home/office internet is rarely blocked):
      pip install -r requirements.txt
      streamlit run app.py
  Anyone on the same office Wi-Fi can open the "Network URL" it prints.
- Or switch the lookup to a paid company-data API (only `lookup()` in
  coastal_companies.py needs changing).

## Good to know
- Free apps sleep after a few days unused; opening the link wakes them (~1 min).
- A full month takes about an hour. The browser tab must stay open; if it
  closes, uploading the same file again resumes from where it stopped.

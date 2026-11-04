# Listing Runner

Runs the Reverb and Etsy Excel bots for the **Auto Bot** page in SellerPilot Hub.
The server is the master; each laptop runs this one runner.

## Setup on a laptop
1. Copy this folder to the laptop, for example `C:\Listing Runner`.
2. On the server, make one token per laptop, with its stores and no owner, so every user's jobs for
   those stores run on it (Etsy and Reverb both):
   `python manage.py create_ingest_token --label worker-1 --scopes reverb etsy --stores IQ IQR`
   This is done once. The laptop takes the oldest waiting job for its stores; while it is busy,
   new jobs for those stores show as pending on the Auto Bot page.
3. Copy `ingest.env.example` to `ingest.env` and set `INGEST_TOKEN`, `REVERB_BOT_DIR` and `ETSY_BOT_DIR`.
4. `pip install requests`, then double click `start.bat`. Leave the window open.

## What it does
- Polls `GET <INGEST_BASE_URL>reverb/next-job/` and `etsy/next-job/`.
- Downloads the job's Excel to `jobs\<job id>\`.
- Runs the bot as `python <script> <excel> [--store X]` in its own console window.
- When the bot stops (finished, captcha timeout, error), uploads the newest
  `Created/Remaining` or `Updated/remain-to-update` files the bot wrote and marks the job complete.
- Every 5 seconds it checks the job. If it was cancelled or deleted on the page,
  it stops that bot's Python and the Chrome on the bot's debug port. The runner keeps running.

The bots, `stores.json`, Chrome profiles and VPN settings are not changed.

`python poller.py --once` handles one job and exits (for testing).

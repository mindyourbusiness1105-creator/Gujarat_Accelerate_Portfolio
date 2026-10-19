# Gujarat Accelerate Network

Public website (companies, buy and sell, jobs and job fairs) plus an admin page at /admin.
Admin uploads the Excel list, posts new items and triggers WhatsApp notifications.

## Run in VS Code
1. Install Python 3.10+ and open this folder in VS Code (File > Open Folder).
2. Terminal > New Terminal, then:
   - Windows: `python -m venv .venv` then `.venv\Scripts\activate`
   - Mac/Linux: `python3 -m venv .venv` then `source .venv/bin/activate`
3. `pip install -r requirements.txt`
4. Copy `.env.example` to `.env` and set SECRET_KEY, ADMIN_PASSWORD (leave the WhatsApp lines empty for now).
5. `python app.py`, then open http://localhost:5000 (public) and http://localhost:5000/admin.
6. In /admin: log in, upload your Excel file, post something. With no WhatsApp credentials it runs as a dry run (nothing is sent).

## Excel format
First sheet named `Companies` (or the first sheet). Header row required. Columns:
name, sector, city, state, region (North/South/East/West/Central), about, sells, buys, jobs, whatsapp, website,
and a column whose header contains "opt" (for example "WhatsApp opt-in (Y/N)") with Y or N.
Several items in one cell: separate with a semicolon. Delete example rows before uploading.

## Connect the WhatsApp Business API (Meta Cloud API)
1. Create a Meta Business account, then at developers.facebook.com create an app with the WhatsApp product.
2. Add your business phone number and note the Phone Number ID. Create a permanent access token (System User).
3. In WhatsApp Manager > Message templates, create a template named `network_update` (language English) with this body:
   `New on the Gujarat Accelerate network: {{1}} | {{2}} | See it and connect: {{3}}`
   Wait for approval (usually minutes to a day).
4. Put WA_TOKEN, WA_PHONE_ID, WA_TEMPLATE, SITE_URL (your real public address) in `.env`, restart.
5. In /admin use "Test WhatsApp" with your own number first. Then publish a post with the notify box ticked.
Notes: messages go to companies marked opt-in Y only. Meta limits how many new people you can message per day until your business is verified. Check Meta's current pricing and limits.

## Deploying on Render
1. Push this project to a GitHub repository and connect that repository to Render.
   Do not commit `.env` or `data.db`; both are excluded by `.gitignore`.
2. In Render, create a new Blueprint instance from the repository. Render reads
   `render.yaml` and generates `SECRET_KEY`; enter a strong, unique `ADMIN_PASSWORD`
   when prompted.
3. After the service deploys, open its `onrender.com` URL to share the public site.
   The admin page is at `/admin`; keep its password private.
4. The repository intentionally does not contain the local database. The new
   service starts with an empty database; log in at `/admin` and import the
   companies Excel file to populate it.

The included Render plan is free and uses ephemeral storage. Database changes can
be lost when the service restarts or is redeployed. For a persistent production
database, attach a paid Render disk, set `DATABASE_PATH` to a file on its mount
path, and keep regular backups before importing or editing live data.

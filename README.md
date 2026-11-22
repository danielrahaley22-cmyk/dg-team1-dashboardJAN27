# dg-team1-dashboardJAN27

Bob's Team 1 dashboard for DG XXL January 2027 (Melbourne), with weekly outreach ticks.

- `index.html` – the dashboard (served by GitHub Pages). Shows code, name and Instagram only, plus Melbourne sales.
- `apps-script/Code.gs` – Google Apps Script attached to the "DG Team 1 Outreach (Bob)" sheet. It sends the page the team list and sales, saves each Focus / Sent / Replied tick and comment to the Outreach tab, and emails a weekly summary every Sunday at 8pm.

To connect: in the sheet, Extensions > Apps Script, paste `Code.gs`, run `setup` once, then deploy as a web app (Execute as: Me, Access: Anyone) and paste its `/exec` URL into `SCRIPT_URL` in `index.html`.

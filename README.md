# EAA Final Report Excel Scraper

Python script for EAA chapter coordinators. It logs into [eaachapters.org](https://www.eaachapters.org) as **whoever runs it**, walks Current / Past Events in a date range you type in, skips cancelled events and events that have not been submitted to headquarters, downloads each event’s Final Report tables, and writes one Excel workbook. Credentials are prompted and never saved.

The workbook includes a **Child Flights** sheet: one row per youth who flew in the range, with how many flights they took in that window.

Young Eagles reports include names, emails, and addresses. Keep the `.xlsx` file private. Do not commit it to git.

## Setup (once per computer)

Python 3.10+ is required.

```bash
cd "EAA Scraper"
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium
```

On Windows, activate with `.venv\Scripts\activate`.

Chromium is only used to complete the site’s login (including reCAPTCHA). Report data is then loaded through the same API the website uses.

## Run

```bash
source .venv/bin/activate
python scrape_eaa_flights.py
```

You will be asked for:

1. Email
2. Password (hidden)
3. Start date
4. End date

Dates can be `YYYY-MM-DD` or other common formats. The range is inclusive. Cancelled rallies and rallies not submitted to headquarters are omitted.

The file is written in the `Reports` folder next to the script:

`Reports/eaa_final_reports_YYYY-MM-DD_to_YYYY-MM-DD.xlsx`

Sheets: Events, Headquarters, Age, Pilot Flights, Parents Registered / Flown / Waitlisted / Not Flown, Volunteers, Child Flights.

If headless login fails, the script opens a visible browser and retries.

## Notes

- Each user should run the script with **their own** EAA Chapters login.
- A short pause sits between API calls so the site is not hammered.
- Child flight counts treat each in-range event a youth appears on as flown as one flight, unless the report row includes a higher flight count.

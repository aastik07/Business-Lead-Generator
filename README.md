# Local Business Lead Generator

A Python automation tool that collects publicly displayed business information from Google Maps and exports it to a clean Excel file — saving hours of manual copy-pasting.

---

## About

Salespeople, marketing agencies, and local businesses often need lists of competitors or potential leads in a specific city. Doing this manually on Google Maps is slow and repetitive.

This tool automates that process. You enter a city, a business type, and how many leads you want. The program opens Google Maps, scrolls through the results automatically, collects the publicly displayed information for each business, removes duplicates, and saves everything into a clean Excel file.

---

## Features

- City-based search
- Business niche search (e.g., Interior Designers, Restaurants, Dentists)
- Selenium browser automation
- Automatic scrolling through results
- Business information extraction
- Duplicate removal (by Google Place ID → Name + Phone → Name + Address)
- Excel export with styled headers
- Automatic filename generation with date
- Graceful error handling

---

## How It Works

```
User enters city + niche + max leads
           ↓
      Chrome opens
           ↓
   Google Maps search
           ↓
  Results are collected
           ↓
   Duplicates removed
           ↓
    Excel generated
```

---

## Technologies

- **Python 3**
- **Selenium** — browser automation
- **Google Chrome** + ChromeDriver
- **openpyxl** — Excel file creation
- **VS Code** — recommended editor

---

## Installation

**Step 1 — Create a virtual environment**
```
python -m venv .venv
```

**Step 2 — Activate it**
```
.venv\Scripts\activate
```

**Step 3 — Install dependencies**
```
pip install -r requirements.txt
```

> **Note:** Make sure Google Chrome is installed. Selenium uses Chrome's built-in driver manager automatically (Selenium 4.6+).

---

## Run

```
python main.py
```

**Example session:**

```
========================================
   LOCAL BUSINESS LEAD GENERATOR
========================================

Enter city: Chandigarh
Enter business niche: Interior Designers
Enter maximum leads: 100

Starting Chrome...

Searching Google Maps for:
Interior Designers in Chandigarh

Collecting businesses...

  [1] ABC Interiors
  [2] XYZ Interior Design
  [3] Modern Spaces
  Duplicate skipped.
  [4] Design Studio
  ...

Completed!
Total leads collected: 87

Excel saved to:
output/interior_designers_chandigarh_2026-09-08.xlsx
```

---

## Excel Output

Each Excel file is saved inside the `output/` folder with a descriptive filename:

```
output/interior_designers_chandigarh_2026-09-08.xlsx
```

**Columns:**

| Column | Description |
|---|---|
| Business Name | The name of the business |
| Rating | Star rating (e.g. 4.5) |
| Reviews | Number of reviews (e.g. 1234) |
| Phone | Publicly listed phone number |
| Address | Publicly listed address |
| Plus Code | Google's location code |
| City | The city you searched |
| Business Niche | The niche you searched |
| Google Maps URL | Clickable hyperlink directly to the listing |

The Excel file includes:
- Bold, dark-blue header row
- Frozen first row
- Auto-filters on all columns
- Auto-adjusted column widths
- Clickable hyperlinks in the Google Maps URL column

---

## Project Use Case

This tool can help:

- **Sales teams** build prospect lists faster
- **Marketing agencies** research local businesses for outreach
- **Local businesses** analyse competitors in their area
- **Lead-generation businesses** automate a time-consuming manual task
- **Market researchers** gather data for reports

---

## Limitations

- Google Maps' HTML structure can change — selectors may need updates if that happens.
- Some businesses may not have phone/address publicly displayed; those fields will show `N/A`.
- Results are not guaranteed to be exhaustive.
- CAPTCHA or temporary rate-limiting by Google may occur; the program will detect this and stop gracefully.
- The project does **not** bypass any anti-bot protections.

---

## Responsible Use

This tool collects only publicly displayed information — the same information visible to any person browsing Google Maps. Use the collected data responsibly and in accordance with Google's Terms of Service and applicable privacy laws.

---

## Future Improvements

Ideas for extending the project (not currently implemented):

- GUI (Tkinter or web-based)
- CSV export option
- Website URL extraction
- Lead scoring
- CRM integration (e.g., HubSpot, Notion)

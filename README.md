# Yandex Maps Parser

Business data scraper for Yandex Maps. Extracts leads with contact enrichment — phones, emails, social media links — and exports to Excel.

## Two versions

**`Yandex_parser_leads.ipynb`** — Base version  
Scroll + collect + export. Simple, fast, reliable.

**`YaMa_V9_FINAL.ipynb`** — Full version (recommended)  
Complete lead generation pipeline with multi-source enrichment.

## What V9 collects

| Field | Source |
|---|---|
| Name, address, phone | Yandex Maps |
| Rating, review count | Yandex Maps |
| Website | Yandex Maps → Yandex Search fallback |
| Email | mailto links → contacts page → page text → OCR |
| Instagram, VK, Telegram, WhatsApp | Yandex Maps + website |
| Average bill | Yandex Maps |

## How it works

```
1. Search Yandex Maps by location + category
2. Auto-scroll to collect all listings up to target count
3. Visit each listing → extract structured data
4. For each website: find email via 4-priority chain
5. Checkpoint saves every N records (resume on failure)
6. Export to Excel with statistics summary
```

## Setup

```bash
pip install selenium pandas beautifulsoup4 requests pillow pytesseract
```

Requires Chrome + ChromeDriver. For Google Colab — Chrome install is handled in the notebook.

## Usage

Set three variables at the top of the notebook:

```python
location = 'москва выхино'   # city or district
title    = 'риэлтор'         # business category
count_of_units = 1000        # max results
```

Run all cells. Excel file downloads automatically (Colab) or saves to working directory.

## Output example

```
============================================================
  ✅ ГОТОВО!
============================================================
  📊 Собрано записей: 847
  📞 С телефоном:     761
  🌐 С сайтом:        534
  📧 С email:         312
  📱 Instagram:       198
  📱 Telegram:        143
============================================================
```

## Stack

Python · Selenium · BeautifulSoup4 · Pandas · Requests · Pytesseract

---

<img width="667" height="447" alt="image" src="https://github.com/user-attachments/assets/ec8fd1a2-cfab-4de0-8e7f-f8ede8371efb" />

*Part of [DenisZDV](https://github.com/DenisZDV) automation toolkit*

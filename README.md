<div align="center">

# Frido Carrier Efficiency Console

**One dashboard for shipment performance, carrier SLAs and carrier billing**

![Python](https://img.shields.io/badge/Python-3.14-3776AB?logo=python&logoColor=white)
![Django](https://img.shields.io/badge/Django-5.x-092E20?logo=django&logoColor=white)
![BigQuery](https://img.shields.io/badge/Google%20BigQuery-data-4285F4?logo=googlebigquery&logoColor=white)
![Vercel](https://img.shields.io/badge/Deploy-Vercel-000000?logo=vercel&logoColor=white)

</div>

## What it does

An internal web app for the Frido operations and finance teams. It reads shipment data, order data and carrier invoices then shows how each courier performs, how much it costs and where billing errors leak money.

It answers three questions.

1. How many orders were placed and how many shipped and delivered
2. Which carrier is fast, which is slow and which breaks its promised time
3. Are carriers billing us correctly and where can we recover money

Load a date window then filter by channel, carrier, warehouse, payment, city tier, zone or weight. Everything updates at once.

## How it works

```
BigQuery  ClickPost shipments   ›  KPIs, TAT, SLA, RTO, carrier score
BigQuery  Unicommerce orders    ›  orders per day, order to ship time
Invoices  Google Drive or upload ›  cost, weight disputes, reconciliation

              feed into

     Django backend  ›  Single page dashboard
```

The backend pulls the data, caches it and rebuilds the report on every filter. The browser only talks to the Django API, never to BigQuery.

## Tech stack

| Part | Tool |
| - | - |
| Backend | Django 5, Python 3.14 |
| Data | Google BigQuery with pyarrow reads |
| Invoices in | Google Drive API, openpyxl, pyxlsb |
| Frontend | Vanilla JavaScript, Chart.js, custom CSS |
| AI summary | OpenAI API |
| Hosting | Vercel, Railway or Docker with Gunicorn |

## Data sources

| Source | Gives | Note |
| - | - | - |
| ClickPost shipments | Carrier, TAT, delivery, RTO, cost | Complete only after an order ships |
| Unicommerce orders | Orders placed, order time | Complete at the moment an order is placed |
| Carrier invoices | Billed amount, weight, charges | One parser per carrier format |
| Item master | SKU expected weight | Used to catch weight over charges |

## What it measures

| Feature | Logic |
| - | - |
| Carrier score | Volume, delivery rate, TAT, RTO and cost blended into one score per carrier |
| SLA compliance | Actual delivery time against the promised time, shown as in TAT and out of TAT |
| Orders per day | Order count from Unicommerce so recent days do not show a false dip |
| Cost analysis | Invoice spend by carrier, category and SKU with cost per parcel and charge per kg |
| Weight disputes | Charged weight against item master weight, flags the excess and the over charge |
| Reconciliation | Billed amount with editable disputes, credit notes and TDS, live payable |
| AI summary | Short written overview anchored on a fixed health score |

## Business impact

| Before | After | Result |
| - | - | - |
| Recent orders looked like a drop | Order count from Unicommerce | True daily orders, no false dip |
| Carrier data spread across sheets | One scored view per carrier | Faster and fairer carrier choice |
| Weight over charges missed | Charged weight checked against master | Money recovered from carriers |
| Invoice checks done by hand | Every carrier format auto parsed | Finance hours saved each month |
| SLA breaches found late | Live in TAT and out of TAT | Early action on slow lanes |

In short it shortens the path from data to decision. Orders are counted right, carriers are held to their promises and billing leaks are caught.

## Project structure

```
carrier_dashboard
├── config              settings and urls
├── dashboard
│   ├── bq.py           BigQuery reads
│   ├── kpi.py          filters, KPIs, scoring
│   ├── invoices.py     invoice parsers and cost analysis
│   ├── gdrive.py       Google Drive import
│   ├── tat.py          TAT and SLA rules
│   ├── views.py        API endpoints and cache
│   └── templates/dashboard  index.html and login.html
├── api                 serverless entry point
└── requirements.txt
```

## Setup

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python manage.py runserver
```

Then sign in and load a window from BigQuery or upload an invoice.

Configure with environment variables in a `.env` file.

| Variable | Meaning |
| - | - |
| `BQ_PROJECT` `BQ_DATASET` `BQ_TABLE` | ClickPost shipment table location |
| `BQ_UC_TABLE` | Unicommerce orders table |
| `BQ_PARTITION_COLUMN` | Partition column for fast reads |
| `GOOGLE_SA_CLIENT_EMAIL` `GOOGLE_SA_PRIVATE_KEY` | Service account for BigQuery and Drive |
| `GDRIVE_INVOICE_FOLDER` | Drive folder that holds invoices |
| `OPENAI_API_KEY` | Key for the AI summary |

## Deployment

| Target | Guide |
| - | - |
| Vercel | `VERCEL_DEPLOY.md` |
| Railway | `RAILWAY.md` |
| Docker | `DEPLOY.md` |

Tip. Partition the ClickPost table on its date column so each load scans only the selected window instead of the full table. This is the biggest speed and cost win.

<div align="center">

Internal use only. Built for the Frido operations team.

</div>

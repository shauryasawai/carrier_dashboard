# Carrier Efficiency Console

A Django web app that ingests a shipment export (`.xlsx`) and renders a
carrier-partner performance dashboard: per-carrier KPIs and a weighted
efficiency score, plus business-mix breakdowns. The Excel format is fixed —
drop in the full dataset and the dashboard recomputes.

## Quick start

```bash
cd carrier_dashboard
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python manage.py runserver
```

Open http://127.0.0.1:8000/ and upload your workbook. No database or
migrations are needed — uploads are parsed in memory.

## Expected columns

Matching is case- and space-insensitive. Required: `Carrier Partner Name`.
Used when present: `Carrier Partner Account Name`, `Shipment Weight`,
`Payment Mode`, `Pickup Pincode`, `Drop Pin Code`, `Pickup Timestamp`,
`Delivery Timestamp`, `OFD1 Timestamp`, `Zone`, `Delivery Type`,
`Latest Status`, `Number of Delivery Attempts`.

Timestamps may be real datetimes or raw Excel serial numbers — both work.

## KPIs (per carrier)

| KPI | Definition |
|-----|------------|
| Pickup → OFD1 TAT | Avg hours, Pickup → OFD1, over rows having both |
| Pickup → Delivery TAT | Avg hours, Pickup → Delivery, over rows having both |
| Delivery success rate | Delivered ÷ Picked (Picked = has a pickup timestamp) |
| First-attempt strike rate | Delivered in 1 attempt ÷ Delivered |
| Average delivery attempts | Mean of the attempts column |

## Efficiency score (0–100)

Each KPI is min-max normalized across the carriers in the current filter set
(TATs and attempts inverted, since lower is better), then combined:

- 30% Pickup → OFD1 TAT
- 25% Pickup → Delivery TAT
- 20% Delivery success rate
- 15% First-attempt strike rate
- 10% Average delivery attempts

The score is **relative** to the carriers shown — useful for ranking within a
filter slice, not as a fixed cross-file benchmark. Weights live in
`dashboard/kpi.py` (`WEIGHTS`); edit there to retune. To switch to absolute
target-based scoring, replace `_score_metric` with a threshold function.

Default filter is **Forward** — reverse-pickup carriers rarely register as
"Delivered", so scoring them on forward-delivery KPIs is misleading. Use the
Delivery type toggle to rank reverse carriers separately.

## Warehouse breakdown

The same KPI block and efficiency score are also computed per **originating
warehouse**, keyed on `Pickup Pincode`, and shown in its own panel ranked by
shipment volume. Because exports often have a long tail of pincodes with only a
handful of shipments, warehouses below a volume threshold
(`WAREHOUSE_MIN_N`, default 20 — in `dashboard/kpi.py`) still appear in the
table with their KPIs but are not assigned a score, and are excluded from the
score normalization pool so they don't distort the ranking. Pincodes that
arrive as ints or floats (e.g. `560037.0`) are normalized to clean strings.

## Layout

```
carrier_dashboard/
├── manage.py
├── requirements.txt
├── config/            # project settings, urls, wsgi
└── dashboard/
    ├── kpi.py         # all parsing + KPI + scoring logic
    ├── views.py       # index + upload/refilter endpoint
    ├── templates/dashboard/index.html
    └── static/dashboard/{styles.css, app.js}
```

## Notes for production

`DEBUG = True` and a placeholder `SECRET_KEY` are set for local use. Before
deploying, set `DEBUG = False`, supply a real secret via environment variable,
restrict `ALLOWED_HOSTS`, and serve static files via WhiteNoise or a CDN. The
in-memory upload cache (`_CACHE` in `views.py`) is single-process; for multi-
worker deploys, swap it for a session/cache backend or re-upload per request.

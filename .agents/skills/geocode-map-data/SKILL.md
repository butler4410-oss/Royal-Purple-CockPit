---
name: geocode-map-data
description: Re-geocode customer and distributor map data for high-precision coordinates. Use when the user reports inaccurate map pin locations, asks to update/refresh map coordinates, or after importing new customer data that needs geocoding.
---

# Geocode Map Data

Re-geocode addresses in `customers.json` and `distributors.json` using the US Census Bureau's free batch geocoder to get rooftop-level coordinate precision (6 decimal places, sub-meter accuracy).

## When to Run

- User reports pins are in the wrong location on the map
- After importing new customer data (CSV upload or manual additions)
- Coordinates have low precision (3-4 decimal places = 100m–1km off)
- Multiple pins stacked at the same spot (geocoded to zip/city centroids)

## How It Works

1. Reads `customers.json` and `distributors.json` from the workspace root
2. Filters to US-only entries (`country == "US"` or no country field)
3. Sends addresses in batches of 500 to the Census Bureau batch geocoder
4. Parses matched results and updates lat/lng to 6 decimal places
5. Keeps original coordinates for any addresses the geocoder can't match
6. Saves updated JSON files back to disk

## Running the Geocoder

Execute the script directly:

```bash
python3 .agents/skills/geocode-map-data/geocode.py
```

The script outputs per-batch match counts and a final summary.

## Census Bureau Batch Geocoder

- **Endpoint:** `https://geocoding.geo.census.gov/geocoder/locations/addressbatch`
- **Method:** POST with multipart CSV upload
- **Batch limit:** 10,000 addresses per request (we use 500 for reliability)
- **Rate limit:** 1 second between batches
- **Coverage:** US addresses only (Canada/International entries are skipped)
- **No API key required**

### Input CSV format

```
id,street_address,city,state,zip
0,10501 MONROE RD,Matthews,NC,28105
```

### Response CSV format

```
"0","10501 MONROE RD, Matthews, NC, 28105","Match","Exact","10501 MONROE RD, MATTHEWS, NC, 28105","-80.733974,35.130281","638684712","R"
```

Fields: id, input_address, match_status, match_quality, matched_address, **lon,lat** (note: longitude first), tiger_id, side

## Data Files

- `customers.json` — ~4,700 entries with fields: `store_name`, `address`, `city`, `state`, `zip`, `country`, `latitude`, `longitude`, `type`, `county`
- `distributors.json` — ~58 entries with the same field structure

## Precision Reference

| Decimal Places | Accuracy   | Example          |
|----------------|------------|------------------|
| 2              | ~1.1 km    | 35.11            |
| 3              | ~111 m     | 35.115           |
| 4              | ~11 m      | 35.1149          |
| 6              | ~0.11 m    | 35.130282        |

## After Running

- Restart the app workflow so the map loads updated coordinates
- The map uses `@st.cache_data` on `load_customers()` — a restart clears the cache
- Non-US entries (Canada, International) retain their original coordinates

## Troubleshooting

- **Timeout errors:** The script retries each batch 3 times with 5-second delays
- **Low match rate:** Addresses with P.O. boxes, suite numbers, or non-standard formats may not match — they keep original coordinates
- **No matches at all:** Check internet connectivity and that the Census Bureau API is available

#!/usr/bin/env python3
"""
Re-geocode customer and distributor addresses using the US Census Bureau
batch geocoder for high-precision (6 decimal place) coordinates.

Usage:
    python3 .agents/skills/geocode-map-data/geocode.py

Reads customers.json and distributors.json from the workspace root,
geocodes US addresses, and saves updated coordinates back.
"""

import requests
import io
import csv
import json
import time
import sys
import os

WORKSPACE = os.path.dirname(os.path.abspath(__file__))
while not os.path.exists(os.path.join(WORKSPACE, "customers.json")):
    WORKSPACE = os.path.dirname(WORKSPACE)
    if WORKSPACE == "/":
        WORKSPACE = "/home/runner/workspace"
        break

CUSTOMERS_PATH = os.path.join(WORKSPACE, "customers.json")
DISTRIBUTORS_PATH = os.path.join(WORKSPACE, "distributors.json")
CENSUS_URL = "https://geocoding.geo.census.gov/geocoder/locations/addressbatch"
BATCH_SIZE = 500
PRECISION = 6


def geocode_batch(entries, batch_size=BATCH_SIZE):
    results = {}
    total = len(entries)

    for start in range(0, total, batch_size):
        batch = entries[start:start + batch_size]
        buf = io.StringIO()
        writer = csv.writer(buf)
        for i, c in enumerate(batch):
            idx = start + i
            writer.writerow([
                idx,
                c.get("address", ""),
                c.get("city", ""),
                c.get("state", ""),
                c.get("zip", ""),
            ])

        payload = {"benchmark": "Public_AR_Current", "vintage": "Current_Current"}
        files = {"addressFile": ("addresses.csv", buf.getvalue(), "text/csv")}

        for attempt in range(3):
            try:
                resp = requests.post(CENSUS_URL, data=payload, files=files, timeout=120)
                if resp.status_code == 200:
                    break
            except Exception as e:
                print(f"  Attempt {attempt + 1} failed: {e}", file=sys.stderr)
                time.sleep(5)
        else:
            print(f"  FAILED batch {start}-{start + len(batch)}", file=sys.stderr)
            continue

        reader = csv.reader(io.StringIO(resp.text))
        for row in reader:
            if len(row) >= 6 and row[2] == "Match":
                idx = int(row[0])
                coords = row[5].strip('"').split(",")
                if len(coords) == 2:
                    try:
                        lng = float(coords[0])
                        lat = float(coords[1])
                        results[idx] = (lat, lng)
                    except ValueError:
                        pass

        matched = sum(1 for i in range(start, start + len(batch)) if i in results)
        print(f"  Batch {start}-{start + len(batch)}: {matched}/{len(batch)} matched")
        time.sleep(1)

    return results


def process_file(filepath, label):
    print(f"\nLoading {label} from {os.path.basename(filepath)}...")
    with open(filepath) as f:
        data = json.load(f)

    us_entries = [
        (i, c) for i, c in enumerate(data)
        if c.get("country", "US") == "US"
    ]
    print(f"Geocoding {len(us_entries)} US {label}...")

    geocoded = geocode_batch([c for _, c in us_entries])

    updated = 0
    for batch_idx, (orig_idx, _) in enumerate(us_entries):
        if batch_idx in geocoded:
            lat, lng = geocoded[batch_idx]
            data[orig_idx]["latitude"] = round(lat, PRECISION)
            data[orig_idx]["longitude"] = round(lng, PRECISION)
            updated += 1

    print(f"Updated {updated}/{len(us_entries)} US {label} coordinates")

    with open(filepath, "w") as f:
        json.dump(data, f, indent=2)

    return updated, len(us_entries)


def main():
    print("=" * 50)
    print("Census Bureau Batch Geocoder")
    print("=" * 50)

    total_updated = 0
    total_entries = 0

    if os.path.exists(CUSTOMERS_PATH):
        u, t = process_file(CUSTOMERS_PATH, "customers")
        total_updated += u
        total_entries += t
    else:
        print(f"customers.json not found at {CUSTOMERS_PATH}")

    if os.path.exists(DISTRIBUTORS_PATH):
        u, t = process_file(DISTRIBUTORS_PATH, "distributors")
        total_updated += u
        total_entries += t
    else:
        print(f"distributors.json not found at {DISTRIBUTORS_PATH}")

    print(f"\nDone! Updated {total_updated}/{total_entries} total coordinates")
    print("Restart the app workflow to load the updated data.")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
Quick fix script to repair AOI tables with missing years
"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'backend'))

import pandas as pd
from storage_service import storage_service

print("=== Fixing AOI Tables with Missing Years ===\n")

# Read the corrupted CSV
aoi_data = storage_service.read_csv('config/aoi-tables.csv')

if aoi_data is None or aoi_data.empty:
    print("No AOI tables found")
    exit(0)

print(f"Found {len(aoi_data)} AOI tables")

# Show tables with missing years
null_years = aoi_data[aoi_data['tahun'].isna() | (aoi_data['tahun'] == '')]
print(f"\nTables with missing years: {len(null_years)}")

if len(null_years) > 0:
    print("\nWhich year should these tables be assigned to?")
    print("Options: 2024, 2055, or type 'skip' to skip")
    year_input = input("Enter year: ").strip()

    if year_input.lower() != 'skip':
        try:
            year = int(year_input)
            # Update missing years
            aoi_data.loc[aoi_data['tahun'].isna() | (aoi_data['tahun'] == ''), 'tahun'] = year

            # Save back
            success = storage_service.write_csv(aoi_data, 'config/aoi-tables.csv')

            if success:
                print(f"\n✅ Successfully updated {len(null_years)} tables to year {year}")
            else:
                print("\n❌ Failed to save changes")
        except ValueError:
            print("\n❌ Invalid year")
else:
    print("\n✅ All tables have valid years!")

print("\nDone!")

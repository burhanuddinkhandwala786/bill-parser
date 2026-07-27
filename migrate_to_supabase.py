import json
import os
import glob
import pandas as pd
from sqlalchemy import create_engine, text
import streamlit as st

# Load connection string from st.secrets or environment
DB_URL = st.secrets["SUPABASE_DB_URL"]
engine = create_engine(DB_URL)

def migrate():
    with engine.begin() as conn:
        print("Starting migration...")
        
        # 1. Find all stores from existing master CSV files
        # Expecting files like 'master_catalog_<store_slug>.csv' or similar existing structure
        csv_files = glob.glob("master_catalog_*.csv")
        
        for file in csv_files:
            store_slug = file.replace("master_catalog_", "").replace(".csv", "")
            display_name = store_slug.replace("_", " ").title()
            
            # Insert store record
            result = conn.execute(
                text("INSERT INTO stores (slug, display_name) VALUES (:slug, :name) ON CONFLICT (slug) DO UPDATE SET display_name = EXCLUDED.display_name RETURNING id;"),
                {"slug": store_slug, "name": display_name}
            )
            store_id = result.fetchone()[0]
            print(f"Migrating store: {store_slug} (ID: {store_id})")
            
            # Migrate Master Catalog CSV
            if os.path.exists(file):
                df = pd.read_csv(file)
                # Map CSV columns to table columns
                for _, row in df.iterrows():
                    conn.execute(
                        text("""
                            INSERT INTO master_skus (store_id, official_sku_name, category, default_unit, gst_rate, selling_price)
                            VALUES (:store_id, :sku, :cat, :unit, :gst, :price)
                        """),
                        {
                            "store_id": store_id,
                            "sku": str(row.get("official_sku_name", row.get("SKU_Name", ""))),
                            "cat": str(row.get("category", "General")),
                            "unit": str(row.get("default_unit", "PCS")),
                            "gst": float(row.get("gst_rate", 18.0)),
                            "price": float(row.get("selling_price", 0.0))
                        }
                    )
                print(f"  -> Migrated {len(df)} SKUs")

            # Migrate Vendor Memory JSON
            memory_file = f"vendor_memory_{store_slug}.json"
            if os.path.exists(memory_file):
                with open(memory_file, "r") as f:
                    memory_data = json.load(f)
                    for raw_name, mapped_sku in memory_data.items():
                        conn.execute(
                            text("""
                                INSERT INTO vendor_mappings (store_id, raw_name, mapped_sku)
                                VALUES (:store_id, :raw, :mapped)
                                ON CONFLICT (store_id, raw_name) DO UPDATE SET mapped_sku = EXCLUDED.mapped_sku
                            """),
                            {"store_id": store_id, "raw": raw_name, "mapped": mapped_sku}
                        )
                print(f"  -> Migrated vendor memory from {memory_file}")

    print("Migration complete! Check Supabase Table Editor to verify.")

if __name__ == "__main__":
    migrate()

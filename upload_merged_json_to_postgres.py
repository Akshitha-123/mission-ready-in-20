#!/usr/bin/env python3
import os
import json
import psycopg2
from psycopg2.extras import Json

# --- CONFIGURATION ---
DB_NAME = os.getenv("DB_NAME", "mrit_db")
DB_USER = os.getenv("DB_USER")
DB_PASSWORD = os.getenv("DB_PASSWORD")
DB_HOST = os.getenv("DB_HOST", "localhost")
DB_PORT = os.getenv("DB_PORT", "5432")

MERGED_DIR = 'MERGED_CONOPS_DRAWS'
TABLE_NAME = 'merged_conops_draws'


# --- CREATE TABLE IF NOT EXISTS ---
CREATE_TABLE_SQL = f'''
CREATE TABLE IF NOT EXISTS {TABLE_NAME} (
    id SERIAL PRIMARY KEY,
    source_directory_id TEXT NOT NULL,
    merged_data JSONB NOT NULL
);
'''

# --- MAIN SCRIPT ---
def main():
    conn = psycopg2.connect(
        dbname=DB_NAME,
        user=DB_USER,
        password=DB_PASSWORD,
        host=DB_HOST,
        port=DB_PORT
    )
    cur = conn.cursor()
    cur.execute(CREATE_TABLE_SQL)
    conn.commit()

    for filename in os.listdir(MERGED_DIR):
        if filename.endswith('.json'):
            file_path = os.path.join(MERGED_DIR, filename)
            with open(file_path, 'r') as f:
                data = json.load(f)
            # Extract source_directory_id from filename (e.g., '0001-merged.json')
            source_directory_id = filename.split('-')[0]
            cur.execute(
                f"INSERT INTO {TABLE_NAME} (source_directory_id, merged_data) VALUES (%s, %s)",
                (source_directory_id, Json(data))
            )
    conn.commit()
    cur.close()
    conn.close()
    print('Upload complete.')

if __name__ == '__main__':
    main()

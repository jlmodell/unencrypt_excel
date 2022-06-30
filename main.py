from datetime import time
from json import load
import msoffcrypto
import os
import sys
from shutil import copyfile
from dotenv import load_dotenv
import io
import pandas as pd
import sqlite3

load_dotenv()

KEY = os.environ.get("KEY", "")
assert KEY != "", "KEY environment variable is not set"

file = sys.argv[1]
assert os.path.isfile(file), "File does not exist"

decrypted = io.BytesIO()

sql_db = "C:\\temp\\schedule.db"


def copy_file(src):
    global file

    dst = src.replace(os.path.basename(src), "_" + os.path.basename(src))

    copyfile(src, dst)

    print("Copied file: " + src + " to " + dst)

    file = dst

    return dst


def decrypt_file(src):
    print("Decrypting file: " + src)

    with open(src, "rb") as f:
        msoffcrypto.OfficeFile(f).load_key(KEY).decrypt(decrypted)


def create_df():
    try:
        df = pd.read_excel(decrypted, sheet_name='Schedule',
                           header=[1])
    except:
        df = pd.read_excel(file, sheet_name='Schedule',
                           header=[1])

    print("Created dataframe")

    df = df.loc[(df['Ready'].str.contains('Y')) |
                (df['Ready'].str.contains('y')) |
                (df['WC Ready'].str.contains('Y')) |
                (df['WC Ready'].str.contains('y'))].copy()

    df.columns = [
        "requested",
        "wh_issue_date",
        "pulled",
        "posted",
        "racks",
        "parts_prep",
        "ready",
        "wc_ready",
        "job_done",
        "request",
        "in_parts_prep_by",
        "l",
        "run_date_time",
        "n",
        "item",
        "wc",
        "tooling",
        "r",
        "description",
        "lot",
        "lot_info",
        "qty",
        "comments",
        "x",
        "mp",
        "pallets",
    ]

    df.fillna("", inplace=True)

    df['wh_issue_date'] = pd.to_datetime(
        df['wh_issue_date'], errors='coerce').dt.strftime('%Y-%m-%d')
    df['request'] = pd.to_datetime(
        df['request'], errors='coerce').dt.strftime('%Y-%m-%d')
    df['in_parts_prep_by'] = pd.to_datetime(
        df['in_parts_prep_by'], errors='coerce').dt.strftime('%Y-%m-%d')
    df['run_date_time'] = pd.to_datetime(
        df['run_date_time'], errors='coerce').dt.strftime('%Y-%m-%d')
    df['x'] = pd.to_datetime(
        df['x'], errors='coerce').dt.strftime('%Y-%m-%d')

    df['r'] = df['r'].astype(str)
    df['qty'] = df['qty'].astype(str)
    df['mp'] = df['mp'].astype(str)
    df['pallets'] = df['pallets'].astype(str)

    return df


def connect_to_sqlite_and_read_into_df():
    try:
        with sqlite3.connect(sql_db) as conn:
            df = pd.read_sql_query(
                """SELECT * FROM 'Released Schedule' WHERE lot NOT LIKE '%W' ORDER BY run_date_time DESC""", conn)

    except Exception as e:
        print(e)
        exit(1)

    return df


def concat_dfs_drop_duplicates_drop_table_insert_table(df1, df2):
    df2 = df2.drop(columns=['id'])

    df = pd.concat([df1, df2], ignore_index=True)

    df = df.drop_duplicates(subset=['lot'])

    print("Original {} -> Updated {}".format(len(df2), len(df)))

    try:
        with sqlite3.connect(sql_db) as conn:

            df.to_sql('Released Schedule', con=conn,
                      if_exists='replace', index=False)

    except Exception as e:
        print(e)
        exit(1)

    return df


if __name__ == "__main__":
    try:
        decrypt_file(copy_file(file))
    except Exception as e:
        print(e)

    df = create_df()
    df2 = connect_to_sqlite_and_read_into_df()

    df3 = concat_dfs_drop_duplicates_drop_table_insert_table(df, df2)

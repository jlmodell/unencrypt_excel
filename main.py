from datetime import datetime
import msoffcrypto
import os
import sys
from shutil import copyfile
from dotenv import load_dotenv
import io
import pandas as pd

load_dotenv()

KEY = os.environ.get("KEY", "")
assert KEY != "", f"Error::{datetime.now():%Y-%m-%d %H:%M:%S}::`Message: KEY environment variable is not set `"

file = r'/mnt/busse3/fs1/pkg_schedule/Released Schedule.xls'
if len(sys.argv) > 1:
    file = sys.argv[1]

print(file)

def copy_file(src) -> str:
    global file

    dst = os.path.join(os.getcwd(), os.path.basename(src))    

    copyfile(src, dst)

    print(f"Completed Task::{datetime.now():%Y-%m-%d %H:%M:%S}::`Copied file: {src} to {dst} `")

    file = dst

    return dst


def decrypt_file(src) -> None:
    global KEY

    dst = os.path.join(os.getcwd(), "_" + os.path.basename(src))
    new_dst=dst.replace('.xls', '.xlsx')

    assert os.path.exists(src), f"Error::{datetime.now():%Y-%m-%d %H:%M:%S}::`Message: {src} does not exist `"

    if os.path.exists(dst) or os.path.exists(new_dst):
        try:
            os.remove(dst)
        except:
            try:
                os.remove(new_dst)
            except:
                pass

    try:
        with open(src, "rb") as f:
            bytes = io.BytesIO(f.read())
        
        dst_bytes = io.BytesIO()
        
        ms = msoffcrypto.OfficeFile(bytes)
        ms.load_key(password=KEY)
        ms.decrypt(dst_bytes)
        print(f"Completed Task::{datetime.now():%Y-%m-%d %H:%M:%S}::`Decrypted file: {src} `")

        df = pd.read_excel(dst_bytes, sheet_name="Schedule")        

    except Exception as e:
        print(f"Error::{datetime.now():%Y-%m-%d %H:%M:%S}::`Message: File was not encrypted `")
        
        if os.environ.get("DEBUG", None) == "true":
            print(f"Details::{datetime.now():%Y-%m-%d %H:%M:%S}::`...` -- EOF")
            print("\t", e)
            print("EOF")
        
        df = pd.read_excel(src, sheet_name="Schedule")
    
    df.to_excel(new_dst, index=False)
    print(f"Completed Task::{datetime.now():%Y-%m-%d %H:%M:%S}::`File written: {new_dst} `")

    os.remove(src)
    print(f"Completed Task::{datetime.now():%Y-%m-%d %H:%M:%S}::`Deleted file: {src} `")


if __name__ == "__main__":
    print(f"Full Task Started::`{datetime.now():%Y-%m-%d %H:%M:%S} `")

    decrypt_file(copy_file(file))

    print(f"Full Task Finished::`{datetime.now():%Y-%m-%d %H:%M:%S} `")
    print()
    print("----------------------------------------------------")
    print()
    print(f"Test Task Started::`{datetime.now():%Y-%m-%d %H:%M:%S} `")

    try:
        df = pd.read_excel(r"_Released Schedule.xlsx")
        print(f"Completed Task::{datetime.now():%Y-%m-%d %H:%M:%S}::`Result: file was successfully decrypted `")
    except Exception as e:
        print(f"Error::{datetime.now():%Y-%m-%d %H:%M:%S}::`Result: file was not decrypted `")
        if os.environ.get("DEBUG", None) == "true":
            print(f"Details::{datetime.now():%Y-%m-%d %H:%M:%S}::`...` -- EOF")
            print("\t", e)
            print("EOF")

    print(f"Full Task Finished::`{datetime.now():%Y-%m-%d %H:%M:%S} `")
'''
import subprocess
import getpass
import pandas as pd
import argparse
import logging
from concurrent.futures import ThreadPoolExecutor
from tqdm import tqdm

class IntegrityClient:
    def __init__(self, hostname="integrity", port="7001", username=None):
        self.hostname = hostname
        self.port = port
        self.username = username
        self.password = None
        self.logger = logging.getLogger(__name__)

    def login(self):
        self.password = getpass.getpass("Enter your Integrity password: ")

    def batch_edit_items(self, item_ids, fields):
        cmd_list = [
            "im", "editissue",
            "--batchEdit",
            f"--hostname={self.hostname}",
            f"--port={self.port}",
            f"--user={self.username}",
            f"--password={self.password}"
        ]
        for field, value in fields.items():
            cmd_list.append(f'--field={field}={value}')
        cmd_list.extend(map(str, item_ids))

        try:
            result = subprocess.run(cmd_list, capture_output=True, text=True, check=True)
            self.logger.info(f" Updated items {item_ids} with fields {fields}")
        except subprocess.CalledProcessError as e:
            self.logger.error(f" Failed to update items {item_ids}: {e.stderr}")

def process_excel(file_path, client):
    df = pd.read_excel(file_path)
    field_cols = [col for col in df.columns if col != "ItemID"]
    grouped = df.groupby(field_cols)

    # Progress bar for batches
    total_batches = len(grouped)
    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = []
        for fields_values, group in tqdm(grouped, total=total_batches, desc="Updating Items", unit="batch"):
            fields = dict(zip(field_cols, fields_values if isinstance(fields_values, tuple) else [fields_values]))
            item_ids = group["ItemID"].tolist()
            futures.append(executor.submit(client.batch_edit_items, item_ids, fields))

if __name__ == "__main__":
    logging.basicConfig(filename="bulk_update.log", level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s')
    parser = argparse.ArgumentParser(description="Bulk update Integrity items from Excel (optimized with progress bar)")
    parser.add_argument("--username", required=True, help="Integrity username")
    parser.add_argument("--file", required=True, help="Path to Excel file")
    args = parser.parse_args()

    client = IntegrityClient(username=args.username)
    client.login()
    process_excel(args.file, client)

'''



import subprocess
import getpass
import pandas as pd
import argparse
import logging
from concurrent.futures import ThreadPoolExecutor
from tqdm import tqdm

class IntegrityClient:
    def __init__(self, hostname="integrity", port="7001", username=None):
        self.hostname = hostname
        self.port = port
        self.username = username
        self.password = None
        self.logger = logging.getLogger(__name__)

    def login(self):
        self.password = getpass.getpass("Enter your Integrity password: ")

    def batch_edit_items(self, item_ids, fields):
        cmd_list = [
            "im", "editissue",
            "--batchEdit",
            f"--hostname={self.hostname}",
            f"--port={self.port}",
            f"--user={self.username}",
            f"--password={self.password}"
        ]
        for field, value in fields.items():
            cmd_list.append(f'--field={field}={value}')
        cmd_list.extend(map(str, item_ids))

        # Show in terminal which items are being updated
        print(f"Updating items: {item_ids} with fields: {fields}")

        try:
            result = subprocess.run(cmd_list, capture_output=True, text=True, check=True)
            self.logger.info(f" Updated items {item_ids} with fields {fields}")
        except subprocess.CalledProcessError as e:
            self.logger.error(f" Failed to update items {item_ids}: {e.stderr}")

def process_excel(file_path, client):
    df = pd.read_excel(file_path)
    field_cols = [col for col in df.columns if col != "ItemID"]
    grouped = df.groupby(field_cols)

    total_batches = len(grouped)
    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = []
        for fields_values, group in tqdm(grouped, total=total_batches, desc="Processing Batches", unit="batch"):
            fields = dict(zip(field_cols, fields_values if isinstance(fields_values, tuple) else [fields_values]))
            item_ids = group["ItemID"].tolist()
            futures.append(executor.submit(client.batch_edit_items, item_ids, fields))

if __name__ == "__main__":
    logging.basicConfig(filename="bulk_update.log", level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s')
    parser = argparse.ArgumentParser(description="Bulk update Integrity items from Excel (with progress and item display)")
    parser.add_argument("--username", required=True, help="Integrity username")
    parser.add_argument("--file", required=True, help="Path to Excel file")
    args = parser.parse_args()

    client = IntegrityClient(username=args.username)
    client.login()
    process_excel(args.file, client)

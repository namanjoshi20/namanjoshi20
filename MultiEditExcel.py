
import subprocess
import getpass
import pandas as pd
import argparse
import logging

class IntegrityClient:
    def __init__(self, hostname="integrity", port="7001", username=None):
        self.hostname = hostname
        self.port = port
        self.username = username
        self.password = None
        self.logger = logging.getLogger(__name__)

    def login(self):
        self.password = getpass.getpass("Enter your Integrity password: ")

    def edit_item(self, item_id, fields):
        cmd_list = [
            "im", "editissue",
            f"--hostname={self.hostname}",
            f"--port={self.port}",
            f"--user={self.username}",
            f"--password={self.password}"
        ]
        for field, value in fields.items():
            cmd_list.append(f'--field={field}={value}')
        cmd_list.append(str(item_id))

        try:
            result = subprocess.run(cmd_list, capture_output=True, text=True, check=True)
            self.logger.info(f"Item {item_id} updated successfully:\n{result.stdout}")
        except subprocess.CalledProcessError as e:
            self.logger.error(f"Failed to update item {item_id}:\n{e.stderr}")

def process_excel(file_path, client):
    try:
        df = pd.read_excel(file_path)
        for _, row in df.iterrows():
            item_id = row["ItemID"]
            fields = {col: str(row[col]) for col in df.columns if col != "ItemID" and pd.notna(row[col])}
            client.logger.info(f"Updating Item {item_id} with fields: {fields}")
            client.edit_item(item_id, fields)
    except Exception as e:
        client.logger.error(f"Failed to process Excel file: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Bulk update Integrity items from Excel")
    parser.add_argument("--username", required=True, help="Integrity username")
    parser.add_argument("--file", required=True, help="Path to Excel file")
    args = parser.parse_args()

    client = IntegrityClient(username=args.username)
    client.login()
    process_excel(args.file, client)
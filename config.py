import json

from reportlab.lib.units import mm
from reportlab.lib.pagesizes import letter, legal, A4, A3


def read_config(file_path):
    try:
        with open(file_path, 'r') as config_file:
            config_data = json.load(config_file)
        return config_data
    except FileNotFoundError:
        print(f"Config file not found at: {file_path}")
        return None
    except json.JSONDecodeError:
        print(f"Error decoding JSON in config file: {file_path}")
        return None


def update_config(file_path, new_config_data):
    try:
        with open(file_path, 'w') as config_file:
            json.dump(new_config_data, config_file, indent=2)
        print("Config file updated successfully.")
    except json.JSONDecodeError:
        print(f"Error decoding JSON in config file: {file_path}")

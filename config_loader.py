import json

def load_config(config_path = 'config.json'):
    with open(config_path, 'r') as config_file:
        return json.load(config_file)

CONFIG = load_config()
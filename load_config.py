import json

def load_config(config_path='config.json'):
    """
    This function loads the config.json file that contains sensitive information like :
    Local Database Credentials
    API Key
    """
    with open(config_path) as config_file:
        config = json.load(config_file)
    return config

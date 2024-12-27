#!filepath: client_config.py
# changed client_config.py: Added VERSION attribute and removed unused import.
class ClientConfig(object):
    HTTP_TIMEOUT = 30
    UPDATE_URLS = ['http://localhost:8000/updates']
    VERSION = "0.0.2"
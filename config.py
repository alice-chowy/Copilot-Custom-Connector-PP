import os
from dotenv import load_dotenv

load_dotenv()

CONFIG = {
    "tenant_id": os.getenv("TENANT_ID"),
    "client_id": os.getenv("CLIENT_ID"),
    "client_secret": os.getenv("CLIENT_SECRET"),
}

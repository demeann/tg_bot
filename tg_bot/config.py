from dataclasses import dataclass
from dotenv import load_dotenv
from os import getenv

load_dotenv()

@dataclass
class Settings:
    bot_token: str = getenv("TELEGRAM_TOKEN", "")
    admin_ids: list[int] = None

    def __post_init__(self):
        raw = getenv("ADMIN_IDS", "")
        self.admin_ids = [int(x.strip()) for x in raw.split(",") if x.strip().isdigit()]

settings = Settings()

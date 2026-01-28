from pydantic_settings import BaseSettings
from pydantic import field_validator
from functools import lru_cache


class Settings(BaseSettings):
    azure_client_id: str
    azure_client_secret: str
    azure_tenant_id: str
    azure_redirect_uri: str = "http://localhost:8365/auth/callback"
    database_url: str = "sqlite+aiosqlite:///./data/ms365_access.db"
    secret_key: str
    app_port: int = 8365

    # Security settings
    allowed_hosts: list[str] = ["localhost", "127.0.0.1"]
    cors_origins: list[str] = ["http://localhost:5173", "http://127.0.0.1:5173"]

    @field_validator("allowed_hosts", "cors_origins", mode="before")
    @classmethod
    def parse_comma_separated(cls, v):
        if isinstance(v, str):
            return [item.strip() for item in v.split(",")]
        return v

    # Server binding (use 0.0.0.0 only in Docker/container environments)
    app_host: str = "127.0.0.1"

    # Local timezone for calendar display (override with LOCAL_TIMEZONE env var)
    local_timezone: str = "UTC"

    # Microsoft Graph API
    graph_base_url: str = "https://graph.microsoft.com/v1.0"
    authority_base: str = "https://login.microsoftonline.com"

    # OAuth scopes
    scopes: list[str] = [
        "User.Read",
        "Mail.ReadWrite",
        "Mail.Send",
        "Calendars.ReadWrite",
        "Files.ReadWrite.All",
        "offline_access",
    ]

    @property
    def authority(self) -> str:
        return f"{self.authority_base}/{self.azure_tenant_id}"

    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"


@lru_cache
def get_settings() -> Settings:
    return Settings()

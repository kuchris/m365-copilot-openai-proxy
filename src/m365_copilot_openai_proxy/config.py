from __future__ import annotations

from typing import Literal

from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    access_token: str = Field(default="", alias="M365_ACCESS_TOKEN")
    time_zone: str = Field(default="Asia/Tokyo", alias="M365_TIME_ZONE")
    model_alias: str = Field(default="m365-copilot", alias="M365_MODEL_ALIAS")
    backend: Literal["substrate", "websocket", "browser"] = Field(default="substrate", alias="M365_BACKEND")
    cdp_port: int = Field(default=9222, alias="M365_CDP_PORT", ge=1, le=65535)
    browser_timeout_seconds: float = Field(default=120, alias="M365_BROWSER_TIMEOUT_SECONDS", gt=0)
    credential_refresh_seconds: float = Field(default=2700, alias="M365_CREDENTIAL_REFRESH_SECONDS", gt=0)
    browser_auto_verify: bool = Field(default=False, alias="M365_BROWSER_AUTO_VERIFY")

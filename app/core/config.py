from pydantic_settings import BaseSettings
import os


class Settings(BaseSettings):
    PROJECT_NAME: str = "FastAPI Project"
    PROJECT_DESCRIPTION: str = "FastAPI Project Description"
    SERVER_HOST: str = "0.0.0.0"
    SERVER_PORT: int = 8000
    DEBUG_MODE: bool = True
    ALLOWED_ORIGINS: list = ["http://localhost:3000"]

    # Supabase 설정
    SUPABASE_URL: str = ""
    SUPABASE_KEY: str = ""

    # Sellmate 설정
    sellmate_id: str
    sellmate_pw: str
    download_path: str
    base_url: str
    smss: str

    class Config:
        env_file = ".env"


settings = Settings()

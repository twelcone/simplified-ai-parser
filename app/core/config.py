from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    app_name: str = "Simplified AI Parser"
    max_file_size_mb: int = 50

    class Config:
        env_file = ".env"


settings = Settings()

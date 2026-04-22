from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """
    Carrega e valida as configurações do ambiente a partir de variáveis de ambiente
    ou de um arquivo .env.
    """

    DB_DRIVER: str = "SQL Server"
    DB_SERVER: str
    DB_NAME: str
    DB_USER: str
    DB_PASSWORD: str

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

settings = Settings()

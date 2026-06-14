from pydantic_settings import BaseSettings
from pathlib import Path


class Settings(BaseSettings):
    # Gemini OCR
    gemini_api_key: str = ""
    gemini_model: str = "gemini-3.1-flash-lite"
    gemini_prompt: str = ""
    gemini_json_schema: str = ""
    proxy: str = ""

    # 数据目录
    data_root: str = ""
    project_subdir: str = "projects"
    template_subdir: str = "template"

    model_config = {
        "env_file": str(Path(__file__).resolve().parent.parent / ".env"),
        "env_file_encoding": "utf-8",
    }

    @property
    def data_dir(self) -> Path:
        if self.data_root:
            return Path(self.data_root)
        return Path(__file__).resolve().parent.parent / "data"

    @property
    def project_dir(self) -> Path:
        return self.data_dir / self.project_subdir

    @property
    def ratecard_dir(self) -> Path:
        return self.data_dir / "ratecard"

    @property
    def template_dir(self) -> Path:
        return self.data_dir / self.template_subdir


settings = Settings()

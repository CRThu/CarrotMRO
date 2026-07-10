from pydantic_settings import BaseSettings
from pathlib import Path


class Settings(BaseSettings):
    # litellm OCR
    litellm_model: str = "xiaomi_mimo/mimo-v2-flash"
    xiaomi_mimo_api_key: str = ""
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

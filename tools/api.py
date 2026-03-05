from openai import OpenAI
import yaml
from pathlib import Path

def load_config(filename: str = "config.yaml") -> dict:
    base_dir = Path(__file__).resolve().parent.parent
    config_path = base_dir / filename
    
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def call(prompt: str, dump: str) -> str:
    config = load_config()
    cfg = config["openrouter"]

    message = prompt + dump

    client = OpenAI(
        base_url = cfg["base_url"],
        api_key = cfg["api_key"]
    )

    answer = client.chat.completions.create(
        model=cfg["model"],
        messages = [{"role": "user", "content": message}]
    )

    return answer.choices[0].message.content.strip()
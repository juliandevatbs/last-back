import logging
from pathlib import Path

PROMPT_DIR = Path(__file__).parent

logger = logging.getLogger(__name__)

def load_prompt(prompt_name: str, **kwargs) -> str:

    prompt_path = PROMPT_DIR / f"{prompt_name}.txt"

    if not prompt_path.exists():

        logger.error(f"Prompt {prompt_name} no existe en {PROMPT_DIR}")
        raise FileNotFoundError(f"Prompt {prompt_name} no existe en {PROMPT_DIR}")

    with open(prompt_path, "r", encoding="utf-8") as f:

        template = f.read()

    import re
    placeholders = re.findall(r'\{(\w+)}', template)
    missing = set(placeholders) - set(kwargs.keys())

    if missing:

        logger.warning(f"Placeholders faltantes {missing}")
    
    
    return template.format(**kwargs)




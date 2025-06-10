import os

CONFIG_FILE = "deepseek_config.txt"


def carregar_clau_deepseek():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return f.read().strip()
    return ""


def guardar_clau_deepseek(clau):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        f.write(clau)
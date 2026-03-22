import json
from yandex_api import YandexDiskAPI
from config import CONFIG
import traceback

api = YandexDiskAPI(CONFIG["TOKEN_ITEMS"])

def explore_path(path, depth=0):
    if depth > 2:
        return
    indent = "  " * depth
    try:
        items = api.get_folder_contents(path, use_cache=False)
        if items:
            print(f"{indent}Scanning: {path} ({len(items)} items)")
            for i in items:
                print(f"{indent} - [{i['type']}] {i['name']}")
                # If it's a directory and it doesn't look like a point folder, dive in
                if i['type'] == 'dir' and not str(i['name']).startswith('п.'):
                     explore_path(i['path'], depth + 1)
        else:
            print(f"{indent}Empty or error: {path}")
    except Exception as e:
        print(f"{indent}Exception for {path}: {e}")

explore_path("/Общая папка/ПНОС/2026/ЭПБ/39-40")

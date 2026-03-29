# -*- coding: utf-8 -*-
import os
import json
import ctypes

STATUS_FILE_NAME = ".pnos_status.json"

def get_status_file_path(point_folder):
    return os.path.join(point_folder, STATUS_FILE_NAME)

def hide_file_windows(filepath):
    """
    Скрывает файл в ОС Windows, устанавливая атрибут FILE_ATTRIBUTE_HIDDEN (0x02).
    """
    try:
        # FILE_ATTRIBUTE_HIDDEN = 0x02
        if os.name == 'nt' and os.path.exists(filepath):
            ctypes.windll.kernel32.SetFileAttributesW(filepath, 0x02)
    except Exception:
        pass

def unhide_file_windows(filepath):
    """
    Убирает атрибут "Скрытый" с файла перед его перезаписью (устанавливает обычный FILE_ATTRIBUTE_NORMAL 0x80).
    Без этого open(..., "w") в Windows может выдавать PermissionError для уже скрытого файла.
    """
    try:
        if os.name == 'nt' and os.path.exists(filepath):
            ctypes.windll.kernel32.SetFileAttributesW(filepath, 0x80)
    except Exception:
        pass

def get_stage_status(point_folder, stage_key):
    """
    Проверяет, завершен ли конкретный этап для данной папки.
    Например: get_stage_status("C:/PNOS/п.1234", "stage1") -> True/False
    """
    path = get_status_file_path(point_folder)
    if not os.path.exists(path):
        return False
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data.get(stage_key, False)
    except Exception:
        return False

def set_stage_status(point_folder, stage_key, status=True):
    """
    Сохраняет статус этапа в скрытый лог-файл (в формате JSON).
    Создает файл, если его нет, и автоматически его скрывает.
    """
    path = get_status_file_path(point_folder)
    data = {}
    
    # Пытаемся прочитать существующие статусы
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            pass
            
    # Обновляем нужный этап
    data[stage_key] = status
    
    # Записываем обратно
    try:
        # Временно снимаем скрытонасть, чтобы Windows не блокировала перезапись файла на "только чтение" / "скрытый"
        unhide_file_windows(path)
        
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        
        # Снова скрываем
        hide_file_windows(path)
    except Exception as e:
        print(f"Ошибка сохранения статуса для {point_folder}: {e}")


# -*- coding: utf-8 -*-
import os
import json
import ctypes
import re

STATUS_FILE_NAME = ".pnos_status.json"
# Шаг 7 прошёл, но актов ГИ нет: здесь причина. Сам шаг (stage6) не
# отмечается — повторный запуск подхватит пункт, когда акты появятся.
GI_WAIT_KEY = "stage6_wait"

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
    Значение может быть любым JSON-совместимым (bool, число, строка, список, словарь).
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


def get_status_value(point_folder, key, default=None):
    """Универсальный геттер произвольного значения из .pnos_status.json."""
    path = get_status_file_path(point_folder)
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data.get(key, default)
    except Exception:
        return default


def reset_point_status(point_folder):
    """Полностью сбрасывает статус пункта — удаляет .pnos_status.json.
    После этого пункт будет обработан всеми этапами заново."""
    path = get_status_file_path(point_folder)
    if os.path.exists(path):
        try:
            unhide_file_windows(path)
            os.remove(path)
            return True
        except Exception as e:
            print(f"Не удалось сбросить статус {point_folder}: {e}")
            return False
    return True


def reset_stage_status(point_folder, stage_key):
    """Сбрасывает статус одного шага (удаляет ключ), сохраняя остальные."""
    path = get_status_file_path(point_folder)
    if not os.path.exists(path):
        return True
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return False
    data.pop(stage_key, None)
    try:
        unhide_file_windows(path)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        hide_file_windows(path)
        return True
    except Exception as e:
        print(f"Не удалось сбросить шаг {stage_key} для {point_folder}: {e}")
        return False


def is_point_blocked(point_folder):
    """
    Пункт заблокирован для обработки, если он помечен «не корректный» (нет
    обязательных папок) или «пропущен» пользователем на шаге 3.
    Возвращает (blocked: bool, reason: str).
    """
    name = os.path.basename(os.path.normpath(point_folder)).lower()
    if "не корректн" in name:
        return True, "не корректный (нет обязательных папок)"

    path = get_status_file_path(point_folder)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = {}
        if data.get("invalid"):
            miss = data.get("missing_folders") or []
            reason = "не корректный"
            if miss:
                reason += f" (нет: {', '.join(miss)})"
            return True, reason
        if data.get("skipped"):
            return True, "пропущен пользователем"
    return False, ""



# ─── Папки пунктов ───────────────────────────────────────────────

# «п.3186», «п. 3186», «П.3186», «пункт 3186», «3186», «п.3186 (не корректный)»
_POINT_DIR = re.compile(r"^\s*(?:(?:пункт|п)\.?\s*)?(\d+)(?!\d)", re.IGNORECASE)

# Этапы 2–7 берут только папки «п.*»: так в ПНОС не попадают _шаблоны и _gi_cache
POINT_DIR_PREFIX = "п."


def point_number_of(dirname: str) -> str | None:
    """Номер пункта из имени папки или None, если имя на пункт не похоже."""
    m = _POINT_DIR.match(dirname)
    return m.group(1) if m else None


def _normalize_points(points) -> list[str]:
    """Номера без повторов; «п.3186» и « 3186 » сводятся к «3186»."""
    out = []
    for p in points:
        raw = str(p).strip()
        num = point_number_of(raw) or raw
        if num and num not in out:
            out.append(num)
    return out


def select_point_folders(local_root: str, points=None) -> list[str]:
    """
    Папки пунктов в local_root, которые перебирают этапы 2–7: подпапки, чьё имя
    начинается с «п.», в порядке os.listdir. Если points задан (номера-строки) —
    только папки с этими номерами; points=None — все пункты, [] — ни одного.
    Заблокированные пункты не отсеиваются: их пропускает сам этап, со своим
    сообщением. Нет папки local_root — пустой список.
    """
    try:
        names = os.listdir(local_root)
    except OSError:
        return []
    wanted = None if points is None else set(_normalize_points(points))
    folders = []
    for d in names:
        dpath = os.path.join(local_root, d)
        if not (os.path.isdir(dpath) and d.startswith(POINT_DIR_PREFIX)):
            continue
        if wanted is not None and point_number_of(d) not in wanted:
            continue
        folders.append(dpath)
    return folders


def missing_points(folders: list[str], points) -> list[str]:
    """Номера из points, для которых среди folders нет папки."""
    found = {point_number_of(os.path.basename(f)) for f in folders}
    return [p for p in _normalize_points(points) if p not in found]


def no_selected_points_message(points) -> str:
    """Текст для журнала, когда среди выбранных пунктов нет ни одной папки."""
    nums = ", ".join(f"п.{p}" for p in _normalize_points(points)) or "список пуст"
    return f"Выбранных пунктов нет в папке ПНОС ({nums})"

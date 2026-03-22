# -*- coding: utf-8 -*-
"""
Модуль работы с Яндекс Диском через REST API.
Поддерживает пагинацию, рекурсивное скачивание, нечёткий поиск папок.
Кеширует результаты get_folder_contents для ускорения повторных запросов.
Теперь полностью АСИНХРОННЫЙ с использованием aiohttp и aiofiles!
"""

import os
import aiohttp
import aiofiles
import asyncio
from urllib.parse import quote
from rapidfuzz import fuzz, process

from config import CONFIG


class YandexDiskAPI:
    """Клиент для работы с REST API Яндекс.Диска."""

    BASE_URL = "https://cloud-api.yandex.net/v1/disk/resources"
    PAGE_LIMIT = 100  # Макс. элементов за один запрос (API макс = 100)

    def __init__(self, token: str):
        self.token = token
        self.headers = {"Authorization": f"OAuth {self.token}"}
        self._cache: dict[str, list] = {}  # Кеш: путь → список items
        self._log_callback = None
        # Защита от бана Яндексом (HTTP 429 Too Many Requests)
        self.semaphore = asyncio.Semaphore(50) 

    def set_log_callback(self, callback):
        """Устанавливает функцию для логирования (для отладки)."""
        self._log_callback = callback

    def _log(self, msg: str):
        if self._log_callback:
            self._log_callback(msg)

    def clear_cache(self):
        """Очищает кеш содержимого папок."""
        self._cache.clear()

    # ─── Получение содержимого папки (с пагинацией и кешем) ──────────

    async def get_folder_contents(self, session: aiohttp.ClientSession, path: str, use_cache: bool = True) -> list | None:
        """
        Получает полный список файлов и папок в указанной директории.
        Обрабатывает пагинацию и кеширует результаты.
        """
        if use_cache and path in self._cache:
            return self._cache[path]

        all_items = []
        offset = 0

        try:
            while True:
                params = {
                    "path": path,
                    "limit": self.PAGE_LIMIT,
                    "offset": offset,
                }
                async with self.semaphore:
                    async with session.get(self.BASE_URL, headers=self.headers, params=params) as response:
                        if response.status != 200:
                            self._log(f"   ⚠ API [{response.status}]: {path}")
                            if response.status == 404:
                                self._log(f"   ⚠ Папка не найдена на Яндекс Диске: {path}")
                            return None if not all_items else all_items

                        data = await response.json()
                        embedded = data.get("_embedded", {})
                        items = embedded.get("items", [])
                        all_items.extend(items)

                        total = embedded.get("total", 0)
                        if offset + len(items) >= total:
                            break
                        offset += self.PAGE_LIMIT

            self._cache[path] = all_items
            return all_items

        except Exception as e:
            self._log(f"   ⚠ Ошибка сети: {e}")
            return None

    async def preload_search_paths(self, session: aiohttp.ClientSession, search_paths: list[str]) -> list[str]:
        """
        Предзагружает содержимое подпалок "ТТП" внутри указанных путей.
        Возвращает список путей к этим подпапкам "ТТП".
        Загрузка происходит параллельно для максимального ускорения.
        """
        self._log(f"   📂 Предзагрузка подпапок 'ТТП' (асинхронно)...")
        extended_paths = []
        loaded = 0

        async def fetch_ttp(base_path):
            ttp_path = f"{base_path}/ТТП"
            items = await self.get_folder_contents(session, ttp_path, use_cache=False)
            return base_path, ttp_path, items

        # Запускаем все запросы одновременно!
        tasks = [fetch_ttp(p) for p in search_paths]
        results = await asyncio.gather(*tasks, return_exceptions=True)

        for res in results:
            if isinstance(res, tuple):
                base_path, ttp_path, items = res
                if items is not None:
                    loaded += 1
                    extended_paths.append(ttp_path)
                else:
                    self._log(f"   ✗ Не удалось загрузить: {base_path.split('/')[-1]}/ТТП")
            else:
                self._log(f"   ✗ Ошибка предзагрузки: {res}")

        self._log(f"   ✓ Успешно загружено {loaded}/{len(search_paths)} папок 'ТТП'")
        return extended_paths

    # ─── Скачивание файла ───────────────────────────────────────────────

    async def download_file(self, session: aiohttp.ClientSession, remote_path: str, local_path: str) -> bool:
        """Скачивает один файл с Я.Диска по указанному пути."""
        try:
            params = {"path": remote_path}
            async with self.semaphore:
                async with session.get(f"{self.BASE_URL}/download", headers=self.headers, params=params) as resp:
                    if resp.status != 200:
                        self._log(f"   ⚠ Не удалось получить ссылку [{resp.status}]: {remote_path}")
                        return False
                    data = await resp.json()
                    download_url = data["href"]

            async with self.semaphore:
                async with session.get(download_url) as r:
                    r.raise_for_status()
                    os.makedirs(os.path.dirname(local_path), exist_ok=True)
                    # Используем aiofiles для асинхронной записи на диск
                    async with aiofiles.open(local_path, "wb") as f:
                        async for chunk in r.content.iter_chunked(8192):
                            await f.write(chunk)
            return True

        except Exception as e:
            self._log(f"   ⚠ Ошибка скачивания: {e}")
            return False

    # ─── Рекурсивное скачивание папки ───────────────────────────────

    async def download_folder_recursive(self, session: aiohttp.ClientSession, remote_path: str, local_path: str, log_callback=None) -> int:
        """
        Рекурсивно скачивает содержимое папки АСИНХРОННО.
        Возвращает количество скачанных файлов.
        """
        os.makedirs(local_path, exist_ok=True)
        items = await self.get_folder_contents(session, remote_path)
        if not items:
            return 0

        tasks = []
        
        async def dl_item(item):
            name = item["name"]
            item_path = item["path"]

            if item["type"] == "file":
                local_file = os.path.join(local_path, name)
                if not os.path.exists(local_file):
                    success = await self.download_file(session, item_path, local_file)
                    if success:
                        if log_callback: log_callback(f"    ↓ {name}")
                        return 1
                    else:
                        if log_callback: log_callback(f"    ✗ Ошибка: {name}")
                        return 0
                return 0
            elif item["type"] == "dir":
                sub_local = os.path.join(local_path, name)
                return await self.download_folder_recursive(session, item_path, sub_local, log_callback)
            return 0

        for item in items:
            tasks.append(dl_item(item))

        # Ждем скачивания всех файлов папки параллельно
        results = await asyncio.gather(*tasks)
        return sum(results)

    # ─── Нечёткий поиск папки по имени ──────────────────────────────

    async def find_folder_by_name(self, session: aiohttp.ClientSession, parent_path: str, target_name: str, threshold: int = None) -> str | None:
        """
        Ищет папку с похожим именем внутри родительской папки.
        """
        if threshold is None:
            threshold = CONFIG["FUZZY_THRESHOLD"]

        items = await self.get_folder_contents(session, parent_path)
        if not items:
            return None

        folders = [item for item in items if item["type"] == "dir"]
        if not folders:
            return None

        names = [f["name"] for f in folders]

        for item in folders:
            if item["name"].lower() == target_name.lower():
                return item["path"]

        match = process.extractOne(target_name, names, scorer=fuzz.ratio)
        if match and match[1] >= threshold:
            found_name = match[0]
            for item in folders:
                if item["name"] == found_name:
                    return item["path"]
        return None

    # ─── Поиск папки пункта п.{номер} ──────────────────────────────

    async def find_point_folder(self, session: aiohttp.ClientSession, search_paths: list[str], point_number: str) -> tuple[str | None, str | None]:
        """
        Ищет папку пункта вида п.{номер} в списке указанных папок.
        """
        expected_name = f"п.{point_number}"
        point_norm = point_number.strip().lower()

        async def check_path(folder_path):
            items = await self.get_folder_contents(session, folder_path)
            if not items:
                return None
                
            folders = [item for item in items if item["type"] == "dir"]
            if not folders:
                return None

            for item in folders:
                item_name_norm = item["name"].replace(" ", "").lower()
                if f"п.{point_norm}" in item_name_norm:
                    return item["path"], folder_path

            names = [f["name"] for f in folders]
            match = process.extractOne(expected_name, names, scorer=fuzz.ratio)
            if match and match[1] >= CONFIG["FUZZY_THRESHOLD"]:
                found_name = match[0]
                for item in folders:
                    if item["name"] == found_name:
                        return item["path"], folder_path
            return None

        # Запускаем поиск по всем путям ОДНОВРЕМЕННО
        tasks = [check_path(p) for p in search_paths]
        results = await asyncio.gather(*tasks)
        
        for res in results:
            if res is not None:
                return res
                
        return None, None

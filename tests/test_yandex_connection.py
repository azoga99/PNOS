# -*- coding: utf-8 -*-
"""
Проверка связи с Яндекс.Диском перед этапом 1.

Раньше обрыв сети никак не назывался: каждый запрос тихо возвращал None,
и этап 1 писал «НЕ НАЙДЕН» по всем пунктам — это принимали за пустые папки
на диске. На ПК с VPN sing-box DNS не отвечал на cloud-api.yandex.net,
хотя ya.ru открывался.
"""

import asyncio
import socket

import aiohttp
import pytest
from aiohttp import web
from aiohttp.client_reqrep import ConnectionKey
from aiohttp.test_utils import TestServer

from yandex_api import API_HOST, YandexDiskAPI, connection_problem

KEY = ConnectionKey(*([None] * len(ConnectionKey._fields)))


def test_dns_failure_is_named():
    exc = aiohttp.ClientConnectorDNSError(KEY, socket.gaierror(11001, "getaddrinfo failed"))
    text = connection_problem(exc)
    assert "DNS" in text and API_HOST in text and "VPN" in text


def test_refused_connection_is_named():
    exc = aiohttp.ClientConnectorError(KEY, ConnectionRefusedError(10061, "refused"))
    text = connection_problem(exc)
    assert "не удалось подключиться" in text and "DNS" not in text


def test_timeout_is_named():
    assert "не ответил" in connection_problem(asyncio.TimeoutError())


def _check(status: int | None):
    """check_connection против локального сервера, отвечающего status."""
    async def handler(request):
        return web.json_response({}, status=status)

    async def run():
        app = web.Application()
        app.router.add_get("/v1/disk/", handler)
        async with TestServer(app) as server:
            api = YandexDiskAPI("token")
            api.DISK_URL = str(server.make_url("/v1/disk/"))
            async with aiohttp.ClientSession() as session:
                return await api.check_connection(session)

    return asyncio.run(run())


def test_ok_connection():
    assert _check(200) is None


@pytest.mark.parametrize("status", [401, 403])
def test_rejected_token(status):
    assert "токен" in _check(status)


def test_server_error():
    assert "HTTP 503" in _check(503)

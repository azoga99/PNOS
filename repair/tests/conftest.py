# -*- coding: utf-8 -*-
"""Общие помощники для тестов модуля «Ремонт»."""

import json
import os

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")


def load_fixture(point: str) -> dict:
    with open(os.path.join(FIXTURES, f"{point}.json"), encoding="utf-8") as f:
        return json.load(f)


def sections_from(rows):
    """dict-строки фикстуры → объекты Section."""
    from repair.model import Section
    return [
        Section(
            number=r["number"], element=r["element"],
            thickness_nom=r["thickness_nom"], diameter=r["diameter"],
            measures=list(r["measures"]), tmin=r["tmin"],
        )
        for r in rows
    ]


def replaced_from(fx) -> dict:
    return {int(k): float(v) for k, v in fx["replaced"].items()}


def perehod_from(fx) -> dict:
    return {int(k): (v[0], v[1]) for k, v in fx["perehod"].items()}

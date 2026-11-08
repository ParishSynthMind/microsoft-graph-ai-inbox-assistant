import random
from datetime import date, datetime

import pytest

from graph_assistant.assistant import Assistant
from graph_assistant.config import Config
from graph_assistant.extract import Extraction, check_extraction
from graph_assistant.fake_graph import FakeGraph
from graph_assistant.graph import GraphClient
from graph_assistant.store import Store

NOW = datetime(2026, 11, 9, 9, 0)  # Monday 09:00 local


def field(value, conf="high", snippet=""):
    return {"value": value, "confidence": conf, "source_snippet": snippet}


class FakeExtractor:
    """Stands in for Claude: answers from a table keyed by subject, then the real checks run."""

    def __init__(self, answers, cost=0.002):
        self.answers = answers
        self.cost = cost
        self.calls = 0

    def extract(self, text, received: date, tz, costs):
        costs.check(text)
        self.calls += 1
        subject = text.split("\n")[1][len("Subject: "):]
        raw = self.answers[subject]
        costs.add(1000, 200)
        return Extraction(check_extraction(raw, text, received), self.cost, 1000, 200)


@pytest.fixture
def fake():
    return FakeGraph()


@pytest.fixture
def graph(fake):
    return GraphClient(fake, lambda: fake.token, sleep=lambda s: None, rng=random.Random(1))


@pytest.fixture
def store(tmp_path):
    return Store(tmp_path / "state.db")


@pytest.fixture
def cfg():
    c = Config()
    c.since_days = 3650
    return c


def make_assistant(graph, extractor, store, cfg):
    a = Assistant(graph, extractor, store, cfg, now=NOW, log=lambda *a: None)
    a.setup()
    return a

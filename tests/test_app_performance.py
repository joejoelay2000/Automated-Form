import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import app


def test_load_companies_prefers_local_cache(monkeypatch):
    calls = {"github": 0, "local": 0}

    def fake_github_config():
        return ("token", "owner/repo", "main")

    def fake_load_github_companies():
        calls["github"] += 1
        return [{"klien": "Remote", "company_name": "Remote"}]

    def fake_load_local_companies():
        calls["local"] += 1
        return [{"klien": "Local", "company_name": "Local"}]

    monkeypatch.setattr(app, "get_github_config", fake_github_config)
    monkeypatch.setattr(app, "load_github_companies", fake_load_github_companies)
    monkeypatch.setattr(app, "load_local_companies", fake_load_local_companies)
    monkeypatch.setattr(app, "get_database", lambda: None)

    app.load_companies.clear()
    first = app.load_companies()
    second = app.load_companies()

    assert first == second == [{"klien": "Local", "company_name": "Local"}]
    assert calls["local"] == 1
    assert calls["github"] == 0

import pytest
from fastapi.testclient import TestClient
from backend.app import app

client = TestClient(app)

def test_theme_server_side_persistence():
    # 1. 设为浅色模式
    res1 = client.post("/api/settings", json={"theme": "light"})
    assert res1.status_code == 200
    get_res1 = client.get("/")
    assert 'data-theme="light"' in get_res1.text

    # 2. 设为深色模式
    res2 = client.post("/api/settings", json={"theme": "dark"})
    assert res2.status_code == 200
    get_res2 = client.get("/")
    assert 'data-theme="dark"' in get_res2.text

"""Тесты для pingcheck."""
from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest

import pingcheck as pc

# --- make_output_filename ---

def test_make_output_filename():
    ts = datetime(2026, 11, 4, 14, 5, 9)
    assert pc.make_output_filename(ts) == '2026-11-04_14-05-09.xlsx'


# --- read_ips_from_excel ---

def test_read_ips_from_excel(tmp_path: Path):
    path = tmp_path / 'ips.xlsx'
    pd.DataFrame({'ip': ['8.8.8.8', None, ' 1.1.1.1 ', '']}).to_excel(
        path, index=False,
    )
    assert pc.read_ips_from_excel(path, 'ip') == ('8.8.8.8', '1.1.1.1')


def test_read_ips_missing_column(tmp_path: Path):
    path = tmp_path / 'ips.xlsx'
    pd.DataFrame({'other': ['x']}).to_excel(path, index=False)
    with pytest.raises(KeyError):
        pc.read_ips_from_excel(path, 'ip')


# --- ping_ips (multiping замокан) ---

class FakeHost:
    def __init__(self, address, sent, recv):
        self.address = address
        self.packets_sent = sent
        self.packets_received = recv
        self.avg_rtt = 1.0 if recv else None
        self.packet_loss = 1 - recv / sent if sent else 1.0
        self.jitter = None
        self.max_rtt = 1.0 if recv else None
        self.min_rtt = 1.0 if recv else None


def test_ping_ips_empty():
    assert pc.ping_ips(()) == ()


def test_ping_ips_results(monkeypatch):
    hosts = [
        FakeHost('8.8.8.8', 2, 2),
        FakeHost('1.1.1.1', 2, 1),
        FakeHost('10.0.0.1', 2, 0),
    ]
    monkeypatch.setattr(pc, 'multiping', lambda **kwargs: hosts)
    results = pc.ping_ips(('8.8.8.8', '1.1.1.1', '10.0.0.1'), count=2)
    by_ip = {r[pc.PingColumn.IP]: r for r in results}
    assert by_ip['8.8.8.8'][pc.PingColumn.SUCCESS_PCT] == 100
    assert by_ip['1.1.1.1'][pc.PingColumn.SUCCESS_PCT] == 50
    assert by_ip['10.0.0.1'][pc.PingColumn.SUCCESS_PCT] == 0
    assert by_ip['10.0.0.1'][pc.PingColumn.ALIVE] is False
    # ping_count пробрасывается в каждый результат
    assert all(r[pc.PingColumn.PING_COUNT] == 2 for r in results)


# --- save_results_to_excel ---

def test_save_results_empty(tmp_path: Path):
    path = tmp_path / 'out.xlsx'
    assert pc.save_results_to_excel(path, ()) is False
    assert not path.exists()


def test_save_results_creates_dir_and_file(tmp_path: Path):
    path = tmp_path / 'nested' / 'out.xlsx'
    results = ({
        pc.PingColumn.IP: '8.8.8.8',
        pc.PingColumn.ALIVE: True,
        pc.PingColumn.SUCCESS_PCT: 100,
        pc.PingColumn.PING_COUNT: 2,
        pc.PingColumn.AVG_RTT: 1.0,
        pc.PingColumn.MIN_RTT: 1.0,
        pc.PingColumn.MAX_RTT: 1.0,
        pc.PingColumn.JITTER: None,
        pc.PingColumn.PACKET_LOSS: 0.0,
        pc.PingColumn.WARNINGS: None,
    },)
    assert pc.save_results_to_excel(path, results) is True
    assert path.exists()
    df = pd.read_excel(path)
    assert list(df.columns) == [
        'ip', 'alive', 'success_pct', 'ping_count',
        'avg_rtt', 'min_rtt', 'max_rtt', 'jitter', 'packet_loss',
        'warnings',
    ]
    assert df.iloc[0]['ip'] == '8.8.8.8'


# --- append_result_to_source ---

def test_append_result_to_source(tmp_path: Path):
    path = tmp_path / 'ips.xlsx'
    pd.DataFrame({'ip': ['8.8.8.8', '1.1.1.1']}).to_excel(path, index=False)
    results = (
        {pc.PingColumn.IP: '8.8.8.8', pc.PingColumn.SUCCESS_PCT: 100},
        {pc.PingColumn.IP: '1.1.1.1', pc.PingColumn.SUCCESS_PCT: 50},
    )
    ts = datetime(2026, 11, 4, 14, 5, 9)
    pc.append_result_to_source(path, 'ip', results, ts)
    df = pd.read_excel(path)
    col = '2026-11-04 14:05:09'
    assert col in df.columns
    assert df[col].tolist() == [100, 50]
    # временный файл не остался
    assert not list(tmp_path.glob('*.tmp*'))

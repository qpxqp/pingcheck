"""Тесты для pingcheck."""
from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest
from icmplib.exceptions import ICMPLibError

import pingcheck as pc

# --- Утилиты ---

def test_make_output_filename():
    ts = datetime(2026, 11, 4, 14, 5, 9)
    assert pc.make_output_filename(ts) == '2026-11-04_14-05-09.xlsx'


def test_normalize_ip():
    assert pc.normalize_ip('192.168.001.001') == '192.168.1.1'
    assert pc.normalize_ip('8.8.8.8') == '8.8.8.8'
    assert pc.normalize_ip('localhost') == 'localhost'
    assert pc.normalize_ip('192.168.1') == '192.168.1'
    assert pc.normalize_ip('hello.world.test') == 'hello.world.test'


def test_chunked():
    lst = [1, 2, 3, 4, 5]
    chunks = list(pc.chunked(lst, 2))
    assert chunks == [[1, 2], [3, 4], [5]]
    assert list(pc.chunked([], 10)) == []


# --- Чтение из Excel ---

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


# --- Хелперы для моков ---

class FakeHost:
    """Универсальный фейковый Host для тестов."""
    def __init__(self, address, sent=1, recv=1):
        self.address = address
        self.packets_sent = sent
        self.packets_received = recv
        self.avg_rtt = 1.0 if recv else None
        self.packet_loss = 1 - recv / sent if sent else 1.0
        self.jitter = None
        self.max_rtt = 1.0 if recv else None
        self.min_rtt = 1.0 if recv else None


# --- safe_ping_hosts ---

def test_safe_ping_hosts_success(monkeypatch):
    def mock_multiping(addresses, **kwargs):
        return [FakeHost(a) for a in addresses]

    monkeypatch.setattr(pc, 'multiping', mock_multiping)
    hosts, errors = pc.safe_ping_hosts(
        ['8.8.8.8', '1.1.1.1'],
        count=1, timeout=1, privileged=False, chunk_size=10
    )
    assert len(hosts) == 2
    assert hosts[0] is not None
    assert hosts[0].address == '8.8.8.8'
    assert errors == [None, None]


def test_safe_ping_hosts_fallback(monkeypatch):
    # Групповой пинг всегда падает
    def mock_multiping_fail(addresses, **kwargs):
        raise ICMPLibError('Batch failed')
    monkeypatch.setattr(pc, 'multiping', mock_multiping_fail)

    # Поштучный пинг падает только на "bad.ip"
    def mock_ping(addr, **kwargs):
        if addr == 'bad.ip':
            raise ICMPLibError('Cannot resolve')
        return FakeHost(addr)

    monkeypatch.setattr(pc, 'ping', mock_ping)
    hosts, errors = pc.safe_ping_hosts(
        ['8.8.8.8', 'bad.ip'],
        count=1, timeout=1, privileged=False, chunk_size=10
    )
    assert len(hosts) == 2
    assert hosts[0] is not None
    assert hosts[0].address == '8.8.8.8'
    assert hosts[1] is None
    assert errors[0] is None
    assert errors[1] == 'Cannot resolve'


# --- ping_ips ---

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
    assert all(r[pc.PingColumn.PING_COUNT] == 2 for r in results)


def test_ping_ips_order_violation(monkeypatch):
    def mock_safe_ping(addresses, **kwargs):
        # Возвращаем хост, но с чужим IP
        h = FakeHost('1.1.1.1')
        return [h], [None]

    monkeypatch.setattr(pc, 'safe_ping_hosts', mock_safe_ping)
    with pytest.raises(RuntimeError, match='Порядок нарушен'):
        pc.ping_ips(('8.8.8.8',), count=1, timeout=1, privileged=False)


def test_ping_ips_hostname_resolve(monkeypatch):
    def mock_safe_ping(addresses, **kwargs):
        h = FakeHost('127.0.0.1')  # localhost резолвится
        return [h], [None]

    monkeypatch.setattr(pc, 'safe_ping_hosts', mock_safe_ping)
    results = pc.ping_ips(('localhost',), count=1, timeout=1, privileged=False)
    assert results[0][pc.PingColumn.IP] == 'localhost'
    assert 'резолвлен в `127.0.0.1`' in results[0][pc.PingColumn.WARNINGS]


# --- Сохранение результатов ---

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


# --- Добавление в источник ---

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

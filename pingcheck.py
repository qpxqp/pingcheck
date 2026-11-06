"""Пинг IP-адресов из Excel-файла с сохранением результатов."""
import argparse
import ipaddress
import logging
import sys
from datetime import datetime
from enum import StrEnum
from pathlib import Path
from typing import Any

import pandas as pd
from icmplib import multiping, ping
from icmplib.exceptions import ICMPLibError


class PingColumn(StrEnum):
    IP = 'ip'
    ALIVE = 'alive'
    SUCCESS_PCT = 'success_pct'
    PING_COUNT = 'ping_count'
    AVG_RTT = 'avg_rtt'
    MIN_RTT = 'min_rtt'
    MAX_RTT = 'max_rtt'
    JITTER = 'jitter'
    PACKET_LOSS = 'packet_loss'
    WARNINGS = 'warnings'


# --- Config ---
IN_USE_MARKER = '====='
BASE_DIR = Path(__file__).parent
RESULTS_DIR_NAME = 'results'
RESULTS_DIR = BASE_DIR / RESULTS_DIR_NAME
TMP_MARKER = '.tmp'
DEFAULT_INPUT_FILE = BASE_DIR / 'ips.xlsx'
DEFAULT_COLUMN_NAME = PingColumn.IP.value
DEFAULT_PING_COUNT = 3
DEFAULT_PING_TIMEOUT = 2
# Форматы timestamp.strftime
FMT_OUTPUT_FILENAME = '%Y-%m-%d_%H-%M-%S'
FMT_COL_NAME = '%Y-%m-%d %H:%M:%S'
FMT_LOGGING = '%Y-%m-%d %H:%M:%S'
# Данные этого столбца будут добавлены в новый столбец в источник
APPEND_DATA = PingColumn.SUCCESS_PCT
DEFAULT_CHUNK_SIZE = 50  # 0 или меньше — без разбивки

LOG_NO_RESULTS = 'Нет результатов для записи'


logger = logging.getLogger(__name__)


def setup_logging(
    log_path: Path,
    log_level: int = logging.INFO,
) -> None:
    """Настраивает логирование: консоль + файл (дописывание)."""
    fmt = logging.Formatter(
        '%(asctime)s [%(levelname)s] %(message)s',
        datefmt=FMT_LOGGING,
    )
    root = logging.getLogger()
    root.setLevel(log_level)
    # консоль
    console = logging.StreamHandler()
    console.setFormatter(fmt)
    root.addHandler(console)
    # файл (дописывание)
    file_handler = logging.FileHandler(
        log_path, mode='a', encoding='utf-8',
    )
    file_handler.setFormatter(fmt)
    root.addHandler(file_handler)


def make_output_filename(timestamp: datetime) -> str:
    """Имя файла без ':' и пробелов."""
    return timestamp.strftime(FMT_OUTPUT_FILENAME) + '.xlsx'


def normalize_ip(ip: str) -> str:
    """Убирает ведущие нули в октетах IPv4: '192.168.001.001'->'192.168.1.1'.

    Возвращает исходную строку, если это не похоже на IPv4.
    """
    parts = ip.strip().split('.')
    if len(parts) != 4:
        return ip
    try:
        return '.'.join(str(int(p)) for p in parts)
    except ValueError:
        return ip


def is_ip_literal(target: str) -> bool:
    """True, если target — IP-литерал (v4 или v6), а не hostname."""
    try:
        ipaddress.ip_address(target)
    except ValueError:
        return False
    return True


def read_ips_from_excel(path: Path, column: str) -> tuple[str, ...]:
    """Читает IP-адреса из указанного столбца Excel-файла.

    Raises:
        FileNotFoundError: файл не существует.
        KeyError: столбец отсутствует.
        ImportError: не установлен движок для чтения .xlsx (openpyxl).
    """
    df = pd.read_excel(path)
    if column not in df.columns:
        raise KeyError(f'Столбец `{column}` не найден в файле `{path}`')
    ips = df[column].dropna().astype(str).str.strip().tolist()
    ips = [ip for ip in ips if ip]  # удаление пустых строк
    logger.info('В файле `%s` найдено %d IP', path, len(ips))
    return tuple(ips)


def chunked(lst: list, size: int):
    """Делит список на чанки указанного размера."""
    for i in range(0, len(lst), size):
        yield lst[i:i + size]


def safe_ping_hosts(
    addresses: list[str],
    *,
    count: int,
    timeout: int,
    privileged: bool,
    chunk_size: int = DEFAULT_CHUNK_SIZE,
) -> tuple[list[Any | None], list[str | None]]:
    """Пингует список. Сначала multiping, при сбое — поштучно."""
    addresses_count = len(addresses)
    if chunk_size <= 0:
        chunk_size = addresses_count or 1
    hosts_out: list[Any | None] = []
    errors_out: list[str | None] = []
    for n, chunk in enumerate(chunked(addresses, chunk_size), start=1):
        logger.debug('Чанк %d из %d адресов', n, len(chunk))
        try:
            hosts = multiping(
                addresses=chunk,
                count=count,
                timeout=timeout,
                privileged=privileged,
            )
            hosts_out.extend(hosts)
            errors_out.extend([None] * len(chunk))
            continue
        except ICMPLibError as e:
            logger.warning(
                'Групповой пинг не выполнен (%s), '
                'пинг поштучно для %d адресов из чанка номер %d',
                e, len(chunk), n,
            )
        for addr in chunk:
            try:
                hosts_out.append(
                    ping(addr, count=count, timeout=timeout,
                         privileged=privileged)
                )
                errors_out.append(None)
            except ICMPLibError as e:
                hosts_out.append(None)
                errors_out.append(str(e))
    return hosts_out, errors_out


def ping_ips(
    ips: tuple[str, ...],
    *,
    count: int = DEFAULT_PING_COUNT,
    timeout: int = DEFAULT_PING_TIMEOUT,
    privileged: bool = True,
    verbose: bool = False,
    chunk_size: int = DEFAULT_CHUNK_SIZE,
) -> tuple[dict, ...]:
    """Пингует IP-адреса и возвращает кортеж словарей."""
    if not ips:
        logger.warning('Список IP пуст — пинговать нечего')
        return ()
    normalized = [normalize_ip(ip) for ip in ips]
    logger.info('Пинг %d адресов...', len(ips))
    hosts, errors = safe_ping_hosts(
        normalized,
        count=count,
        timeout=timeout,
        privileged=privileged,
        chunk_size=chunk_size,
    )
    results: list[dict] = []
    success = 0
    failed = 0
    log = logger.info if verbose else logger.debug
    log('Результаты:')
    for orig, norm, host, err in zip(
        ips, normalized, hosts, errors, strict=True,
    ):
        warnings_list: list[str] = []
        if orig != norm:
            warnings_list.append(
                f'IP нормализован для пинга: `{orig}` -> `{norm}`'
            )
        if err is not None:
            warnings_list.append(f'не пингуется: {err}')
        if host is None:
            results.append({
                PingColumn.IP: orig,
                PingColumn.ALIVE: False,
                PingColumn.SUCCESS_PCT: 0,
                PingColumn.PING_COUNT: 0,
                PingColumn.AVG_RTT: None,
                PingColumn.MIN_RTT: None,
                PingColumn.MAX_RTT: None,
                PingColumn.JITTER: None,
                PingColumn.PACKET_LOSS: None,
                PingColumn.WARNINGS: '; '.join(warnings_list) or None,
            })
            failed += 1
            log('%-20s %s', orig, f'FAIL: {err}')
            continue
        if is_ip_literal(norm) and host.address != norm:
            raise RuntimeError(
                f'Порядок нарушен: ожидали `{norm}`, получили `{host.address}`'
            )
        sent = getattr(host, 'packets_sent', 0)
        recv = getattr(host, 'packets_received', 0)
        pct = round(recv / sent * 100) if sent else 0
        is_alive = recv > 0
        if host.address != norm:
            warnings_list.append(f'`{orig}` резолвлен в `{host.address}`')
        results.append({
            PingColumn.IP: orig,
            PingColumn.ALIVE: is_alive,
            PingColumn.SUCCESS_PCT: pct,
            PingColumn.PING_COUNT: count,
            PingColumn.AVG_RTT: getattr(host, 'avg_rtt', None),
            PingColumn.PACKET_LOSS: getattr(host, 'packet_loss', None),
            PingColumn.JITTER: getattr(host, 'jitter', None),
            PingColumn.MAX_RTT: getattr(host, 'max_rtt', None),
            PingColumn.MIN_RTT: getattr(host, 'min_rtt', None),
            PingColumn.WARNINGS: '; '.join(warnings_list) or None,
        })
        if is_alive:
            success += 1
            is_alive_msg = 'OK'
        else:
            failed += 1
            is_alive_msg = 'FAIL'
        log('%-20s %s', orig, is_alive_msg)
    logger.info('Успешно: %d', success)
    logger.info('Неуспешно: %d', failed)
    return tuple(results)


def save_results_to_excel(path: Path, results: tuple[dict, ...]) -> bool:
    """Сохраняет результаты в Excel. Возвращает True, если что-то сохранено."""
    if not results:
        logger.warning(LOG_NO_RESULTS)
        return False
    path.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(results)
    df = df.reindex(columns=[
        PingColumn.IP, PingColumn.ALIVE, PingColumn.SUCCESS_PCT,
        PingColumn.PING_COUNT, PingColumn.AVG_RTT, PingColumn.MIN_RTT,
        PingColumn.MAX_RTT, PingColumn.JITTER, PingColumn.PACKET_LOSS,
        PingColumn.WARNINGS,
    ])
    # RTT — числа с плавающей точкой; None останется NaN
    for col in (PingColumn.AVG_RTT, PingColumn.MIN_RTT, PingColumn.MAX_RTT,
                PingColumn.JITTER, PingColumn.PACKET_LOSS):
        df[col] = pd.to_numeric(df[col], errors='coerce')
    df.to_excel(path, index=False)
    logger.info('Результаты сохранены в `%s`', path)
    return True


def append_result_to_source(
    path: Path,
    column: str,
    results: tuple[dict, ...],
    timestamp: datetime,
    append_column: str = APPEND_DATA,
) -> None:
    """Добавляет результаты в Excel в новый столбец."""
    col_name = timestamp.strftime(FMT_COL_NAME)
    result_map = {r[PingColumn.IP]: r[append_column] for r in results}
    df = pd.read_excel(path)
    if column not in df.columns:
        raise KeyError(f'Столбец `{column}` не найден в файле `{path}`')
    df[col_name] = df[column].astype(str).str.strip().map(result_map)
    tmp = path.with_name(path.stem + TMP_MARKER + path.suffix)
    try:
        df.to_excel(tmp, index=False)
        tmp.replace(path)
    finally:
        tmp.unlink(missing_ok=True)
    logger.info(
        'В файле `%s` добавлен столбец `%s` с данными `%s`',
        path, col_name, append_column,
    )


def parse_args() -> argparse.Namespace:
    """Разбирает аргументы командной строки."""
    parser = argparse.ArgumentParser(
        description='Пинг IP-адресов из Excel и запись результатов.'
    )
    parser.add_argument(
        '-i',
        '--input-file',
        type=Path,
        default=DEFAULT_INPUT_FILE,
        help=(
            f'Путь к исходному Excel-файлу '
            f'(по умолчанию: `{DEFAULT_INPUT_FILE}`)'
        ),
    )
    parser.add_argument(
        '--column',
        type=str,
        default=DEFAULT_COLUMN_NAME,
        help=(
            f'Имя столбца с IP-адресами '
            f'(по умолчанию: {DEFAULT_COLUMN_NAME})'
        ),
    )
    parser.add_argument(
        '-c',
        '--count',
        type=int,
        default=DEFAULT_PING_COUNT,
        help=f'Количество ICMP-пакетов (по умолчанию: {DEFAULT_PING_COUNT})',
    )
    parser.add_argument(
        '-t',
        '--timeout',
        type=int,
        default=DEFAULT_PING_TIMEOUT,
        help=(
            f'Таймаут одного пакета в секундах '
            f'(по умолчанию: {DEFAULT_PING_TIMEOUT})'
        ),
    )
    parser.add_argument(
        '--no-privileged',
        action='store_false',
        dest='privileged',
        default=True,
        help=(
            'Использовать UDP вместо ICMP '
            '(не требует прав администратора для Linux/macOS)'
        ),
    )
    parser.add_argument(
        '--no-report',
        action='store_false',
        dest='write_report',
        default=True,
        help=f'Не сохранять отчёт в `{RESULTS_DIR_NAME}/...`',
    )
    parser.add_argument(
        '--no-append-source',
        action='store_false',
        dest='append_source',
        default=True,
        help=(
            'Не добавлять колонку `YYYY-MM-DD hh:mm:ss` '
            'с результатами в исходный файл'
        ),
    )
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Показывать результат по каждому IP в логе',
    )
    parser.add_argument(
        '--chunk-size',
        type=int,
        default=DEFAULT_CHUNK_SIZE,
        help=(
            f'Размер чанка для группового пинга адресов '
            f'(по умолчанию: {DEFAULT_CHUNK_SIZE}, 0 — без разбивки)'
        ),
    )
    return parser.parse_args()


def main() -> int:
    """Точка входа. Возвращает код выхода (0 — успех)."""
    args = parse_args()
    input_file = args.input_file
    column = args.column
    timestamp = datetime.now()
    logger.info('%s Process started %s', IN_USE_MARKER, IN_USE_MARKER)
    try:
        ips = read_ips_from_excel(input_file, column)
        results = ping_ips(
            ips,
            count=args.count,
            timeout=args.timeout,
            privileged=args.privileged,
            verbose=args.verbose,
            chunk_size=args.chunk_size,
        )
        if not results:
            logger.warning(LOG_NO_RESULTS)
            return 1
        if args.write_report:
            save_results_to_excel(
                RESULTS_DIR / make_output_filename(timestamp),
                results=results,
            )
        else:
            logger.info(
                'Запись отчёта в `results/...` отключена (--no-report)'
            )
        if args.append_source:
            append_result_to_source(input_file, column, results, timestamp)
        else:
            logger.info(
                'Запись в исходный файл отключена (--no-append-source)'
            )
    except FileNotFoundError:
        logger.error('Файл не найден: `%s`', input_file)
        return 1
    except KeyError as e:
        logger.error('Ошибка структуры файла: %s', e)
        return 1
    except ImportError as e:
        logger.error(
            'Не хватает зависимости для работы с Excel: %s. '
            'Необходимо выполнить: `uv sync`',
            e,
        )
        return 2
    # except ICMPLibError as e:
    #     logger.error('Пинг не выполнен: %s', e)
    #     return 1
    except PermissionError as e:
        logger.error(
            'Недостаточно прав доступа. '
            'Убедитесь, что файл закрыт другими приложениями: `%s`',
            e,
        )
        return 1
    except Exception:
        logger.exception('Непредвиденная ошибка:')
        return 2
    finally:
        elapsed = (datetime.now() - timestamp).total_seconds()
        logger.info(
            '%s Process finished in %.2f sec %s',
            IN_USE_MARKER, elapsed, IN_USE_MARKER,
        )
    return 0


if __name__ == '__main__':
    setup_logging(Path(__file__).with_suffix('.log'))
    sys.exit(main())

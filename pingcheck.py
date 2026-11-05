"""Пинг IP-адресов из Excel-файла с сохранением результатов."""
import argparse
import logging
import sys
from datetime import datetime
from enum import StrEnum
from pathlib import Path

import pandas as pd
from icmplib import multiping
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
RESULTS_DIR = BASE_DIR / 'results'
TMP_MARKER = '.tmp'
DEFAULT_INPUT_FILE = BASE_DIR / 'ips.xlsx'
DEFAULT_COLUMN_NAME = PingColumn.IP.value
DEFAULT_PING_COUNT = 2
DEFAULT_PING_TIMEOUT = 3
# Форматы timestamp.strftime
FMT_OUTPUT_FILENAME = '%Y-%m-%d_%H-%M-%S'
FMT_COL_NAME = '%Y-%m-%d %H:%M:%S'
FMT_LOGGING = '%Y-%m-%d %H:%M:%S'
# Данные этого столбца будут добавлены в новый столбец в источник
APPEND_DATA = PingColumn.SUCCESS_PCT

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


def ping_ips(
    ips: tuple[str, ...],
    *,
    count: int = DEFAULT_PING_COUNT,
    timeout: int = DEFAULT_PING_TIMEOUT,
    privileged: bool = True,
    verbose: bool = False,
) -> tuple[dict, ...]:
    """Пингует IP-адреса и возвращает кортеж словарей."""
    if not ips:
        logger.warning('Список IP пуст — пинговать нечего')
        return ()
    normalized = [normalize_ip(ip) for ip in ips]
    logger.info('Пинг %d адресов...', len(ips))
    hosts = multiping(
        addresses=normalized,
        count=count,
        timeout=timeout,
        privileged=privileged,
    )
    results: list[dict] = []
    success = 0
    failed = 0
    log = logger.info if verbose else logger.debug
    for orig, norm, host in zip(ips, normalized, hosts, strict=True):
        if host.address != norm:
            raise RuntimeError(
                f'Порядок нарушен: ожидали `{norm}`, получили `{host.address}`'
            )
        sent = getattr(host, 'packets_sent', 0)
        recv = getattr(host, 'packets_received', 0)
        pct = round(recv / sent * 100) if sent else 0
        is_alive = recv > 0
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
            PingColumn.WARNINGS: (
                f'IP нормализован для пинга: `{orig}` -> `{norm}`'
                if orig != norm else
                None
            ),
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
        help='Не сохранять отчёт в `results/YYYY-MM-DD_hh-mm-ss.xlsx`',
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
    return parser.parse_args()


def main() -> int:
    """Точка входа. Возвращает код выхода (0 — успех)."""
    args = parse_args()
    input_file = args.input_file
    column = args.column
    count = args.count
    timeout = args.timeout
    privileged = args.privileged
    write_report = args.write_report
    append_source = args.append_source
    timestamp = datetime.now()
    logger.info('%s Process started %s', IN_USE_MARKER, IN_USE_MARKER)
    try:
        ips = read_ips_from_excel(input_file, column)
        results = ping_ips(
            ips,
            count=count,
            timeout=timeout,
            privileged=privileged,
            verbose=args.verbose,
        )
        if not results:
            logger.warning(LOG_NO_RESULTS)
            return 1
        if write_report:
            save_results_to_excel(
                RESULTS_DIR / make_output_filename(timestamp),
                results=results,
            )
        else:
            logger.info(
                'Запись отчёта в `results/...` отключена (--no-report)'
            )
        if append_source:
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
    except ICMPLibError as e:
        logger.error('Пинг не выполнен: %s', e)
        return 1
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
        logger.info('%s Process finished %s', IN_USE_MARKER, IN_USE_MARKER)
    return 0


if __name__ == '__main__':
    setup_logging(Path(__file__).with_suffix('.log'))
    sys.exit(main())

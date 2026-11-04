#!/usr/bin/env python3
"""
Listing Runner: runs the Reverb and Etsy Excel bots for SellerPilot Hub (Auto Bot page).

The server is the master. This laptop:
  1. polls GET <INGEST_BASE_URL>reverb/next-job/ and .../etsy/next-job/ with its ingest token,
  2. downloads that job's Excel,
  3. runs the existing bot script with the Excel path (and --store when set),
  4. when the bot stops, uploads the finished and remaining Excel files and marks the job complete.

While a bot runs, GET .../jobs/<id>/ is checked every few seconds. When the job is
cancelled or deleted on the server, only that bot's Python run and its Chrome window
are stopped. The runner keeps going.

The bots, their stores.json, Chrome profiles and VPN settings are never changed here.

Usage:
  python poller.py              # run forever
  python poller.py --once       # handle at most one job, then exit
"""
from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import urljoin

import requests

HERE = Path(__file__).resolve().parent
ENV_FILE = HERE / 'ingest.env'
LOG_FILE = HERE / 'ingest.log'
JOBS_DIR = HERE / 'jobs'

DEFAULT_POLL_INTERVAL = 20      # seconds between next-job polls when idle
DEFAULT_CANCEL_CHECK_EVERY = 5  # seconds between GET /jobs/<id>/ probes
DEFAULT_TIMEOUT = 60
VENDORS = ('reverb', 'etsy')

# job_type -> (script under the bot folder, output folder under the bot folder, done prefix, remaining prefix)
JOB_TYPES: dict[str, tuple[str, str, str, str]] = {
    'reverb-create': (r'create_listing\create_listing.py', 'output', 'Created', 'Remaining'),
    'reverb-stock': (r'price_stock\update_price_stock.py', r'output\price_stock', 'Updated', 'remain-to-update'),
    'reverb-sku': (r'sku_replace\sku_replace.py', r'output\sku_replace', 'Updated', 'remain-to-update'),
    'etsy-create': (r'create_listing\create_listing.py', 'output', 'Created', 'Remaining'),
    'etsy-stock': (r'price_stock\update_price_stock.py', r'output\price_stock', 'Updated', 'remain-to-update'),
    'etsy-sku': (r'sku_replace\sku_replace.py', r'output\sku_replace', 'Updated', 'remain-to-update'),
}

EXIT_NOTES = {
    0: 'Bot finished.',
    2: 'Bot finished. Some rows failed; they are in the remaining file.',
    3: 'Bot stopped on a login or human check that was not finished in 5 minutes. Chrome was left open.',
    130: 'Bot was stopped by the user on the laptop.',
}


def log(msg: str) -> None:
    ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    line = f'[{ts}] [runner] {msg}'
    print(line, flush=True)
    try:
        with LOG_FILE.open('a', encoding='utf-8') as fh:
            fh.write(line + '\n')
    except Exception:
        pass


def load_env(path: Path) -> dict[str, str]:
    data: dict[str, str] = {}
    if not path.exists():
        return data
    for raw in path.read_text(encoding='utf-8').splitlines():
        line = raw.strip()
        if not line or line.startswith('#') or '=' not in line:
            continue
        key, value = line.split('=', 1)
        data[key.strip()] = value.strip().strip('"').strip("'")
    return data


# --------------------------------------------------------------------------- #
# HTTP (same rules as the Costco/HEB poller)                                  #
# --------------------------------------------------------------------------- #

def _normalize_api_url(api_url: str) -> str:
    """Force https:// on non-loopback hosts. A 301 from http would turn POST into GET."""
    url = (api_url or '').strip()
    if not url.lower().startswith('http://'):
        return url
    rest = url[len('http://'):]
    host = rest.split('/', 1)[0].split(':', 1)[0].lower()
    if host in ('localhost', '127.0.0.1', '::1') or host.endswith('.local'):
        return url
    return 'https://' + rest


def _endpoint(api_url: str, suffix: str) -> str:
    return api_url.rstrip('/') + '/' + suffix.lstrip('/')


class Server:
    def __init__(self, base_url: str, token: str, *, timeout: int, verify: bool):
        self.base = _normalize_api_url(base_url).rstrip('/') + '/'
        self.token = token
        self.timeout = timeout
        self.verify = verify

    def _headers(self) -> dict[str, str]:
        return {'Authorization': f'Bearer {self.token}'}

    def vendor_url(self, vendor: str) -> str:
        return self.base + vendor + '/'

    def _get(self, url: str, **kwargs):
        return requests.get(
            url, headers=self._headers(), timeout=self.timeout, verify=self.verify,
            allow_redirects=False, **kwargs,
        )

    def poll_next_job(self, vendor: str) -> dict[str, Any] | None:
        """A job dict, or None. Never raises, so a server outage never stops the runner."""
        try:
            resp = self._get(_endpoint(self.vendor_url(vendor), 'next-job/'))
        except Exception as exc:
            log(f'{vendor} poll error: {exc}')
            return None
        if resp.status_code in (301, 302, 303, 307, 308):
            log(f'{vendor} poll redirect -> {resp.headers.get("Location", "")!r}. Use https:// in INGEST_BASE_URL.')
            return None
        if resp.status_code != 200:
            log(f'{vendor} poll failed: HTTP {resp.status_code} {resp.text[:200]}')
            return None
        try:
            data = resp.json()
        except ValueError:
            return None
        return data if isinstance(data, dict) and data.get('job_id') else None

    def download(self, file_url: str, dest: Path) -> None:
        url = urljoin(self.base, file_url)
        with self._get(url, stream=True) as resp:
            if resp.status_code != 200:
                raise RuntimeError(f'Excel download failed: HTTP {resp.status_code} {resp.text[:200]}')
            dest.parent.mkdir(parents=True, exist_ok=True)
            with dest.open('wb') as fh:
                for chunk in resp.iter_content(chunk_size=65536):
                    fh.write(chunk)

    def should_stop(self, vendor: str, job_id: str) -> bool:
        """True when the job was cancelled or deleted. Network errors keep the bot running."""
        try:
            resp = self._get(_endpoint(self.vendor_url(vendor), f'jobs/{job_id}/'))
        except Exception as exc:
            log(f'  status check error: {exc}')
            return False
        if resp.status_code == 404:
            log('  job was deleted on the server.')
            return True
        if resp.status_code != 200:
            log(f'  status check failed: HTTP {resp.status_code} {resp.text[:200]}')
            return False
        try:
            return bool(resp.json().get('cancelled'))
        except ValueError:
            return False

    def upload_results(
        self, vendor: str, job_id: str, done: Path | None, remaining: Path | None, exit_code: int | None, note: str,
    ) -> bool:
        files = {}
        handles = []
        try:
            for field, path in (('done_file', done), ('remaining_file', remaining)):
                if path is not None:
                    fh = path.open('rb')
                    handles.append(fh)
                    files[field] = (path.name, fh, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            data = {'job_id': job_id, 'note': note}
            if exit_code is not None:
                data['exit_code'] = str(exit_code)
            resp = requests.post(
                self.vendor_url(vendor), data=data, files=files or None, headers=self._headers(),
                timeout=max(self.timeout, 120), verify=self.verify, allow_redirects=False,
            )
        except Exception as exc:
            log(f'  result upload error: {exc}')
            return False
        finally:
            for fh in handles:
                fh.close()
        if resp.status_code != 200:
            log(f'  result upload failed: HTTP {resp.status_code} {resp.text[:200]}')
            return False
        log(f'  results uploaded: {resp.json().get("stats")}')
        return True

    def complete(self, vendor: str, job_id: str, status: str, note: str) -> None:
        try:
            resp = requests.post(
                _endpoint(self.vendor_url(vendor), f'jobs/{job_id}/complete/'),
                json={'status': status, 'note': note[:2000]}, headers=self._headers(),
                timeout=self.timeout, verify=self.verify, allow_redirects=False,
            )
        except Exception as exc:
            log(f'  complete error: {exc}')
            return
        if resp.status_code == 200:
            log(f'  job {job_id} marked {status}.')
        else:
            log(f'  complete failed: HTTP {resp.status_code} {resp.text[:200]}')


# --------------------------------------------------------------------------- #
# Bot process                                                                  #
# --------------------------------------------------------------------------- #

def _bot_python(bot_dir: Path, configured: str) -> str:
    if configured:
        return configured
    venv = bot_dir / '.venv' / 'Scripts' / 'python.exe'
    return str(venv) if venv.is_file() else sys.executable


def _debug_port(bot_dir: Path) -> int:
    """Chrome debug port from the bot's stores.json. Read only, never written."""
    try:
        cfg = json.loads((bot_dir / 'stores.json').read_text(encoding='utf-8'))
        return int(cfg.get('debug_port') or 9222)
    except Exception:
        return 9222


def stop_bot(proc: subprocess.Popen, port: int) -> None:
    """Stop this bot's Python run and the Chrome it drives. The runner itself stays up."""
    if proc.poll() is None:
        subprocess.run(['taskkill', '/F', '/T', '/PID', str(proc.pid)], capture_output=True, text=True)
    # The bot's Chrome is the only chrome.exe started with this remote-debugging port.
    script = (
        "Get-CimInstance Win32_Process | Where-Object { ($_.Name -eq 'chrome.exe' -or "
        "$_.Name -eq 'chromedriver.exe') -and $_.CommandLine -and "
        f"$_.CommandLine -like '*--remote-debugging-port={port}*' }} | "
        "ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"
    )
    subprocess.run(['powershell', '-NoProfile', '-Command', script], capture_output=True, text=True)
    try:
        proc.wait(timeout=15)
    except subprocess.TimeoutExpired:
        pass


def newest_output(folder: Path, prefix: str, since: float) -> Path | None:
    """The newest <prefix>(YYYY-MM-DD).xlsx (or its -HHMMSS fallback) written after the bot started."""
    pattern = re.compile(rf'^{re.escape(prefix)}\(\d{{4}}-\d{{2}}-\d{{2}}\)(-\d+)?\.xlsx$', re.I)
    found = [
        p for p in folder.glob('*.xlsx')
        if pattern.match(p.name) and p.stat().st_mtime >= since - 2
    ]
    return max(found, key=lambda p: p.stat().st_mtime) if found else None


def _safe_name(name: str) -> str:
    clean = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', '_', Path(name or 'input.xlsx').name).strip(' .')
    return clean if clean.lower().endswith('.xlsx') else (clean or 'input') + '.xlsx'


def handle_job(job: dict[str, Any], server: Server, bot_dirs: dict[str, Path], cfg: dict[str, str]) -> None:
    vendor = job['vendor']
    job_id = job['job_id']
    job_type = job.get('job_type') or ''
    store = (job.get('store_name') or '').strip()
    log(f'Claimed {job_type} job {job_id} ({job.get("file_name")}){f" store {store}" if store else ""}')

    spec = JOB_TYPES.get(job_type)
    bot_dir = bot_dirs.get(vendor)
    if spec is None or bot_dir is None or not bot_dir.is_dir():
        server.complete(vendor, job_id, 'failed', f'This laptop cannot run {job_type}: bot folder not set or missing.')
        return
    script_rel, out_rel, done_prefix, remain_prefix = spec
    script = bot_dir / script_rel
    if not script.is_file():
        server.complete(vendor, job_id, 'failed', f'Bot script not found on {cfg.get("RUNNER_NAME") or "laptop"}: {script}')
        return

    excel = JOBS_DIR / job_id / _safe_name(job.get('file_name') or 'input.xlsx')
    try:
        server.download(job['file_url'], excel)
    except Exception as exc:
        server.complete(vendor, job_id, 'failed', f'Could not download the Excel: {exc}')
        return

    # Arguments stay separate list items: paths contain spaces.
    args = [_bot_python(bot_dir, cfg.get('POLLER_PYTHON', '')), str(script), str(excel)]
    if store:
        args += ['--store', store]
    port = _debug_port(bot_dir)
    cancel_every = max(2, int(cfg.get('CANCEL_CHECK_EVERY') or DEFAULT_CANCEL_CHECK_EVERY))
    started = time.time()
    log(f'  running {script} ...')
    proc = subprocess.Popen(args, cwd=str(script.parent), creationflags=subprocess.CREATE_NEW_CONSOLE)

    try:
        next_check = time.time() + cancel_every
        while proc.poll() is None:
            time.sleep(1)
            if time.time() >= next_check:
                next_check = time.time() + cancel_every
                if server.should_stop(vendor, job_id):
                    log('  stopping the bot (cancelled or deleted on the server).')
                    stop_bot(proc, port)
                    return
    except KeyboardInterrupt:
        log('  runner stopped by user. Stopping the bot.')
        stop_bot(proc, port)
        server.complete(vendor, job_id, 'failed', 'The laptop runner was stopped. Send the file again.')
        raise

    code = proc.returncode
    out_dir = bot_dir / out_rel
    done = newest_output(out_dir, done_prefix, started)
    remaining = newest_output(out_dir, remain_prefix, started)
    note = EXIT_NOTES.get(code, f'Bot exited with code {code}.')
    if done is None and remaining is None:
        note += ' The bot wrote no result files.'
    log(f'  bot exited with code {code}. done={done} remaining={remaining}')
    server.upload_results(vendor, job_id, done, remaining, code, note)
    server.complete(vendor, job_id, 'done' if (done or remaining) else 'failed', note)


# --------------------------------------------------------------------------- #
# Main loop                                                                    #
# --------------------------------------------------------------------------- #

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description='Run Reverb/Etsy Excel jobs from SellerPilot Hub')
    parser.add_argument('--once', action='store_true', help='Handle at most one job, then exit')
    parser.add_argument('--interval', type=int, default=0, help='Seconds between polls when idle')
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    cfg = load_env(ENV_FILE)
    base = cfg.get('INGEST_BASE_URL', '')
    token = cfg.get('INGEST_TOKEN', '')
    if not base or not token:
        log(f'Set INGEST_BASE_URL and INGEST_TOKEN in {ENV_FILE}')
        return 2
    bot_dirs = {
        vendor: Path(cfg[key])
        for vendor, key in (('reverb', 'REVERB_BOT_DIR'), ('etsy', 'ETSY_BOT_DIR'))
        if cfg.get(key)
    }
    vendors = [v for v in VENDORS if v in bot_dirs]
    if not vendors:
        log('Set REVERB_BOT_DIR and/or ETSY_BOT_DIR in ingest.env')
        return 2
    server = Server(
        base, token,
        timeout=int(cfg.get('INGEST_TIMEOUT') or DEFAULT_TIMEOUT),
        verify=(cfg.get('INGEST_VERIFY_SSL', '1') != '0'),
    )
    interval = args.interval or int(cfg.get('POLLER_INTERVAL') or DEFAULT_POLL_INTERVAL)
    log(f'Listing Runner started. Polling {", ".join(vendors)} every {interval}s at {server.base}')

    while True:
        # One job at a time: both bots share Chrome debug port and profiles.
        job = None
        for vendor in vendors:
            job = server.poll_next_job(vendor)
            if job:
                job.setdefault('vendor', vendor)
                break
        if job:
            try:
                handle_job(job, server, bot_dirs, cfg)
            except KeyboardInterrupt:
                raise
            except Exception as exc:
                log(f'job crashed: {exc!r}')
                server.complete(job['vendor'], job['job_id'], 'failed', f'Runner error: {exc!r}'[:500])
            if args.once:
                return 0
            continue
        if args.once:
            return 0
        time.sleep(interval)


def _run_forever() -> int:
    """Restart after an unexpected error instead of dying. Ctrl+C stops it."""
    while True:
        try:
            return main()
        except KeyboardInterrupt:
            log('interrupted - bye')
            return 130
        except SystemExit:
            raise
        except Exception as exc:
            log(f'FATAL: {exc!r} - restarting in 30s')
            try:
                time.sleep(30)
            except KeyboardInterrupt:
                return 130


if __name__ == '__main__':
    sys.exit(_run_forever())

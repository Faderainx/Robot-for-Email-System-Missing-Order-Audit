"""组装 ECOPV 工作台便携包。

PyInstaller 负责把 Python 工作台打成 dist/ECOPV_Workbench；本脚本再把
前端、脱敏配置、参考表和空的持久化目录放到同一目录，并生成 zip。
不复制当前机器的邮箱密码、工单密码、登录态、缓存或历史数据库。
"""

from __future__ import annotations

import argparse
import hashlib
import shutil
import zipfile
from pathlib import Path


SANITIZED_CONFIG = """email:
  imap_server: imap.qiye.aliyun.com
  imap_port: 993
  address: ''
  audit_addresses:
    - report@example.com
  password: ''
  mailbox: INBOX
  cache_enabled: false
workorder:
  url: https://mp.ecopv-epr.com/workbench/main
  username: ''
  password: ''
  timeout: 30
  query_interval_seconds: 1.5
  query_cache_path: storage/query_cache.json
  query_cache_ttl_days: 7
  show_cached_inputs: true
  input_step_delay_seconds: 0.15
  cached_input_hold_seconds: 0.3
  headless: true
  browser_channel: msedge
  headless_fallback_to_headed: true
  login_state_path: storage/login_state.json
  login_wait_seconds: 120
  field_timeout_seconds: 5
  dropdown_timeout_seconds: 5
  query_click_settle_seconds: 0.3
  table_refresh_timeout_seconds: 3
  max_consecutive_query_failures: 3
  agent_aliases: {}
  force_live_query: true
  max_total_seconds: 0
  date_window_months: 1
  selectors: {}
reference_tables:
  internal_emails: data/internal_emails.xlsx
  agent_emails: data/agent_emails.xlsx
  project_names: data/project_names.xlsx
fuzzy_match:
  opencc_config: s2t.json
  high_threshold: 90
  medium_threshold: 80
  low_threshold: 80
workorder_match:
  date_tolerance_days: 60
  customer_threshold: 85
output:
  dir: output
  categorized: true
  missing_template: 漏单清单_{timestamp}.xlsx
  filtered_template: 过滤清单_{timestamp}.xlsx
logging:
  dir: logs
  level: INFO
ocr:
  fallback: true
llm:
  api_key_env: DEEPSEEK_API_KEY
  api_key: ''
  base_url: https://api.deepseek.com
  model: deepseek-flash
  thinking_mode: disabled
  timeout: 30
  max_tokens: 1000
  temperature: 0.1
"""

USAGE = """# ECOPV 邮件复核工作台（便携版）

## 启动

双击 `启动独立工作台.bat`。首次启动后会自动打开浏览器，页面地址是
`http://127.0.0.1:8765`。关闭启动窗口即可停止本机服务。

## 数据与留痕

- `data` 保存内部邮箱、代理和项目名称参考表，可直接替换为新版 `.xlsx`。
- `storage` 保存操作人员修改、历史留痕和 SQLite 数据库。
- `output` 保存处理完成总表、未完成总表、漏单表及复核导出结果。
- `cache` 和 `logs` 由程序自动生成。

便携包不带当前机器的邮箱密码、工单密码、登录态或历史数据。需要抓取新邮件或
查询工单时，请在原有处理端填写账号；操作人员只使用本工作台时无需安装 Python。

## 更新参考表

把新的 `agent_emails.xlsx`、`internal_emails.xlsx` 或 `project_names.xlsx` 放入
`data` 并覆盖同名文件，刷新页面即可沿用新表。不要删除 `storage`，否则会丢失
本机操作留痕。
"""


def _copy_file(src: Path, dest: Path) -> None:
    if not src.is_file():
        raise FileNotFoundError(src)
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dest)


def assemble(source_root: Path, pyinstaller_dist: Path, package_root: Path) -> Path:
    if package_root.exists():
        shutil.rmtree(package_root)
    package_root.mkdir(parents=True)

    # PyInstaller 的 onedir 运行时（exe + _internal）
    for item in pyinstaller_dist.iterdir():
        target = package_root / item.name
        if item.is_dir():
            shutil.copytree(item, target)
        else:
            shutil.copy2(item, target)

    _copy_file(source_root / "workbench.html", package_root / "workbench.html")
    (package_root / "config.yaml").write_text(SANITIZED_CONFIG, encoding="utf-8")
    for name in ("agent_emails.xlsx", "internal_emails.xlsx", "internal_email_cache.xlsx", "project_names.xlsx"):
        _copy_file(source_root / "data" / name, package_root / "data" / name)

    for folder in ("storage", "storage/workbench_history", "output", "output/workbench_history", "cache", "logs"):
        path = package_root / folder
        path.mkdir(parents=True, exist_ok=True)
        (path / ".keep").write_text("", encoding="utf-8")

    (package_root / "启动独立工作台.bat").write_text(
        '@echo off\nchcp 65001 >nul\ncd /d "%~dp0"\nstart "ECOPV 工作台" "%~dp0ECOPV_Workbench.exe"\n',
        encoding="utf-8",
    )
    (package_root / "启动独立工作台-调试.bat").write_text(
        '@echo off\nchcp 65001 >nul\ncd /d "%~dp0"\n"%~dp0ECOPV_Workbench.exe" --no-browser\npause\n',
        encoding="utf-8",
    )
    (package_root / "使用说明.md").write_text(USAGE, encoding="utf-8")
    return package_root


def make_zip(package_root: Path, zip_path: Path) -> tuple[int, str]:
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for path in sorted(package_root.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(package_root.parent))
    digest = hashlib.sha256(zip_path.read_bytes()).hexdigest()
    return zip_path.stat().st_size, digest


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dist", type=Path, required=True, help="PyInstaller dist/ECOPV_Workbench 目录")
    parser.add_argument("--output", type=Path, required=True, help="便携目录")
    parser.add_argument("--zip", dest="zip_path", type=Path, required=True, help="zip 输出路径")
    args = parser.parse_args()

    source_root = Path(__file__).resolve().parent
    package_root = assemble(source_root, args.dist.resolve(), args.output.resolve())
    size, digest = make_zip(package_root, args.zip_path.resolve())
    required = [
        package_root / "ECOPV_Workbench.exe",
        package_root / "_internal",
        package_root / "workbench.html",
        package_root / "config.yaml",
        package_root / "data" / "project_names.xlsx",
        package_root / "启动独立工作台.bat",
    ]
    missing = [str(path) for path in required if not path.exists()]
    if missing:
        raise SystemExit("便携包校验失败，缺少: " + ", ".join(missing))
    print(f"package={package_root}")
    print(f"zip={args.zip_path.resolve()}")
    print(f"zip_bytes={size}")
    print(f"sha256={digest}")


if __name__ == "__main__":
    main()

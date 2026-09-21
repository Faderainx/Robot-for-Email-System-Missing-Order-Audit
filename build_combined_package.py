"""组装“服务端 Python GUI + Edge 独立工作台”合并交付包。

服务端源码保留现有 GUI/邮件处理能力，但不复制 live 配置、缓存、登录态、
历史数据库和输出；工作台使用已冻结的 ECOPV_Workbench 运行时。
"""

from __future__ import annotations

import argparse
import hashlib
import shutil
import zipfile
from pathlib import Path

from build_portable_workbench import SANITIZED_CONFIG


SERVER_README = """# ECOPV 邮件审核平台：服务端 GUI + Edge 工作台

## 组成

- `服务端源码/`：Python/PyQt GUI、邮件抓取、解析、LLM、工单查询和数据同步。
- `独立工作台/`：无需 Python 的工作台运行时，供 Edge 浏览器访问。

## 服务端启动

服务端电脑需要 Python 3.10+ 及 `requirements.txt` 中的依赖。进入
`服务端源码` 后运行 `启动.bat`，即可打开实际 PyQt GUI。

## Edge 操作端启动

推荐在同一台服务端电脑启动 `独立工作台/ECOPV_Workbench.exe`，然后让操作人员
在 Edge 打开 `http://服务器IP:8765/`。局域网部署可使用：

```text
ECOPV_Workbench.exe --host 0.0.0.0 --public-host 服务器IP
```

服务器防火墙需放行 TCP 8765。默认数据、附件、SQLite 和历史记录都留在服务器。

## 配置与安全

包内 `config.yaml` 已脱敏。首次抓取邮件或查询工单前，在服务端配置账号/API key；
不要把 `storage`、`cache`、`output`、`logs` 或 `session_state.json` 复制到其他电脑。
操作人员只需 Edge，不需要在电脑上安装 Python 或 PyQt。
"""


def _copy_file(src: Path, dest: Path) -> None:
    if not src.is_file():
        raise FileNotFoundError(src)
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dest)


def _source_ignore(directory: str, names: list[str]) -> set[str]:
    ignored = {
        ".git", "__pycache__", "cache", "storage", "output", "logs", "deliveries",
        "dist", "work", "spec", ".pytest_cache", "test_output", "config.yaml",
        "session_state.json", "_runlogs", "node_modules", "findings.md", "progress.md",
        "task_plan.md", "AGENTS.md", "PROJECT.md", "test_offline.py", "test_modules.py",
        "probe_one_query.py", "probe_one_query.bat", "probe_workorder.py", "probe_workorder.bat",
        "build_portable_workbench.py", "build_combined_package.py", "portable_workbench.py",
    }
    return {
        name for name in names
        if name in ignored or name.endswith(".pyc") or name.startswith("test_")
    }


def assemble(source_root: Path, workbench_package: Path, package_root: Path) -> None:
    if package_root.exists():
        shutil.rmtree(package_root)
    package_root.mkdir(parents=True)

    server = package_root / "服务端源码"
    shutil.copytree(source_root, server, ignore=_source_ignore)
    (server / "config.yaml").write_text(SANITIZED_CONFIG, encoding="utf-8")
    # 只携带参考表；不携带运行期间生成的缓存/历史数据。
    data = server / "data"
    if data.exists():
        shutil.rmtree(data)
    for name in ("agent_emails.xlsx", "internal_emails.xlsx", "internal_email_cache.xlsx", "project_names.xlsx"):
        _copy_file(source_root / "data" / name, data / name)
    for folder in ("storage", "output", "cache", "logs"):
        (server / folder).mkdir(parents=True, exist_ok=True)
        (server / folder / ".keep").write_text("", encoding="utf-8")

    workbench = package_root / "独立工作台"
    shutil.copytree(workbench_package, workbench)
    (package_root / "使用说明.md").write_text(SERVER_README, encoding="utf-8")
    (package_root / "启动Edge工作台-局域网.bat").write_text(
        '@echo off\nchcp 65001 >nul\ncd /d "%~dp0独立工作台"\nset /p HOST=请输入服务器局域网IP（例如 192.168.1.20）：\nECOPV_Workbench.exe --host 0.0.0.0 --public-host %HOST%\n',
        encoding="utf-8",
    )


def zip_package(package_root: Path, zip_path: Path) -> tuple[int, str]:
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
    parser.add_argument("--workbench", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--zip", dest="zip_path", type=Path, required=True)
    args = parser.parse_args()
    source_root = Path(__file__).resolve().parent
    assemble(source_root, args.workbench.resolve(), args.output.resolve())
    size, digest = zip_package(args.output.resolve(), args.zip_path.resolve())
    required = [
        args.output / "服务端源码" / "启动.bat",
        args.output / "服务端源码" / "gui.py",
        args.output / "服务端源码" / "config.yaml",
        args.output / "独立工作台" / "ECOPV_Workbench.exe",
        args.output / "独立工作台" / "workbench.html",
    ]
    missing = [str(path) for path in required if not path.exists()]
    if missing:
        raise SystemExit("合并包校验失败，缺少: " + ", ".join(missing))
    print(f"package={args.output.resolve()}")
    print(f"zip={args.zip_path.resolve()}")
    print(f"zip_bytes={size}")
    print(f"sha256={digest}")


if __name__ == "__main__":
    main()

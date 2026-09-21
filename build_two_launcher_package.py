"""组装面向操作人员的双启动 ECOPV 发布包。

发布包根目录只保留两个启动按钮：

* ``启动GUI.bat``：启动已编译的 PyQt GUI；
* ``启动独立工作台.bat``：启动已编译的 Edge 工作台服务。

两个可执行文件都通过 ``ECOPV_APP_ROOT`` 指向同一个 ``其他`` 目录，因此配置、
参考表、历史留痕、缓存和导出结果不会因为分别启动两个程序而分裂到不同位置。
"""

from __future__ import annotations

import argparse
import hashlib
import shutil
import zipfile
from pathlib import Path

from build_portable_workbench import SANITIZED_CONFIG


USAGE = """# ECOPV 邮件审核平台（双启动版）

## 启动

根目录只有两个启动文件：

- `启动GUI.bat`：启动邮件抓取、解析和处理 GUI；
- `启动独立工作台.bat`：启动 Edge 浏览器工作台。

请先把整个文件夹完整解压，不要单独移动或删除 `其他` 文件夹。操作人员电脑不需要
安装 Python、PyQt 或其他编程环境；独立工作台需要安装 Microsoft Edge（或系统默认
浏览器）。压缩包已内置 RAR 解析用的 7-Zip 和图片 OCR（Tesseract），不需要另行安装。

## 配置与数据

- `其他/config.yaml` 是脱敏配置。首次抓取邮件、调用 LLM 或查询工单前，请在 GUI
  的配置位置填写邮箱账号、密码和 API Key；发布包不会携带原电脑的密码或登录态。
- `其他/data` 保存内部邮箱、代理和项目名称参考表，可以替换同名 `.xlsx` 后刷新数据。
- `其他/storage`、`其他/output`、`其他/cache`、`其他/logs` 保存数据库、历史留痕、
  导出表格、缓存和日志。不要删除这些目录。

## 共享数据

GUI 和独立工作台都使用 `其他` 作为同一个数据根目录，操作记录和导出结果会在两边
保持一致。若需要给其他电脑使用，请复制整个解压后的文件夹，不要只复制某个 exe。
"""


def _copy_tree_contents(source: Path, target: Path) -> None:
    if not source.is_dir():
        raise FileNotFoundError(f"找不到构建目录: {source}")
    target.mkdir(parents=True, exist_ok=True)
    for item in source.iterdir():
        destination = target / item.name
        if item.is_dir():
            shutil.copytree(item, destination, dirs_exist_ok=True)
        else:
            shutil.copy2(item, destination)


def _copy_file(source: Path, target: Path) -> None:
    if not source.is_file():
        raise FileNotFoundError(f"找不到文件: {source}")
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, target)


def _copy_portable_seven_zip(shared_root: Path) -> None:
    """把 RAR 解析所需的最小 7-Zip 运行文件放进发布包。"""
    candidates = [
        Path(r"C:\Program Files\7-Zip"),
        Path(r"C:\Program Files (x86)\7-Zip"),
    ]
    source_dir = next((path for path in candidates if (path / "7z.exe").is_file()), None)
    if source_dir is None:
        raise FileNotFoundError("找不到本机 7-Zip 安装目录，无法打包 RAR 解析组件")
    target_dir = shared_root / "tools" / "7z"
    target_dir.mkdir(parents=True, exist_ok=True)
    for name in ("7z.exe", "7z.dll", "License.txt", "readme.txt"):
        source = source_dir / name
        if source.is_file():
            shutil.copy2(source, target_dir / name)
    if not (target_dir / "7z.exe").is_file() or not (target_dir / "7z.dll").is_file():
        raise FileNotFoundError("7-Zip 运行文件不完整，至少需要 7z.exe 和 7z.dll")


def _copy_portable_tesseract(source_root: Path, shared_root: Path) -> None:
    """复制图片 OCR 的轻量 Tesseract 运行时和中英文模型。"""
    source_dir = source_root / "tools" / "tesseract"
    required = (
        source_dir / "tesseract.exe",
        source_dir / "tessdata" / "chi_sim.traineddata",
        source_dir / "tessdata" / "eng.traineddata",
        source_dir / "tessdata" / "osd.traineddata",
    )
    missing = [str(path) for path in required if not path.is_file()]
    if missing:
        raise FileNotFoundError("Tesseract OCR 资源不完整: " + ", ".join(missing))
    target_dir = shared_root / "tools" / "tesseract"
    shutil.copytree(source_dir, target_dir, dirs_exist_ok=True)


def _write_shared_root(source_root: Path, shared_root: Path) -> None:
    shared_root.mkdir(parents=True, exist_ok=True)
    (shared_root / "config.yaml").write_text(SANITIZED_CONFIG, encoding="utf-8")
    _copy_file(source_root / "workbench.html", shared_root / "workbench.html")
    for name in (
        "agent_emails.xlsx",
        "internal_emails.xlsx",
        "internal_email_cache.xlsx",
        "project_names.xlsx",
    ):
        _copy_file(source_root / "data" / name, shared_root / "data" / name)
    for folder in (
        "storage",
        "storage/workbench_history",
        "output",
        "output/workbench_history",
        "cache",
        "logs",
    ):
        directory = shared_root / folder
        directory.mkdir(parents=True, exist_ok=True)
        (directory / ".keep").write_text("", encoding="utf-8")
    _copy_portable_seven_zip(shared_root)
    _copy_portable_tesseract(source_root, shared_root)
    (shared_root / "使用说明.md").write_text(USAGE, encoding="utf-8")


def _write_launchers(package_root: Path) -> None:
    (package_root / "启动GUI.bat").write_text(
        "@echo off\n"
        "chcp 65001 >nul\n"
        "setlocal EnableExtensions\n"
        "set \"ECOPV_APP_ROOT=%~dp0其他\"\n"
        "set \"PATH=%~dp0其他\\tools\\7z;%PATH%\"\n"
        "set \"ECOPV_LAUNCH_LOG=%~dp0其他\\logs\\launcher_gui.log\"\n"
        ">>\"%ECOPV_LAUNCH_LOG%\" echo [%date% %time%] 启动GUI\n"
        "if not exist \"%~dp0其他\\GUI\\ECOPV_GUI.exe\" (>>\"%ECOPV_LAUNCH_LOG%\" echo [ERROR] 找不到 GUI 程序 & exit /b 2)\n"
        "start \"ECOPV 邮件审核 GUI\" \"%~dp0其他\\GUI\\ECOPV_GUI.exe\"\n",
        encoding="utf-8",
    )
    (package_root / "启动独立工作台.bat").write_text(
        "@echo off\n"
        "chcp 65001 >nul\n"
        "setlocal EnableExtensions\n"
        "set \"ECOPV_APP_ROOT=%~dp0其他\"\n"
        "set \"ECOPV_LAUNCH_LOG=%~dp0其他\\logs\\launcher_workbench.log\"\n"
        ">>\"%ECOPV_LAUNCH_LOG%\" echo [%date% %time%] 启动独立工作台\n"
        "if not exist \"%~dp0其他\\工作台\\ECOPV_Workbench.exe\" (>>\"%ECOPV_LAUNCH_LOG%\" echo [ERROR] 找不到工作台程序 & exit /b 2)\n"
        "start \"ECOPV 独立工作台\" \"%~dp0其他\\工作台\\ECOPV_Workbench.exe\"\n",
        encoding="utf-8",
    )


def assemble(
    source_root: Path,
    gui_dist: Path,
    workbench_package: Path,
    package_root: Path,
) -> Path:
    if package_root.exists():
        shutil.rmtree(package_root)
    package_root.mkdir(parents=True, exist_ok=True)

    shared_root = package_root / "其他"
    _write_shared_root(source_root, shared_root)

    gui_target = shared_root / "GUI"
    _copy_file(gui_dist / "ECOPV_GUI.exe", gui_target / "ECOPV_GUI.exe")
    gui_internal = gui_dist / "_internal"
    if gui_internal.is_dir():
        shutil.copytree(gui_internal, gui_target / "_internal", dirs_exist_ok=True)
    else:
        raise FileNotFoundError(f"GUI 构建目录缺少 _internal: {gui_dist}")
    # 构建命令为开发机调试方便把配置/参考表也塞进了 PyInstaller 运行时；
    # 正式包统一从共享根目录读取，必须删掉其中可能含有旧凭据的副本。
    for bundled_name in ("config.yaml", "workbench.html"):
        bundled_path = gui_target / "_internal" / bundled_name
        if bundled_path.exists():
            bundled_path.unlink()
    bundled_data = gui_target / "_internal" / "data"
    if bundled_data.exists():
        shutil.rmtree(bundled_data)
    # 工单自动化默认使用操作人员电脑已有的 Microsoft Edge；Playwright 自带的
    # Chromium 和 headless shell 合计约 800 MB，正式 Edge 版不再复制它们。
    bundled_browsers = (
        gui_target
        / "_internal"
        / "playwright"
        / "driver"
        / "package"
        / ".local-browsers"
    )
    if bundled_browsers.exists():
        shutil.rmtree(bundled_browsers)
    if not (gui_target / "ECOPV_GUI.exe").is_file():
        raise FileNotFoundError(f"GUI 构建目录缺少 ECOPV_GUI.exe: {gui_dist}")

    workbench_target = shared_root / "工作台"
    _copy_file(workbench_package / "ECOPV_Workbench.exe", workbench_target / "ECOPV_Workbench.exe")
    _copy_file(workbench_package / "workbench.html", shared_root / "workbench.html")
    internal = workbench_package / "_internal"
    if internal.is_dir():
        shutil.copytree(internal, workbench_target / "_internal", dirs_exist_ok=True)
    else:
        raise FileNotFoundError(f"工作台构建目录缺少 _internal: {workbench_package}")

    _write_launchers(package_root)
    return package_root


def _scan_for_secrets(package_root: Path) -> list[str]:
    forbidden = (
        "GNkEjPkny5ocGja",
        "ba8e4ENj&H2",
        "huiyan.song@eu-helper.com",
    )
    hits: list[str] = []
    for path in package_root.rglob("*"):
        if not path.is_file() or path.suffix.lower() in {".exe", ".dll", ".pyd", ".zip", ".xlsx"}:
            continue
        try:
            content = path.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue
        if any(token in content for token in forbidden):
            hits.append(str(path))
    return hits


def make_zip(package_root: Path, zip_path: Path) -> tuple[int, str]:
    if zip_path.exists():
        zip_path.unlink()
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for path in sorted(package_root.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(package_root.parent))
    digest = hashlib.sha256(zip_path.read_bytes()).hexdigest()
    return zip_path.stat().st_size, digest


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--gui-dist", type=Path, required=True, help="PyInstaller GUI dist/ECOPV_GUI 目录")
    parser.add_argument("--workbench-package", type=Path, required=True, help="已构建的工作台目录")
    parser.add_argument("--output", type=Path, required=True, help="最终解压目录")
    parser.add_argument("--zip", dest="zip_path", type=Path, required=True, help="最终 zip 路径")
    args = parser.parse_args()

    source_root = Path(__file__).resolve().parent
    package_root = assemble(
        source_root,
        args.gui_dist.resolve(),
        args.workbench_package.resolve(),
        args.output.resolve(),
    )
    required = (
        package_root / "启动GUI.bat",
        package_root / "启动独立工作台.bat",
        package_root / "其他" / "GUI" / "ECOPV_GUI.exe",
        package_root / "其他" / "工作台" / "ECOPV_Workbench.exe",
        package_root / "其他" / "config.yaml",
        package_root / "其他" / "data" / "project_names.xlsx",
        package_root / "其他" / "workbench.html",
        package_root / "其他" / "tools" / "7z" / "7z.exe",
        package_root / "其他" / "tools" / "7z" / "7z.dll",
        package_root / "其他" / "tools" / "tesseract" / "tesseract.exe",
        package_root / "其他" / "tools" / "tesseract" / "tessdata" / "chi_sim.traineddata",
        package_root / "其他" / "tools" / "tesseract" / "tessdata" / "eng.traineddata",
    )
    missing = [str(path) for path in required if not path.exists()]
    if missing:
        raise SystemExit("双启动包校验失败，缺少: " + ", ".join(missing))
    root_items = sorted(path.name for path in package_root.iterdir())
    if set(root_items) != {"启动GUI.bat", "启动独立工作台.bat", "其他"}:
        raise SystemExit(f"双启动包根目录包含非启动项: {root_items}")
    secrets = _scan_for_secrets(package_root)
    if secrets:
        raise SystemExit("双启动包疑似包含敏感配置: " + ", ".join(secrets))

    size, digest = make_zip(package_root, args.zip_path.resolve())
    print(f"package={package_root}")
    print(f"zip={args.zip_path.resolve()}")
    print(f"zip_bytes={size}")
    print(f"sha256={digest}")
    print(f"root_items={root_items}")


if __name__ == "__main__":
    main()

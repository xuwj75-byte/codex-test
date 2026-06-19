#!/usr/bin/env python3
"""
本地文件整理工具（Windows 友好版）

安全原则：
- 只处理用户明确指定的 source 文件夹。
- 默认 dry-run（预演）模式，只生成计划，不移动/复制文件。
- 不删除文件，不修改文件内容，不覆盖同名文件。
"""

from __future__ import annotations

import argparse
import csv
import logging
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, Set, Tuple


TYPE_FOLDERS: Dict[str, Set[str]] = {
    "Word文档": {".doc", ".docx", ".dot", ".dotx", ".rtf"},
    "PDF文件": {".pdf"},
    "Excel表格": {".xls", ".xlsx", ".xlsm", ".xlt", ".xltx", ".csv"},
    "PPT演示": {".ppt", ".pptx", ".pps", ".ppsx"},
    "图片": {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".heic", ".svg"},
    "视频": {".mp4", ".mov", ".avi", ".mkv", ".wmv", ".flv", ".mpeg", ".mpg", ".webm"},
    "音频": {".mp3", ".wav", ".flac", ".aac", ".m4a", ".ogg", ".wma"},
    "压缩包": {".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz"},
    "文本文件": {".txt", ".md", ".log", ".ini", ".json", ".xml", ".yaml", ".yml"},
}
OTHER_TYPE = "其他文件"
PLAN_FILE_NAME = "organize_plan.csv"
LOG_FILE_NAME = "organize_log.txt"


@dataclass(frozen=True)
class PlanItem:
    source_path: Path
    target_path: Path
    file_type: str
    file_size: int
    modified_time: datetime
    has_name_conflict: bool


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="安全整理指定文件夹中的文件：默认只预演，添加 --execute 才会移动或复制。"
    )
    parser.add_argument("--source", required=True, help="要整理的源文件夹路径，例如 C:\\Users\\me\\Downloads")
    parser.add_argument("--target", required=True, help="整理后的目标文件夹路径，例如 D:\\整理后的文件")
    parser.add_argument("--execute", action="store_true", help="正式执行。未添加此参数时只生成整理计划。")
    parser.add_argument("--copy", action="store_true", help="复制文件而不是移动文件。必须与 --execute 一起才会真正复制。")
    parser.add_argument("--plan", default=PLAN_FILE_NAME, help=f"整理计划 CSV 文件路径，默认 {PLAN_FILE_NAME}")
    parser.add_argument("--log", default=LOG_FILE_NAME, help=f"日志文件路径，默认 {LOG_FILE_NAME}")
    return parser.parse_args()


def setup_logging(log_path: Path) -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )


def validate_folder(path_text: str, label: str) -> Path:
    path = Path(path_text).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"{label}不存在：{path}")
    if not path.is_dir():
        raise NotADirectoryError(f"{label}不是文件夹：{path}")
    return path


def get_file_type(path: Path) -> str:
    suffix = path.suffix.lower()
    for folder_name, extensions in TYPE_FOLDERS.items():
        if suffix in extensions:
            return folder_name
    return OTHER_TYPE


def iter_source_files(source: Path, target: Path) -> Iterable[Path]:
    """只递归扫描用户指定的 source 文件夹，跳过 target 内文件避免自我整理。"""
    for path in source.rglob("*"):
        if not path.is_file():
            continue
        resolved = path.resolve()
        if is_relative_to(resolved, target):
            logging.info("跳过目标文件夹内的文件：%s", resolved)
            continue
        yield resolved


def is_relative_to(path: Path, possible_parent: Path) -> bool:
    try:
        path.relative_to(possible_parent)
        return True
    except ValueError:
        return False


def unique_target_path(desired_path: Path, reserved_paths: Set[Path]) -> Tuple[Path, bool]:
    """返回不会覆盖已有文件、也不会与本次计划中其他文件冲突的目标路径。"""
    conflict = desired_path.exists() or desired_path in reserved_paths
    if not conflict:
        reserved_paths.add(desired_path)
        return desired_path, False

    stem = desired_path.stem
    suffix = desired_path.suffix
    parent = desired_path.parent
    counter = 1
    while True:
        candidate = parent / f"{stem}_{counter:03d}{suffix}"
        if not candidate.exists() and candidate not in reserved_paths:
            reserved_paths.add(candidate)
            return candidate, True
        counter += 1


def build_plan(source: Path, target: Path) -> list[PlanItem]:
    plan: list[PlanItem] = []
    reserved_paths: Set[Path] = set()

    for file_path in iter_source_files(source, target):
        stat = file_path.stat()
        modified_time = datetime.fromtimestamp(stat.st_mtime)
        file_type = get_file_type(file_path)
        desired_target = target / file_type / str(modified_time.year) / file_path.name
        final_target, has_conflict = unique_target_path(desired_target, reserved_paths)
        plan.append(
            PlanItem(
                source_path=file_path,
                target_path=final_target,
                file_type=file_type,
                file_size=stat.st_size,
                modified_time=modified_time,
                has_name_conflict=has_conflict,
            )
        )
    return plan


def write_plan_csv(plan: list[PlanItem], plan_path: Path) -> None:
    with plan_path.open("w", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(["原文件路径", "目标文件路径", "文件类型", "文件大小", "最后修改时间", "是否会重名"])
        for item in plan:
            writer.writerow(
                [
                    str(item.source_path),
                    str(item.target_path),
                    item.file_type,
                    item.file_size,
                    item.modified_time.strftime("%Y-%m-%d %H:%M:%S"),
                    "是" if item.has_name_conflict else "否",
                ]
            )


def execute_plan(plan: list[PlanItem], copy_mode: bool) -> None:
    action_name = "复制" if copy_mode else "移动"
    for item in plan:
        item.target_path.parent.mkdir(parents=True, exist_ok=True)
        logging.info("%s：%s -> %s", action_name, item.source_path, item.target_path)
        if copy_mode:
            shutil.copy2(item.source_path, item.target_path)
        else:
            shutil.move(str(item.source_path), str(item.target_path))


def main() -> int:
    args = parse_args()
    log_path = Path(args.log).expanduser().resolve()
    setup_logging(log_path)

    try:
        source = validate_folder(args.source, "源文件夹")
        target = Path(args.target).expanduser().resolve()

        logging.info("源文件夹：%s", source)
        logging.info("目标文件夹：%s", target)
        logging.info("模式：%s%s", "正式执行" if args.execute else "预演 dry-run", "（复制）" if args.copy else "（移动）")

        plan = build_plan(source, target)
        plan_path = Path(args.plan).expanduser().resolve()
        write_plan_csv(plan, plan_path)
        logging.info("已生成整理计划：%s，共 %s 个文件", plan_path, len(plan))

        if args.execute:
            target.mkdir(parents=True, exist_ok=True)
            execute_plan(plan, args.copy)
            logging.info("正式执行完成。")
        else:
            logging.info("当前是预演 dry-run，未移动、未复制、未删除任何文件。确认计划后添加 --execute 再正式执行。")

        return 0
    except Exception as exc:  # noqa: BLE001 - CLI 需要把错误写入日志并给新手可读提示
        logging.error("发生错误：%s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

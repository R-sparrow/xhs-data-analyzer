from pathlib import Path
from typing import Any

import pandas as pd


NOTE_LIST_COLUMNS = [
    "笔记标题",
    "首次发布时间",
    "体裁",
    "曝光",
    "观看量",
    "封面点击率",
    "点赞",
    "评论",
    "收藏",
    "涨粉",
    "分享",
    "人均观看时长",
]

DAILY_SHEETS = {
    "impressions": ("曝光数（天）", ["日期", "曝光数"]),
    "views": ("观看数（天）", ["日期", "观看数"]),
    "cover_click_rate": ("封面点击率（天）", ["日期", "封面点击率"]),
    "avg_watch_time": ("平均观看时长（天）", ["日期", "人均观看时长"]),
    "followers_gained": ("涨粉数（天）", ["日期", "涨粉数"]),
}


def _as_existing_file(file_path: Any) -> Path:
    path = Path(file_path)
    if not path.is_file():
        raise FileNotFoundError(f"Excel 文件不存在：{path}")
    return path


def _require_columns(
    dataframe: pd.DataFrame, required_columns: list[str], source_name: str
) -> None:
    missing_columns = [
        column for column in required_columns if column not in dataframe.columns
    ]
    if missing_columns:
        raise ValueError(
            f"{source_name} 缺少必要字段: {', '.join(missing_columns)}"
        )


def read_note_list_excel(file_path: Any) -> pd.DataFrame:
    """读取笔记列表明细表，并返回 V1 所需的原始字段。"""
    path = _as_existing_file(file_path)
    dataframe = pd.read_excel(
        path,
        sheet_name="Sheet1",
        header=1,
        engine="openpyxl",
    )

    _require_columns(dataframe, NOTE_LIST_COLUMNS, f"{path.name} / Sheet1")
    return dataframe.loc[:, NOTE_LIST_COLUMNS].copy()


def read_note_daily_metrics(file_path: Any) -> dict[str, pd.DataFrame]:
    """读取单篇笔记数据明细表中的 V1 日级指标 Sheet。"""
    path = _as_existing_file(file_path)
    required_sheet_names = [sheet_name for sheet_name, _ in DAILY_SHEETS.values()]

    with pd.ExcelFile(path, engine="openpyxl") as workbook:
        missing_sheets = [
            sheet_name
            for sheet_name in required_sheet_names
            if sheet_name not in workbook.sheet_names
        ]

    if missing_sheets:
        raise ValueError(
            f"{path.name} 缺少必需 Sheet：{', '.join(missing_sheets)}"
        )

    result = {}
    for key, (sheet_name, required_columns) in DAILY_SHEETS.items():
        dataframe = pd.read_excel(
            path,
            sheet_name=sheet_name,
            header=0,
            engine="openpyxl",
        )
        _require_columns(
            dataframe,
            required_columns,
            f"{path.name} / {sheet_name}",
        )
        result[key] = dataframe.loc[:, required_columns].copy()

    return result

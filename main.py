from pathlib import Path

from excel_reader import read_note_daily_metrics, read_note_list_excel


def main() -> None:
    project_dir = Path(__file__).resolve().parent
    note_list_path = project_dir / "data" / "笔记列表明细表.xlsx"
    note_detail_path = project_dir / "data" / "弃医之后的生活（74）-数据明细表.xlsx"

    note_list = read_note_list_excel(note_list_path)
    print("第一类 Excel 读取成功")
    print(f"行数：{note_list.shape[0]}，列数：{note_list.shape[1]}")
    print(f"列名：{list(note_list.columns)}")

    daily_metrics = read_note_daily_metrics(note_detail_path)
    print("第二类 Excel 读取成功")
    print(f"返回的 key：{list(daily_metrics.keys())}")
    for key, dataframe in daily_metrics.items():
        print(f"{key}: shape={dataframe.shape}")


if __name__ == "__main__":
    main()

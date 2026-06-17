import pandas as pd
from excel_exporter import ExcelExporter, export_to_excel


def example_using_class():
    print("=== 示例 1: 使用 ExcelExporter 类 ===")

    exporter = ExcelExporter("output/class_example.xlsx")

    users_data = [
        {"id": 1, "name": "张三", "age": 28, "department": "技术部"},
        {"id": 2, "name": "李四", "age": 32, "department": "市场部"},
        {"id": 3, "name": "王五", "age": 25, "department": "财务部"},
        {"id": 4, "name": "赵六", "age": 35, "department": "技术部"},
    ]
    exporter.add_sheet(users_data, "用户信息")

    orders_data = pd.DataFrame({
        "订单号": ["ORD001", "ORD002", "ORD003", "ORD004", "ORD005"],
        "用户ID": [1, 2, 1, 3, 4],
        "金额": [199.00, 350.50, 89.90, 1299.00, 520.00],
        "状态": ["已完成", "待发货", "已完成", "已取消", "已完成"],
        "下单时间": ["2024-01-15", "2024-01-16", "2024-01-17", "2024-01-18", "2024-01-19"],
    })
    exporter.add_sheet(orders_data, "订单数据")

    products_data = pd.DataFrame({
        "产品ID": [101, 102, 103, 104],
        "产品名称": ["笔记本电脑", "无线鼠标", "机械键盘", "显示器"],
        "价格": [5999, 129, 499, 1299],
        "库存": [50, 200, 80, 30],
    })
    exporter.add_sheet(products_data, "产品列表")

    file_path = exporter.export()
    print(f"Excel 文件已导出至: {file_path}")


def example_using_function():
    print("\n=== 示例 2: 使用 export_to_excel 函数 ===")

    sales_summary = {
        "月份": ["1月", "2月", "3月", "4月", "5月", "6月"],
        "销售额(万元)": [120, 135, 150, 128, 160, 175],
        "目标完成率": ["92%", "102%", "108%", "95%", "112%", "118%"],
    }

    inventory = [
        {"仓库": "北京仓", "商品种类": 230, "库存总量": 5000, "预警库存": 300},
        {"仓库": "上海仓", "商品种类": 280, "库存总量": 6500, "预警库存": 400},
        {"仓库": "广州仓", "商品种类": 190, "库存总量": 4200, "预警库存": 250},
        {"仓库": "成都仓", "商品种类": 150, "库存总量": 3800, "预警库存": 200},
    ]

    data_sheets = [
        {"name": "销售汇总", "data": sales_summary},
        {"name": "库存统计", "data": inventory, "header_style": True},
    ]

    file_path = export_to_excel(data_sheets, "output/function_example.xlsx")
    print(f"Excel 文件已导出至: {file_path}")


def example_custom_options():
    print("\n=== 示例 3: 自定义选项 ===")

    exporter = ExcelExporter("output/custom_example.xlsx")

    data1 = [{"A": 1, "B": 2, "C": 3}, {"A": 4, "B": 5, "C": 6}]
    exporter.add_sheet(data1, "默认样式")

    data2 = [{"X": 10, "Y": 20, "Z": 30}, {"X": 40, "Y": 50, "Z": 60}]
    exporter.add_sheet(data2, "无样式", header_style=False, auto_width=False, freeze_panes=None)

    data3 = pd.DataFrame({f"列{i}": range(5) for i in range(1, 11)})
    exporter.add_sheet(data3, "宽表格测试", freeze_panes="B2")

    file_path = exporter.export()
    print(f"Excel 文件已导出至: {file_path}")


def example_sheet_name_edge_cases():
    print("\n=== 示例 4: Sheet 名称边界情况处理 ===")

    exporter = ExcelExporter("output/edge_cases_example.xlsx")

    long_name = "这是一个非常非常非常非常非常非常非常长的Sheet名称超过31个字符"
    exporter.add_sheet([{"数据": "长名称测试"}], long_name)

    bad_chars_name = "销售/报告\\2024*汇总?表:Q1[机密]"
    exporter.add_sheet([{"数据": "非法字符测试"}], bad_chars_name)

    exporter.add_sheet([{"数据": 1}], "月度报表")
    exporter.add_sheet([{"数据": 2}], "月度报表")
    exporter.add_sheet([{"数据": 3}], "月度报表")

    exporter.add_sheet([{"数据": 4}], "")
    exporter.add_sheet([{"数据": 5}], None)

    print("Sheet 名称处理结果:")
    for i, sheet in enumerate(exporter.sheets, 1):
        print(f"  Sheet {i}: '{sheet['name']}' (长度={len(sheet['name'])})")

    file_path = exporter.export()
    print(f"Excel 文件已导出至: {file_path}")


def example_sheet_reordering():
    print("\n=== 示例 5: Sheet 顺序调整功能 ===")

    exporter = ExcelExporter("output/reorder_example.xlsx")

    exporter.add_sheet([{"id": 1, "name": "张三"}], "用户信息")
    exporter.add_sheet([{"订单号": "ORD001", "金额": 199}], "订单数据")
    exporter.add_sheet([{"产品ID": 101, "名称": "笔记本"}], "产品列表")
    exporter.add_sheet([{"月份": "1月", "销售额": 120}], "销售汇总")
    exporter.add_sheet([{"仓库": "北京仓", "库存": 5000}], "库存统计")

    print("初始顺序:", exporter.get_sheet_names())

    exporter.swap_sheets("用户信息", "销售汇总")
    print("交换'用户信息'和'销售汇总'后:", exporter.get_sheet_names())

    exporter.move_sheet(4, 0)
    print("将索引4的Sheet移到最前:", exporter.get_sheet_names())

    exporter.move_sheet("订单数据", 4)
    print("将'订单数据'移到最后:", exporter.get_sheet_names())

    exporter.reorder_sheets([0, 1, "产品列表", 3, "订单数据"])
    print("重新排列顺序后:", exporter.get_sheet_names())

    file_path = exporter.export()
    print(f"Excel 文件已导出至: {file_path}")

    print("\n最终 Excel 中的 Sheet 顺序:")
    xl = pd.ExcelFile(file_path)
    for i, name in enumerate(xl.sheet_names, 1):
        print(f"  {i}. {name}")
    xl.close()


if __name__ == "__main__":
    import os
    os.makedirs("output", exist_ok=True)

    example_using_class()
    example_using_function()
    example_custom_options()
    example_sheet_name_edge_cases()
    example_sheet_reordering()

    print("\n✅ 所有示例执行完成！")

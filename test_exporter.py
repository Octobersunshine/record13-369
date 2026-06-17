import os
import tempfile
import pandas as pd
import pytest
from excel_exporter import ExcelExporter, export_to_excel


class TestExcelExporter:
    def setup_method(self):
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_output.xlsx")

    def teardown_method(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
        os.rmdir(self.temp_dir)

    def test_add_dataframe_sheet(self):
        exporter = ExcelExporter(self.test_file)
        df = pd.DataFrame({"A": [1, 2, 3], "B": [4, 5, 6]})
        exporter.add_sheet(df, "Sheet1")
        assert len(exporter.sheets) == 1
        assert exporter.sheets[0]["name"] == "Sheet1"

    def test_add_list_of_dicts_sheet(self):
        exporter = ExcelExporter(self.test_file)
        data = [{"name": "张三", "age": 28}, {"name": "李四", "age": 32}]
        exporter.add_sheet(data, "用户")
        assert len(exporter.sheets) == 1
        assert isinstance(exporter.sheets[0]["data"], pd.DataFrame)
        assert list(exporter.sheets[0]["data"].columns) == ["name", "age"]

    def test_add_single_dict_sheet(self):
        exporter = ExcelExporter(self.test_file)
        data = {"name": "张三", "age": 28}
        exporter.add_sheet(data, "单个用户")
        assert len(exporter.sheets[0]["data"]) == 1

    def test_export_single_sheet(self):
        exporter = ExcelExporter(self.test_file)
        data = [{"id": 1, "value": "test"}]
        exporter.add_sheet(data, "TestSheet")
        result = exporter.export()
        assert os.path.exists(result)
        assert result.endswith("test_output.xlsx")

    def test_export_multiple_sheets(self):
        exporter = ExcelExporter(self.test_file)
        exporter.add_sheet([{"A": 1}], "Sheet1")
        exporter.add_sheet([{"B": 2}], "Sheet2")
        exporter.add_sheet([{"C": 3}], "Sheet3")

        result = exporter.export()
        assert os.path.exists(result)

        with pd.ExcelFile(result) as xl:
            assert xl.sheet_names == ["Sheet1", "Sheet2", "Sheet3"]

    def test_export_no_sheets_raises_error(self):
        exporter = ExcelExporter(self.test_file)
        with pytest.raises(ValueError, match="No sheets to export"):
            exporter.export()

    def test_invalid_data_type_raises_error(self):
        exporter = ExcelExporter(self.test_file)
        with pytest.raises(ValueError, match="data must be a DataFrame"):
            exporter.add_sheet("invalid data", "Sheet1")

    def test_export_to_excel_function(self):
        sheets = [
            {"name": "数据1", "data": [{"X": 1, "Y": 2}]},
            {"name": "数据2", "data": [{"A": 10, "B": 20}]},
        ]
        result = export_to_excel(sheets, self.test_file)
        assert os.path.exists(result)

        with pd.ExcelFile(result) as xl:
            assert xl.sheet_names == ["数据1", "数据2"]

        df1 = pd.read_excel(result, sheet_name="数据1")
        assert list(df1.columns) == ["X", "Y"]
        assert df1.iloc[0]["X"] == 1

    def test_export_without_header_style(self):
        exporter = ExcelExporter(self.test_file)
        data = [{"col1": 1, "col2": 2}]
        exporter.add_sheet(data, "NoStyle", header_style=False)
        result = exporter.export()
        assert os.path.exists(result)

    def test_export_without_auto_width(self):
        exporter = ExcelExporter(self.test_file)
        data = [{"col1": 1, "col2": 2}]
        exporter.add_sheet(data, "NoAutoWidth", auto_width=False)
        result = exporter.export()
        assert os.path.exists(result)

    def test_export_without_freeze_panes(self):
        exporter = ExcelExporter(self.test_file)
        data = [{"col1": 1, "col2": 2}]
        exporter.add_sheet(data, "NoFreeze", freeze_panes=None)
        result = exporter.export()
        assert os.path.exists(result)

    def test_export_empty_dataframe(self):
        exporter = ExcelExporter(self.test_file)
        df = pd.DataFrame(columns=["col1", "col2", "col3"])
        exporter.add_sheet(df, "EmptySheet")
        result = exporter.export()
        assert os.path.exists(result)

        df_read = pd.read_excel(result, sheet_name="EmptySheet")
        assert list(df_read.columns) == ["col1", "col2", "col3"]
        assert len(df_read) == 0

    def test_sheet_names_are_preserved(self):
        exporter = ExcelExporter(self.test_file)
        exporter.add_sheet([{"a": 1}], "用户信息表")
        exporter.add_sheet([{"b": 2}], "订单明细表")
        exporter.add_sheet([{"c": 3}], "产品分类表")

        result = exporter.export()
        with pd.ExcelFile(result) as xl:
            assert xl.sheet_names == ["用户信息表", "订单明细表", "产品分类表"]

    def test_data_content_accuracy(self):
        exporter = ExcelExporter(self.test_file)
        original_data = [
            {"id": 1, "name": "张三", "age": 28},
            {"id": 2, "name": "李四", "age": 32},
            {"id": 3, "name": "王五", "age": 25},
        ]
        exporter.add_sheet(original_data, "UserData")
        result = exporter.export()

        df_read = pd.read_excel(result, sheet_name="UserData")
        assert len(df_read) == 3
        assert df_read.iloc[0]["name"] == "张三"
        assert df_read.iloc[1]["age"] == 32
        assert list(df_read["id"]) == [1, 2, 3]


if __name__ == "__main__":
    pytest.main([__file__, "-v"])

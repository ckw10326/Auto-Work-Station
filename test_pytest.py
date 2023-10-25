'''
pytest -s test_text_gen.py
pytest -s test_file_process.py
pytest 測試的檔案清單
pytest --collect-only
'''
import os
import itertools
from text_gen import copy_plan_file, text_gen
from function_file_process import files_list, check_plan


def test_text_gen(monkeypatch):
    """模擬測試程式"""
    inputs = iter(["2", "3"])  # 使用迭代器模擬多次的輸入
    inputs = itertools.cycle(inputs)  # 將迭代器循環使用
    monkeypatch.setattr('builtins.input', lambda _: next(inputs))  # 模擬輸入值
    letter_title = 'HRSG Chimney-General Arrangement of Concrete Roof & Layout of Permanent Shutter to Roof Slab'
    letter_rev = "A"
    letter_num = 'TPC-TC(C0)-CD-23-0002'
    letter_date = '2023/02/02'

    contents0_expected = "本文係中部施工處對統包商提送「HRSG Chimney-General Arrangement of Concrete Roof & Layout of Permanent Shutter to Roof Slab」Rev.A所提審查意見(共3頁)，未逾合約規範，已電傳泰興公司，擬陳閱後文存。"
    contents1_expected = "檢送台中電廠燃氣機組更新改建計畫「HRSG Chimney-General Arrangement of Concrete Roof & Layout of Permanent Shutter to Roof Slab」，Rev.A，中部施工處之審查意見（如附，共3頁）供卓參，請查照。"
    contents2_expected = "依據GE/CTCI 112年02月02日TPC-TC(C0)-CD-23-0002號辦理。"
    contents3_expected = "本文係統包商提送「HRSG Chimney-General Arrangement of Concrete Roof & Layout of Permanent Shutter to Roof Slab」Rev.A，本組無意見，已Email通知泰興公司，擬陳閱後文存。"
    result = text_gen(letter_title, letter_rev, letter_num, letter_date)

    assert result[0] == contents0_expected
    assert result[1] == contents1_expected
    assert result[2] == contents2_expected
    assert result[3] == contents3_expected

def test_files_list():
    """測試"""
    # 測試情境1：沒有指定關鍵字
    xpath = "/path/to/directory"
    expected_output = [
        "/path/to/directory/file1.txt",
        "/path/to/directory/subdirectory/file2.txt"
    ]
    result = files_list(xpath)
    assert result.sort() == expected_output.sort()

    # 測試情境2：指定關鍵字
    xpath = "/path/to/directory"
    search_str = "file2"
    expected_output = [
        "/path/to/directory/subdirectory/file2.txt"
    ]
    result = files_list(xpath, search_str)
    assert result.sort() == expected_output.sort()

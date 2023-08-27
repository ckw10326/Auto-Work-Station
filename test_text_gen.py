'''
pytest -s test_text_gen.py
pytest -s test_file_process.py
pytest 測試的檔案清單
pytest --collect-only
'''
import os
import itertools
from text_gen import copy_plan_file, text_gen

def test_copy_plan_file(tmpdir):
    """
    創建測試檔案
    模擬輸出輸入
    """
    # 創建測試檔案
    test_file = tmpdir.join("HT-_test_file.txt")
    test_file.write("Test content")

    # 執行被測試的函數
    print(test_file)
    copy_plan_file(test_file)

    # 檢查是否生成了正確的套印檔案和傳真檔案
    file_folder = os.path.split(test_file)[0]
    file_name = os.path.splitext(os.path.split(test_file)[1])[0]
    print_dest_file = os.path.join(file_folder, f"套印_{file_name}.rtf")
    print(print_dest_file)
    fax_dest_file = os.path.join(file_folder, f"傳真_{file_name}.doc")
    print(print_dest_file)
    assert os.path.exists(print_dest_file)
    assert os.path.exists(fax_dest_file)

    # 清理測試生成的檔案
    os.remove(print_dest_file)
    os.remove(fax_dest_file)

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

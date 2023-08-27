'''
pytest -s test_text_gen.py
'''
import os
from text_gen import copy_plan_file, text_gen

def test_copy_plan_file(tmpdir):
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


def test_text_gen():
    letter_title = 'HRSG Chimney-General Arrangement of Concrete Roof & Layout of Permanent Shutter to Roof Slab'
    letter_rev = "A"
    letter_num = 'TPC-TC(C0)-CD-23-0002'
    letter_date = '2023/02/02'
    expected_output = None
    result = text_gen(letter_title, letter_rev, letter_num, letter_date)
    assert result == expected_output
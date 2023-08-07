import unittest
from text_gen import text_gen
from text_gen import file_path_process
from text_gen import copy_plan_file

class TestTextGen(unittest.TestCase):
    def test_text_gen_with_valid_input(self):
        # 設定測試的輸入值
        letter_title = "HRSG Chimney-General Arrangement of Concrete Roof & Layout of Permanent Shutter to Roof Slab"
        letter_rev = "A"
        letter_num = "TPC-TC(C0)-CD-23-0002"
        letter_date = "2023/02/02"

        # 呼叫 `text_gen` 函數
        actual_text = text_gen(letter_title, letter_rev, letter_num, letter_date)

        # 檢查 `text_gen` 函數的輸出
        self.assertIsNone(actual_text)

    def test_text_gen_with_invalid_input(self):
        # 設定測試的輸入值
        letter_title = ""
        letter_rev = ""
        letter_num = ""
        letter_date = ""

        # 呼叫 `text_gen` 函數
        with self.assertRaises(Exception):
            text_gen(letter_title, letter_rev, letter_num, letter_date)

class TestFilePathProcess(unittest.TestCase):
    def test_file_path_process_with_valid_input(self):
        # 設定測試的輸入值
        file_plan_no = "1"
        doc_end_num = "1234"
        path_way = "1"

        # 呼叫 `file_path_process` 函數
        actual_file_doc_num, actual_file_source_folder, actual_file_dest_folder = file_path_process(file_plan_no, doc_end_num, path_way)

        # 檢查 `file_path_process` 函數的輸出
        expected_file_doc_num = "HT-D1-CTC-GEL-23-1234"
        expected_file_source_folder = constants.SOURCE_FOLDER + "/" + expected_file_doc_num
        expected_file_dest_folder = constants.DEST_FOLDER + "/" + constants.PLAN_LIST[int(file_plan_no)] + "_" + expected_file_doc_num
        self.assertEqual(actual_file_doc_num, expected_file_doc_num)
        self.assertEqual(actual_file_source_folder, expected_file_source_folder)
        self.assertEqual(actual_file_dest_folder, expected_file_dest_folder)

    def test_file_path_process_with_invalid_input(self):
        # 設定測試的輸入值
        file_plan_no = "99"
        doc_end_num = "1234"
        path_way = "1"

        # 呼叫 `file_path_process` 函數
        with self.assertRaises(Exception):
            file_path_process(file_plan_no, doc_end_num, path_way)

    def test_file_path_process_with_invalid_doc_end_num(self):
        # 設定測試的輸入值
        file_plan_no = "1"
        doc_end_num = "1234567890"
        path_way = "1"

        # 呼叫 `file_path_process` 函數
        with self.assertRaises(Exception):
            file_path_process(file_plan_no, doc_end_num, path_way)


if __name__ == "__main__":
    unittest.main()

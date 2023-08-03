import csv
import pandas as pd


class DocItem:
    def __init__(self, drawing_no_value, drawing_title_value,
                 drawing_vision_value, letter_num_value, letter_date_value,
                 letter_titl_value, file_path):
        self.drawing_no_value = drawing_no_value  # 圖號
        self.drawing_title_value = drawing_title_value  # 圖名
        self.drawing_vision_value = drawing_vision_value  # 版次
        self.letter_num_value = letter_num_value  # 來文號碼
        self.letter_date_value = letter_date_value  # 來文日期
        self.letter_titl_value = letter_titl_value  # 來文名稱
        self.file_path = file_path  # 路徑

    # 寫入csv檔案
    def process_data0(self):
        # 執行資料處理
        row0 = [
            self.drawing_no_value,
            self.drawing_title_value,
            self.drawing_vision_value,
            self.letter_num_value,
            self.letter_date_value,
            self.letter_titl_value,
            self.file_path
        ]
        # 將處理後的資料轉換為DataFrame
        df = pd.DataFrame(row0)
        print(df)
        # 將DataFrame儲存為CSV檔案
        csv_path = "/workspaces/Auto-Work-Station/00dest/processed_data.csv"
        df.to_csv(csv_path, index=False)
        print("資料已成功處理並儲存為 processed_data.csv 檔案。")

    @staticmethod
    def load_data0():
        # 從CSV檔案中讀取資料並返回DataFrame
        df = pd.read_csv('processed_data.csv')
        return df

    def process_data(self):
        # 在這裡進行資料處理的操作
        # 您可以使用 self 屬性來訪問 DocItem 的屬性值並進行處理

        # 範例程式碼只是將屬性值印出來作為示範
        print("Drawing No:", self.drawing_no_value)
        print("Drawing Title:", self.drawing_title_value)
        print("Drawing Vision:", self.drawing_vision_value)
        print("Letter Num:", self.letter_num_value)
        print("Letter Date:", self.letter_date_value)
        print("Letter Title:", self.letter_titl_value)
        print("File Path:", self.file_path)
        # 範例程式碼將屬性值寫入 CSV 檔案
        row = [
            self.drawing_no_value,
            self.drawing_title_value,
            self.drawing_vision_value,
            self.letter_num_value,
            self.letter_date_value,
            self.letter_titl_value,
            self.file_path
        ]

        try:
            path = "/workspaces/Auto-Work-Station/00dest/data.csv"
            df = pd.read_csv(path)
        except FileNotFoundError:
            df = pd.DataFrame(columns=['Drawing No', 'Drawing Title', 'Drawing Vision',
                              'Letter Num', 'Letter Date', 'Letter Title', 'File Path'])

        with open(path, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(row)

        print("資料已寫入 data.csv")

    def load_data():
        # 在這裡實作從資料源（例如 CSV 檔案）讀取資料的邏輯
        # 範例程式碼只是返回一個固定的 DocItem 物件作為示範
        return DocItem("123", "Sample Drawing", "A", "001", "2021-01-01", "Sample Letter", "/path/to/file.csv")

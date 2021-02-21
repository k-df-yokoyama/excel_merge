import openpyxl


class RiskList:
    def __init__(self):
        self.file_name = None
        self.sheet_name = None
        self.min_row = 1
        self.min_col = 1
        self.wb = None
        self.ws = None
        self.header_cells = None
        self.risk_list = []
        self.risk_idx_list = []
        self.risk_key = '学籍番号'

    def load_from_excel_book(self, file_name, sheet_name, min_row, min_col):
        # print(file_name)
        self.min_row = min_row
        self.min_col = min_col
        self.file_name = file_name
        self.sheet_name = sheet_name
        self.wb = openpyxl.load_workbook(file_name)
        self.ws = self.wb[sheet_name]

        self.get_risk_list()

    def get_risk_list(self):
        # self.header_cells = None

        for row in self.ws.iter_rows(min_row=self.min_row, min_col=self.min_col):
            if row[0].row == self.min_row:
                self.header_cells = row
            else:
                row_dic = {}
                for k, v in zip(self.header_cells, row):
                    row_dic[k.value] = v.value
                self.risk_list.append(row_dic)

    def get_risk_idx_list(self):
        self.risk_idx_list = [d.get(self.risk_key) for d in self.risk_list]
        # print(risk_idx_list)
        return self.risk_idx_list

    def get_risk_data_dict(self, risk_idx):
        for d in self.risk_list:
            if d.get(self.risk_key) == risk_idx:
                return d

    def get_risk_data_dict_row_idx(self, risk_idx):
        i = 0
        for d in self.risk_list:
            i += 1
            if d.get(self.risk_key) == risk_idx:
                return i

    def set_risk_data(self, risk_idx, col_key, value):
        target_row_idx = risk_idx + self.min_row  # タイトル行があるので -1 は不要
        target_col_idx = self.get_risk_data_dict_col_idx(col_key) + self.min_col - 1
        self.ws.cell(target_row_idx, target_col_idx).value = value

    def save_excel_book(self):
        self.wb.save(self.file_name)

    def get_risk_data_dict_col_idx(self, col_key):
        # self.header_cells
        idx = 0
        for d in self.header_cells:
            idx += 1
            if d.value == col_key:
                return idx
        return -1

    def print_risk_list(self):
        print(self.risk_list)

    def print_rows(self):
        for row in self.ws.iter_rows(min_row=self.min_row, min_col=self.min_col):
            i = 0
            value = ""
            for col in row:
                i += 1
                if i == 1:
                    value += col.value
                else:
                    value = value + "," + col.value
            print(value)


if __name__ == '__main__':

    risk_list = RiskList()
    risk_list.load_from_excel_book("./school_members2.xlsx", "Sheet1", 2, 2)

    risk_list.print_risk_list()
    risk_idx_list = risk_list.get_risk_idx_list()
    print(risk_idx_list)

    risk_key_value = "026"
    risk_data_dict = risk_list.get_risk_data_dict(risk_key_value)
    print("risk_data is " + str(risk_data_dict))
    if risk_data_dict is not None:
        row_idx = risk_list.get_risk_data_dict_row_idx(risk_key_value)
        print(" row_idx is " + str(row_idx))
        print(" name is " + risk_data_dict.get('名前'))

    risk_key_value = "017"
    risk_data_dict = risk_list.get_risk_data_dict(risk_key_value)
    print("risk_data is " + str(risk_data_dict))
    if risk_data_dict is not None:
        row_idx = risk_list.get_risk_data_dict_row_idx(risk_key_value)
        print(" row_idx is " + str(row_idx))
        print(" name is " + risk_data_dict.get('名前'))

    risk_key_value = "000"
    risk_data_dict = risk_list.get_risk_data_dict(risk_key_value)
    print("risk_data is " + str(risk_data_dict))
    if risk_data_dict is not None:
        row_idx = risk_list.get_risk_data_dict_row_idx(risk_key_value)
        print(" row_idx is " + str(row_idx))
        print(" name is " + risk_data_dict.get('名前'))

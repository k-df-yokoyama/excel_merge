from RiskList import RiskList


def update_risk_data_by_idx_and_key(_dict_of_risk_data_of_merge_from, _dict_of_risk_data_of_merge_to, targett_primary_key_value, target_col_key):
    if dict_of_risk_data_of_merge_from.get(target_col_key) != _dict_of_risk_data_of_merge_to.get(target_col_key):
        # 値が異なる場合には、セルの行番号と列番号を取得する
        row_idx_from = risk_list_merge_from.get_row_index_from_primary_key_value(targett_primary_key_value)
        col_idx_from = risk_list_merge_to.get_col_idx_from_col_key(target_col_key)
        print("From:row=" + str(row_idx_from) + ", col=" + str(col_idx_from))

        row_idx_to = risk_list_merge_to.get_row_index_from_primary_key_value(targett_primary_key_value)
        col_idx_to = risk_list_merge_to.get_col_idx_from_col_key(target_col_key)
        print("To:  row=" + str(row_idx_to) + ", col=" + str(col_idx_to))

        # risk_data_from と risk_data_to で値の異なるカラムを更新する
        risk_list_merge_to.set_risk_data(row_idx_to, target_col_key, _dict_of_risk_data_of_merge_from.get(target_col_key))


if __name__ == '__main__':
    min_row = 2
    min_col = 2

    risk_list_merge_from = RiskList()
    risk_list_merge_from.load_from_excel_book("./school_members2.xlsx", "Sheet1", min_row, min_col)

    risk_list_merge_to = RiskList()
    risk_list_merge_to.load_from_excel_book("./school_members1.xlsx", "Sheet1", min_row, min_col)

    # fromのRiskListから、主キーの値のリストを取得
    primary_key_value_list = risk_list_merge_from.get_primary_key_value_list()

    for current_primary_key_value in primary_key_value_list:
        # 指定された主キーの値(current_primary_key_value)に対応する行のリスクデータを辞書型で取得
        dict_of_risk_data_of_merge_from = risk_list_merge_from.get_dictionary_of_one_risk_data(current_primary_key_value)
        print(dict_of_risk_data_of_merge_from)
        
        # 指定された主キーの値(current_primary_key_value)に対応する行のリスクデータを辞書型で取得
        dict_of_risk_data_of_merge_to = risk_list_merge_to.get_dictionary_of_one_risk_data(current_primary_key_value)
        print(dict_of_risk_data_of_merge_to)

        # dict_of_risk_data_of_merge_from と dict_of_risk_data_of_merge_to で、
        # target_col_key で指定したカラムの値を更新する
        update_risk_data_by_idx_and_key(dict_of_risk_data_of_merge_from, dict_of_risk_data_of_merge_to, current_primary_key_value, '名前')
        update_risk_data_by_idx_and_key(dict_of_risk_data_of_merge_from, dict_of_risk_data_of_merge_to, current_primary_key_value, 'クラス')

        risk_list_merge_to.save_excel_book()

from RiskList import RiskList


def update_risk_data_by_idx_and_key(_risk_data_dict_from, _risk_data_dict_to, target_risk_idx, target_col_key):
    if _risk_data_dict_from.get(target_col_key) != _risk_data_dict_to.get(target_col_key):
        # 値が異なる場合には、セルの行番号と列番号を取得する
        row_idx_from = risk_list_merge_from.get_risk_data_dict_row_idx(target_risk_idx)
        col_idx_from = risk_list_merge_to.get_risk_data_dict_col_idx(target_col_key)
        print("From:row=" + str(row_idx_from) + ", col=" + str(col_idx_from))

        row_idx_to = risk_list_merge_to.get_risk_data_dict_row_idx(target_risk_idx)
        col_idx_to = risk_list_merge_to.get_risk_data_dict_col_idx(target_col_key)
        print("To:  row=" + str(row_idx_to) + ", col=" + str(col_idx_to))

        # risk_data_from と risk_data_to で値の異なるカラムを更新する
        risk_list_merge_to.set_risk_data(row_idx_to, target_col_key, _risk_data_dict_from.get(target_col_key))


if __name__ == '__main__':
    min_row = 2
    min_col = 2

    risk_list_merge_to = RiskList()
    risk_list_merge_to.load_from_excel_book("./school_members1.xlsx", "Sheet1", min_row, min_col)

    risk_list_merge_from = RiskList()
    risk_list_merge_from.load_from_excel_book("./school_members2.xlsx", "Sheet1", min_row, min_col)

    risk_idx_list = risk_list_merge_from.get_risk_idx_list()

    for risk_idx in risk_idx_list:
        # risk_idx を持つ行を取得する
        risk_data_dict_from = risk_list_merge_from.get_risk_data_dict(risk_idx)
        print(risk_data_dict_from)
        risk_data_dict_to = risk_list_merge_to.get_risk_data_dict(risk_idx)
        print(risk_data_dict_to)

        # risk_data_from と risk_data_to で、target_col_key で指定したカラムの値が異なるかを調べる
        update_risk_data_by_idx_and_key(risk_data_dict_from, risk_data_dict_to, risk_idx, '名前')
        update_risk_data_by_idx_and_key(risk_data_dict_from, risk_data_dict_to, risk_idx, 'クラス')

        risk_list_merge_to.save_excel_book()

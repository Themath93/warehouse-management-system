

wb_sht_list = ['SHIPMENT_INFORMATION', 'LOCAL_LIST', 'IR_ORDER']

abbreviation=['si','lc','br','br_d','lc','svmx','else','sc','tool']

#약어모음집 dict
abb_sheets_names = dict(zip(wb_sht_list,abbreviation))

#key_value 반대로
reverse_abb_sheets_names = dict(map(reversed,abb_sheets_names.items()))


#####################################################################################
#####################################################################################
def return_dict(dict_name=1):
    """
    약자가 values면 1을 , 약자가 key값인 경우 0 입력
    """
    if dict_name == 1:
        tmp_dict = abb_sheets_names
    elif dict_name == 0:
        tmp_dict = reverse_abb_sheets_names
    else:
        #잘못된 0,1 외의 값이 들어올경우
        pass

    return tmp_dict

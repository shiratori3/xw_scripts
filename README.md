# 一些xlwings脚本

## 当前脚本实现功能

1. 函数：

    - 单元格MD5计算（MDF、MDF_SINGLE）

2. marco（宏）：

    - merge：合并工作表
    - to_csv：批量转换为csv文件
    - clean_num：数值清洗
    - change_sign：变更正负号

### 函数功能介绍

- 单元格MD5计算（MDF、MDF_SINGLE）

> __注意：由于转换器的设置，整数数值会自动变成带.0的浮点数，会导致MD5数值与原值有差异。使用时请留意是否包含整数以免出现计算差异__
>
> 当前支持：
>
> 1. MDF_SINGLE：将整个区域的单元格的值按顺序合并起来计算其MD5数值，并返还单个MD5数值。
> 2. MDF：计算整个区域中的各个单元格值的MD5值，并返回对应的MD5数组。

### 宏功能介绍

- merge：合并工作表

> 当前支持：
>
> 1. 合并当前工作簿的所有工作表
> 2. 合并当前工作簿的文件路径下指定格式（支持csv、xls、xlsx）的所有文件内的首个或多个工作表至当前工作表或对应工作表（以文件名命名或以工作表命名）

    # 现有宏名称及对应功能
    merge_cur_wb_to_cur_sheet_NoRow1 # 合并当前工作簿所有工作表除首行外的数据至当前工作表
    merge_cur_wb_to_first_sheet_NoRow1 # 合并当前工作簿所有工作表除首行外的数据至Sheet1或首个工作表
    merge_csv_to_cur_sheet # 合并当前路径下的所有csv文件的首个工作表至当前工作表
    merge_csv_to_diff_sheets # 合并当前路径下的所有csv文件的首个工作表至对应工作表（以对应文件名命名）
    merge_xls_first_sheet_to_cur_sheet_NoRow1 # 合并当前路径下的所有xls文件首个工作表除首行外的数据至当前工作表
    merge_xlsx_first_sheet_to_cur_sheet_NoRow1 # 合并当前路径下的所有xlsx文件首个工作表除首行外的数据至当前工作表
    merge_xlsx_first_sheet_to_diff_sheets_with_bookname # 合并当前路径下的所有xlsx文件首个工作表除首行外的数据至对应工作表（以对应文件名命名）
    merge_xlsx_first_sheet_to_diff_sheets_with_sheetname # 合并当前路径下的所有xlsx文件首个工作表除首行外的数据至对应工作表（以对应工作表命名）

- to_csv：批量转换为csv文件

> 当前支持：
>
> 1. 将当前工作簿的文件路径下对应格式的文件均转换成对应名称的csv文件

    # 现有宏名称及对应功能
    from_xlsx_to_csv # 当前路径下的xlsx文件批量转成csv
    from_xls_to_csv # 当前路径下的xls文件批量转成csv

- clean_num：数值清洗

> 当前支持：
>
> 1. 将当前工作表中，选中区域中含数字[0-9]的非数值单元格清洗成纯数值，数值单元格仅做格式变换。

    # 现有宏名称及对应功能
    clean_num_with2dec_round # 清洗为2位小数的数值，多于2位部分四舍五入
    clean_num_with2dec_truncate # 清洗为2位小数的数值，多于2位部分截断（向下取整）
    clean_num_with4dec_round # 清洗为4位小数的数值，多于4位部分四舍五入
    clean_num_with0dec_truncate # 清洗为0位小数的数值，多于0位部分截断（向下取整）

- change_sign：变更正负号

> 当前支持：
>
> 1. 将当前工作表中，选中区域中为数值格式的单元格的数值乘以-1，实现变更正负的功能。

    # 现有宏名称及对应功能
    change_selected_sign # 变更选中区域的数值的正负号
    change_selected_sign_preformat_text # 先把区域格式变为数值再变更选中区域的数值的正负号

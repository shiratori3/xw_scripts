# 一些xlwings脚本

## 当前脚本实现功能

1. 函数：

    - 单元格MD5计算（udf_md5_single、udf_md5）
    - 分配（udf_assign、udf_assign_multi）
    - 抽样（udf_sample、udf_sample_assign）

2. 宏：

    - merge：合并工作表
    - to_csv：批量转换为csv文件
    - clean_num：数值清洗
    - change_sign：变更正负号

### 函数功能介绍

- 单元格MD5计算（udf_md5_single、udf_md5）

> __注意：由于转换器的设置，整数数值会自动变成带.0的浮点数，会导致MD5数值与原值有差异。使用时请留意是否包含整数以免出现计算差异__
>
> 当前支持：
>
> 1. udf_md5_single：将整个区域的单元格的值按顺序合并起来计算并返回其MD5数值。
>    - 参数：
>       - data: 需要合并计算md5值的区域
>    - 返回值：
>       - md5: 单个数值，为data区域的所有值的合并字符串的md5值
> 2. udf_md5：计算整个区域中的各个单元格值的MD5值（空格也会纳入计算），并返回对应的MD5数组。
>    - 参数：
>       - data: 需要合并计算md5值的区域
>    - 返回值：
>       - range_md5: 数组公式，其范围大小与data一致，其中每个单元格的值为data对应区域的值的字符串的md5值

- 分配（udf_assign、udf_assign_multi）

> 当前支持：
>
> 1. udf_assign：根据指定的名称及对应数量返回对应名称列表。
>    - 参数：
>       - names: 需要分配任务的名称
>       - nums_to_assign: 需要分配任务的名称对应的分配量
>    - 返回值：
>       - name_list: 数组公式，一维名称列表，其行数等于nums_to_assign总和，值为按names顺序填充nums_to_assign次names对应字符串的字符串列表。
> 2. udf_assign_multi：根据指定的名称、任务名及对应数量返回对应名称数组。
>    - 参数：
>       - names: 需要分配任务的名称
>       - tasks: 需要分配的任务类型
>       - nums_to_assign: 需要分配任务的名称对应的分配量
>    - 返回值：
>       - task_name_list: 数组公式，两列多行的二维名称列表，其行数等于nums_to_assign总和，首列的值为按tasks顺序填充nums_to_assign次tasks对应字符串的字符串列表，第二列的值为按tasks、names顺序填充nums_to_assign次names对应字符串的字符串列表。

- 抽样（udf_sample）

> __注意，返回值为数组公式，重新打开文件后会发生变更，因此需要把对应结果复制至其他区域或粘贴为值。__
>
> 当前支持：
>
> 1. udf_sample：根据指定的抽样区域及抽样大小返回对应布尔数组结果。
>    - 参数：
>       - data: 需要抽样的范围
>       - sample_num: 抽样个数或百分比，录入需为正整数或0~1之间的小数
>    - 返回值：
>       - sample_array: 数组公式，一维布尔数组，大小与data行数一致，其中True的个数根据sample_num计算得出。
>           - sample_num大于data行数则返回错误提示文字。
>           - sample_num为0~1之间的小数：True个数为sample_num乘以data行数后舍去小数位
>           - sample_num为正整数：True个数为sample_num。
>
> 2. udf_sample_assign：根据指定的抽样区域、分配名称及分配量返回对应名称数组结果。
>    - 参数：
>       - data: 需要抽样的范围
>       - names: 需要分配抽样任务的名称
>       - nums_to_assign: 需要分配抽样任务的名称对应的分配量
>    - 返回值：
>       - sample_name_list: 数组公式，一维名称列表，大小与data行数一致，默认值为空，并按names顺序随机填充nums_to_assign次names中的对应字符串。
>           - nums_to_assign总和大于data行数则返回错误提示文字。

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


## todo list

1. 函数：

    - 暂无

2. 宏：

    - 暂无

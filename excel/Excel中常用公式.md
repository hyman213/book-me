# Excel中常用公式

## 文本

### FIND/FINDB函数

**在一个文本值中查找另一个文本值（区分大小写）-类似indexOf**

**函数 FIND 和 FINDB 用于在第二个文本串中定位第一个文本串，并返回第一个文本串的起始位置的值，该值从第二个文本串的第一个字符算起。**

| 函数名 | 语法 | 例子 |
| ---- | ------- | --- |
| FIND | FIND(find_text, within_text, [start_num]) | =FIND("M",A2) |
| FINDB | FINDB(find_text, within_text, [start_num]) | =FIND("M",A2,3) |

- **find_text**    必需。 要查找的文本
- **within_text**    必需。 包含要查找文本的文本
- **start_num**    可选。 指定开始进行查找的字符。 within_text 中的首字符是编号为 1 的字符。 如果省略 start_num，则假定其值为 1

**注：**

- **无论默认语言设置如何，函数 FIND 始终将每个字符（不管是单字节还是双字节）按 1 计数**
- **当启用支持 DBCS 的语言的编辑并将其设置为默认语言时，FINDB 会将每个双字节字符按 2 计数。 否则，FINDB 会将每个字符按 1 计数。**支持 DBCS 的语言包括日语、中文（简体）、中文（繁体）以及朝鲜语。 
- **当不存在文本时，可能会返回 #VALUE! 错误，一般结合ISNUMBER一起使用IF(ISNUMBER(FIND("2013",A1)),"2013","2016")**



## 查找和引用

### ROW/COLUMN函数

**返回引用的行号或列号**

| 函数名 | 语法             | 例子                |
| ------ | ---------------- | ------------------- |
| ROW    | ROW([reference]) | =ROW(C10)           |
| ROWS   | ROWS(array)      | =ROWS(C1:E4)，结果4 |
| COLUMN  | COLUMN([reference]) | =COLUMN() |
| COLUMNS | COLUMNS(array)      | =COLUMNS(E1:F7)，结果2 |

ROW 函数语法具有下列参数（COLUMN与此类似）

- **Reference**    可选。 需要得到其行号的单元格或单元格区域。
  - 如果省略 reference，则假定是对函数 ROW 所在单元格的引用。
  - 如果 reference 为一个单元格区域，并且 ROW 作为垂直数组输入，则 ROW 将以垂直数组的形式返回 reference 的行号。
  - Reference 不能引用多个区域。

ROWS函数语法具有下列参数：

- **Array**    必需。 需要得到其行数的数组、数组公式或对单元格区域的引用。

### MATCH函数
在范围单元格中搜索特定的项，然后返回该项在此区域中的相对位置
| 函数名 | 语法             | 例子                |
| ------ | ---------------- | ------------------- |
| MATCH    | **MATCH(lookup_value, lookup_array, [match_type])** | ==MATCH(25,A1:A3,0)     |

- **lookup_value**    必需。 要在 **lookup_array** 中匹配的值。**lookup_value** 参数可以为值（数字、文本或逻辑值）或对数字、文本或逻辑值的单元格引用。
- **lookup_array**    必需。 要搜索的单元格区域。
- **match_type**    可选。 数字 -1、0 或 1。 **match_type** 参数指定 Excel 如何将 **lookup_value** 与 **lookup_array** 中的值匹配。 此参数的默认值为 1。
  - 1或省略，**MATCH** 查找小于或等于 **lookup_value** 的最大值。 **lookup_array** 参数中的值必须以升序排序，例如：...-2, -1, 0, 1, 2, ..., A-Z, FALSE, TRUE。
  - **MATCH** 查找完全等于 **lookup_value** 的第一个值。 **lookup_array** 参数中的值可按任何顺序排列
  - **MATCH** 查找大于或等于 **lookup_value 的最小值。 lookup_array 参数中的值必须按降序排列，例如：TRUE, FALSE, Z-A, ...2, 1, 0, -1, -2, ... 等等。**

### INDEX函数

INDEX 函数返回表格或区域中的值或值的引用
| 函数名 | 语法             | 例子                |
| ------ | ---------------- | ------------------- |
| INDEX | **INDEX(array, row_num, [column_num])** | =INDEX(A2:B3,2,2)，位于区域 A2:B3 中第二行和第二列交叉处的数值。 |

- **数组**    必需。 单元格区域
- **row_num**    必需。 选择数组中的某行，函数从该行返回数值
- **column_num**    可选。 选择数组中的某列，函数从该列返回数值

一般结合`MATCH`和`INDEX`函数一起使用, 示例`=IF(MAX($G8:$BD8)>0,INDEX($G$3:$BD$3,1,MATCH(MAX($G8:$BD8),$G8:$BD8,0)),"")`

### [VLOOKUP函数](https://support.office.com/zh-cn/article/vlookup-函数-0bbc8083-26fe-4963-8ab8-93a18ad188a1)

在表格或区域中按行查找内容

`VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`

- `lookup_value`, 要查找的值

1. `table_array`, 查阅值所在的区域。 请记住，查阅值应该始终位于所在区域的第一列，这样 VLOOKUP 才能正常工作。 例如，如果查阅值位于单元格 C2 内，那么您的区域应该以 C 开头。
2. `col_index_num`, 区域中包含返回值的列号。 例如，如果指定 B2：D11 作为区域，那么应该将 B 算作第一列，C 作为第二列，以此类推。
3. `range_lookup`,（可选）如果需要返回值的近似匹配，可以指定 TRUE；如果需要返回值的精确匹配，则指定 FALSE。 如果没有指定任何内容，默认值将始终为 TRUE 或近似匹配。

### HLOOKUP函数

在表格或区域中按行查找内容

`HLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`

- `lookup_value`, 要查找的值

1. `table_array`, 查阅值所在的区域。 请记住，查阅值应该始终位于所在区域的第一列，这样 VLOOKUP 才能正常工作。 例如，如果查阅值位于单元格 C2 内，那么您的区域应该以 C 开头。
2. `col_index_num`, 区域中包含返回值的列号。 例如，如果指定 B2：D11 作为区域，那么应该将 B 算作第一列，C 作为第二列，以此类推。
3. `range_lookup`,（可选）如果需要返回值的近似匹配，可以指定 TRUE；如果需要返回值的精确匹配，则指定 FALSE。 如果没有指定任何内容，默认值将始终为 TRUE 或近似匹配。












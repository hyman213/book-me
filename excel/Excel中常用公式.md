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

### COUNTIF函数

`COIUNTIF`函数用于统计一个区域中符合条件的单元格个数

`=IF(COUNTIF(A:A,A2)>1,"重复","")`, 先使用COIUNTIF函数计算出A列区域中有多少个与A2相同的姓名。然后使用IF函数判断，如果COIUNTIF函数的结果大于1，就说明有重复了

`COUNTIF(range, criteria)`

- 参数range 表示条件区域——对单元格进行计数的区域。

- 参数criteria 表示条件——条件的形式可以是数字、表达式或文本，甚至可以使用通配符

`=COUNTIF(B1:B7,">"&D1)`, 计算B1:B7中大于D1的个数

### AVERAGEIF

求某个区域内满足给定条件指定的单元格的平均值

`averageif(range, criteria, [average_range])`

* 参数Range表示：**条件区**——第二个参数条件所在的范围
* 参数Criteria表示：**条件**——是用来定义计算平均值的单元格
* 参数Average_range：**平均值区域**——参与计算平均值的单元格。（这参数可以省略，当条件区和平均值区域一致时）

示例：`=AVERAGEIF(B2:B7,"男",C2:C7)`

如果需要判断多个条件，可使用**AVERAGEIFS**

`=AVERAGEIFS(average_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)`

### DATEIF

主要用于计算两个日期之间的天数、月数或年数。

`DATEDIF(start_date,end_date,unit)`

- 参数1：start_date，表示起始日期

- 参数2：end_date，表示结束日期。参数1和参数2可以是带引号的文本串（例如："2014-1-1"）、系列号或者其他公式或函数的结果

- 参数3：unit为所需信息的返回时间单位代码。各代码含义如下：
  - "y"返回时间段中的整年数
  - "m”返回时间段中的整月数
  - "d"返回时间段中的天数
  - "md”参数1和2的天数之差，忽略年和月
  - "ym“参数1和2的月数之差，忽略年和日
  - "yd”参数1和2的天数之差，忽略年。按照月、日计算天数

计算年龄：`=DATEDIF(A2,TODAY(),"y")`

说明：TODAY函数返回系统当前的日期。DATEDIF函数以A2的出生年月作为开始日期，以系统日期作为结束日期，第三参数使用Y，表示计算两个日期之间的整年数。

### MID

`MID(text, start_num, num_chars)`

* text是需要查找的字符串文本
* start_num是查找字符串文本中的起始位置，从第一个字符计算，默认为1
* num_chars是所从起始位置开始的提取字符串个数，num_chars不可为负数，如大于文本长度，则提取剩余文本。

提取出生年月：`=--TEXT(MID(A2,7,8),"0-00-00")`

### TEXT

将数值转化为自己想要的文本格式

`=text(value,format_text）`

* Value为数字值
* Format_text为设置单元格格式中自己所要选用的文本格式。






















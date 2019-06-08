# Excel中常用公式

## 文本

###  在一个文本值中查找另一个文本值（区分大小写）-类似indexOf

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

### 返回引用的行号或列号

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
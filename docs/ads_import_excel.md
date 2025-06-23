# %ads_import_excel

## 简介

读取 Excel 数据，创建 SAS 数据集。

## 语法

### 参数

#### 必选参数

- [file](#file)
- [outdata](#outdata)
- [sheet_name](#sheet_name)

#### 可选参数

- [dbms](#dbms)
- [range_attr](#range_attr)
- [range_data](#range_data)
- [all_chars](#all_chars)
- [clear_format](#clear_format)
- [clear_informat](#clear_informat)
- [ignore_empty_line](#ignore_empty_line)
- [warning_var_name_empty](#warning_var_name_empty)
- [warning_var_name_not_meet_v7](#warning_var_name_not_meet_v7)
- [warning_var_name_len_gt_8](#warning_var_name_len_gt_8)

#### 调试参数

- [debug](#debug)

### 参数说明

#### file

**Syntax** : _physical_path_ | _filename_

指定 Excel 文件路径。

> [!IMPORTANT]
>
> 1. `file` 文件必须是 `.xlsx`, `.xlsm`, `.xls` 格式
> 2. `file` 文件必须包含一张定义了需要导入的变量和数据的工作表 `sheet_name`，可使用参数 [sheet_name](#sheet_name) 指定工作表名称
> 3. `file` 文件的工作表 `sheet_name` 中，必须包含一个 $2 \times C$ 单元格范围 `range_attr`，该单元格范围中的第一行必须是对变量标签的定义，第二行必须是对变量名的定义，可使用 [range_attr](#range_attr) 指定这个范围
> 4. `file` 文件的工作表 `sheet_name` 中，必须包含一个 $R \times C$ 单元格范围 `range_data`，该单元格范围中的数据即为需要导入的数据，可使用 [range_data](#range_data) 指定这个范围

**Usage** :

```sas
file = %str(~\原始数据\方案偏离清单-20250125.xlsx)
```

---

#### outdata

**Syntax** : _data-set-name_<(_data-set-option_)>

指定需要创建的 SAS 数据集名称，可使用数据集选项。

**Usage** :

```sas
outdata = dv
outdata = dv(where = (dvyn = "是"))
```

---

#### sheet_name

**Syntax** : _string_

指定 Excel 文件中包含变量定义和数据定义的工作表名称。

**Usage** :

```
sheet_name = %str(方案偏离清单)
```

---

#### dbms

指定源数据类型。

**Syntax** : _dbms_ | `#auto`

**Default** : `#auto`

默认情况下，参数 [dbms](#dbms) 的值取决于参数 [file](#file) 指定的文件的后缀名，如下表所示：

| 后缀名           | _dbms_ 取值 |
| ---------------- | ----------- |
| `.xls`           | `excel`     |
| `.xlsx`, `.xlsm` | `xlsx`      |

---

#### range_attr

指定 [sheet_name](#sheet_name) 工作表中变量定义的单元格范围。

**Syntax** : _range_ | `#null`

> [!IMPORTANT]
>
> - _range_ 是符合正则表达式 `^[A-Za-z]+[0-9]+:[A-Za-z]+[0-9]+$` 的字符串，例如：`A1:F255`
> - _range_ 指定的范围有以下三种类型：
>   - 若 _range_ 指定的范围是一行（例如 `A1:F1`），则程序自动检测这一行单元格中的文本是否均为 `VALIDVARNAME=V7` 下的合法变量名，若合法，则这一行用作输出数据集的变量标签和变量名；
>   - 若 _range_ 指定的范围是两行（例如 `A1:F2`），则第一行被用作输出数据集的变量标签，第二行被用作输出数据集的变量名；
>   - 若 _range_ 指定的范围是 $n$ 行 ($n \ge 3$)，则第一行被用作输出数据集的变量标签，第二行被用作输出数据集的变量名，第 $n$ ($n \ge 3$) 行将被忽略。

**Default** : `#null`

默认情况下，[sheet_name](#sheet_name) 工作表中的前两行被视为变量定义的单元格范围。

**Usage** :

```sas
range = %str(A1:U2)
```

---

#### range_data

指定 [sheet_name](#sheet_name) 工作表中数据定义的单元格范围。

**Syntax** : _range_ | `#null`

> [!IMPORTANT]
>
> - _range_ 是符合正则表达式 `^[A-Za-z]+[0-9]+:[A-Za-z]+[0-9]+$` 的字符串，例如：`A1:F255`

**Default** : `#null`

默认情况下，[sheet_name](#sheet_name) 工作表中的第三行及之后被视为数据定义的单元格范围。

**Usage** :

```sas
range = %str(A3:U255)
```

---

#### all_chars

指定是否将所有变量都视为字符型变量。

**Syntax** : `true` | `false`

**Default** : `true`

> [!IMPORTANT]
>
> - `all_chars = true` 时，宏程序会以字符形式读取所有变量，这在 Excel 文件频繁更新的情况下非常有用，
>   因为可能存在某一次数据的更新导致通过 `PROC IMPORT` 过程读入 SAS 时，某个变量的类型由数值型变为字符型，进而导致相关程序不得不频繁更新，给编程带来不必要的麻烦；
> - `all_chars = true` 时，日期型变量的值会以 Excel 内部数值的字符串形式进行表示，可以使用 `input(xxx, 8.) + "30DEC1899"d` 转为 SAS 日期型变量；
> - `all_chars = true` 时，部分数值可能以科学计数法的形式表示，可以使用 `input(xxx, best32.)` 转为一般形式；
> - `all_chars = true` 时，变量的前导空格将会被删除；
> - `all_chars = true` 时，所有变量的输出格式和输入格式都将被清除，无需指定 `clear_format = true` 或 `clear_informat = true`。

> [!IMPORTANT]
>
> 修改参数 `all_chars` 的值实际上是在修改宏变量 [EFI_ALLCHARS](https://documentation.sas.com/doc/zh-CN/pgmsascdc/9.4_3.5/proc/p12uk352fte2h1n1efh4r44dmxmp.htm#n0eph5ecq6gxten1cl8kswhtepon) 的值。

---

#### clear_format

指定是否清除所有变量的输出格式。

**Syntax** : `true` | `false`

**Default** : `true`

> [!WARNING]
>
> `all_chars = true` 时，参数 `clear_format` 无效。

---

#### clear_informat

指定是否清除所有变量的输入格式。

**Syntax** : `true` | `false`

**Default** : `true`

> [!WARNING]
>
> `all_chars = true` 时，参数 `clear_informat` 无效。

---

#### ignore_empty_line

指定是否忽略空行。

**Syntax** : `true` | `false`

**Default** : `true`

---

#### warning_var_name_empty

指定当变量名为空时，是否输出警告信息。

**Syntax** : `true` | `false`

**Default** : `true`

---

#### warning_var_name_not_meet_v7

指定当变量名包含 `VALIDVARNAME=V7` 下的非法字符时，是否输出警告信息。

> [!NOTE]
>
> - 使用函数 [NOTNAME](https://support.sas.com/documentation/cdl/en/lrdict/64316/HTML/default/viewer.htm#a002197357.htm) 判断变量名是否符合 `VALIDVARNAME=V7` 的要求
> - 有关 SAS 系统选项 `VALIDVARNAME` 的更多信息，请参考 [VALIDVARNAME= System Option - SAS Support](https://support.sas.com/documentation/cdl/en/acreldb/63647/HTML/default/viewer.htm#a000436063.htm)

**Syntax** : `true` | `false`

**Default** : `true`

---

#### warning_var_name_len_gt_8

指定当变量名长度超过 8 时，是否输出警告信息。

**Syntax** : `true` | `false`

**Default** : `true`

---

#### debug

指定是否删除中间过程生成的数据集。

> [!NOTE]
>
> 这是一个用于开发者调试的参数，通常不需要关注。

**Syntax** : `true` | `false`

**Default** : `false`

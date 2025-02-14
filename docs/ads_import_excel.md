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

- [range_attr](#range_attr)
- [range_data](#range_data)
- [warning_var_name_empty](#warning_var_name_empty)
- [warning_var_name_not_meet_v7](#warning_var_name_not_meet_v7)
- [warning_var_name_len_gt_8](#warning_var_name_len_gt_8)

#### 调试参数

- [debug](#debug)

### 参数说明

#### file

**Syntax** : _filename_

指定 Excel 文件路径。

> [!IMPORTANT]
>
> 1. `file` 文件必须是 `.xlsx` 格式
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

#### range_attr

指定 [sheet_name](#sheet_name) 工作表中变量定义的单元格范围。

**Syntax** : _range_ | `#null`

> [!IMPORTANT]
>
> - _range_ 是符合正则表达式 `^[A-Za-z]+[0-9]+:[A-Za-z]+[0-9]+$` 的字符串，例如：`A1:F255`
> - _range_ 指定的范围中，第一行必须是变量标签的定义，第二行必须是变量名的定义

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

# %ads_empty

## 简介

读取 ADaM Specification 文件，根据变量定义创建空白 ADaM 数据集。

## 语法

### 参数

#### 必选参数

- [spec](#spec)

#### 可选参数

- [select](#select)
- [prefix](#prefix)
- [suffix](#suffix)
- [sheet_name](#sheet_name)
- [range](#range)

#### 调试参数

- [debug](#debug)

### 参数说明

#### spec

**Syntax** : _filename_

指定 ADaM Specification 文件路径。

> [!IMPORTANT]
>
> 1. `spec` 文件必须是 `.xlsx` 格式
> 2. `spec` 文件必须包含一张定义了所需全部 ADaM 数据集变量的工作表 `sheet_name`，可使用参数 [sheet_name](#sheet_name) 指定工作表名称
> 3. `spec` 文件的工作表 `sheet_name` 中，A~F 列的定义如下：
>
>    | 列名 | 列标题   | 示例        |
>    | ---- | -------- | ----------- |
>    | A    | 数据集   | _ADSL_      |
>    | B    | 变量名   | _RANDDT_    |
>    | C    | 变量标签 | _入组日期_  |
>    | D    | 变量类型 | _Num_       |
>    | E    | 长度     | _8_         |
>    | F    | 显示格式 | _YYMMDD10._ |
>
>    - `A` 列不能为空，程序将根据 `A` 列决定输出的数据集名称
>    - `B` 列不能为空，程序将跳过未命名的变量
>    - 当 `A` 列的值相同时，`B` 列的值不允许重复，否则程序将直接退出
>    - `C` 列的值不能为空，否则程序将复制 `B` 列的值作为 `C` 列的值，并发出警告
>    - `D` 列的值只能是 `Char` 或 `Num`，当 `D` 列为空时，程序将其视为 `Char`，并发出警告
>    - `E` 列的值不能为空，否则程序将根据 `D` 列的值自动设置 `E` 列的值，对于 `Char`，设置为 `200`，对于 `Num`，设置为 `8`

**Usage** :

```sas
spec = %str(~\ADS程序\分析数据库编程说明\ADaM_specification_V1.0_20250115.xlsx)
```

---

#### select

**Syntax** : _dataset-1_ <, _dataset-2_ <, ...>>

指定需要创建的空白 ADaM 数据集名称（列表）。

**Default** : `#ALL`

默认情况下，宏程序会尝试创建 [sheet_name](#sheet_name) 工作表中定义的所有 ADaM 数据集。

**Usage** :

```sas
select = adsl
select = adsl adae addv
```

---

#### prefix

**Syntax** : _string_

指定空白 ADaM 数据集的名称前缀。

**Default** : ` `

**Usage** :

```sas
prefix = %str(tmp_)
```

---

#### suffix

**Syntax** : _string_

指定空白 ADaM 数据集的名称后缀。

**Default** : `_empty`

**Usage**:

```sas
suffix = %str(_empty)
```

---

#### sheet_name

**Syntax** : _string_

指定 ADaM Specification 文件中包含变量定义的工作表名称。

**Default** : `变量说明`

**Usage** :

```
sheet_name = %str(变量说明)
```

---

#### range

指定 [sheet_name](#sheet_name) 工作表中的特定范围。

**Syntax** : _range_

> [!IMPORTANT]
>
> - _range_ 是符合正则表达式 `^[A-Za-z]+[0-9]+:[A-Za-z]+[0-9]+$` 的字符串，例如：`A1:F255`
> - _range_ 指定的范围中，第一行必须是标题行

**Default** : `#ALL`

默认情况下，[sheet_name](#sheet_name) 工作表中的所有变量定义都会被读取。

**Usage** :

```sas
range = %str(A1:F255)
```

---

#### debug

指定是否删除中间过程生成的数据集。

> [!NOTE]
>
> 这是一个用于开发者调试的参数，通常不需要关注。

**Syntax** : `true` | `false`

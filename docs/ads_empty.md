# %ads_empty

## 简介

读取 ADaM Specification 文件，根据变量定义创建空白 ADaM 数据集。

## 语法

### 参数

#### 必选参数

- [spec](#spec)

#### 可选参数

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
> 仅支持 `.xlsx` 格式的 ADaM Specification 文件

**Usage** :

```sas
spec = %str(~\ADS程序\分析数据库编程说明\ADaM_specification_V1.0_20250115.xlsx)
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

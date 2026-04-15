# %convert_to_xpt

## 简介

`sas7bdat` -> `xpt` 批量转换。

## 语法

### 参数

#### 必选参数

- [libref](#libref)
- [dir](#dir)
- [format](#format)

### 参数说明

#### libref

**Syntax** : _libref_

指定需转换的 `.sas7bdat` 文件所在的逻辑库名称。

**Usage** :

```sas
libref = adam
```

---

#### dir

**Syntax** : _physical_path_

指定转换后生成的 `.xpt` 文件存放的目录的路径。

**Usage** :

```sas
dir = %str(~\分析数据\XPT)
```

> [!NOTE]
>
> 如果 `dir` 指定的目录不存在，则该目录将自动创建。

---

#### format

**Syntax** : `v5` | `v8` | `auto`

指定 XPT 格式的版本。

> [!WARNING]
>
> 若指定 `format = auto`，可能会导致 `dir` 中同时存在 `v5` 和 `v8` 格式的 XPT 文件。

**Default** : `v8`

**Usage** :

```sas
format = v5
```

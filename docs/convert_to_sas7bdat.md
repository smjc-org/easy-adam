# %convert_to_sas7bdat

## 简介

`xpt` -> `sas7bdat` 批量转换。

## 语法

### 参数

#### 必选参数

- [dir](#dir)

#### 可选参数

- [libref](#libref)

### 参数说明

#### dir

**Syntax** : _physical_path_

指定需转换的 `.xpt` 文件所在目录的路径。

**Usage** :

```sas
dir = %str(~\分析数据\XPT)
```

---

#### libref

**Syntax** : _libref_

指定转换后生成的 `.sas7bdat` 文件存放的逻辑库名称。

**Default**: `work`

**Usage** :

```sas
libref = work
```

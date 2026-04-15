# %convert_to_xpt

## 简介

`sas7bdat` -> `xpt` 批量转换。

## 语法

### 参数

#### 必选参数

- [sas7bdat_dir](#base_dir)
- [xpt_dir](#compare_dir)
- [format](#outdata)

#### 调试参数

- [debug](#debug)

### 参数说明

#### sas7bdat_dir

**Syntax** : _physical_path_

指定需转换的 `.sas7bdat` 文件所在目录的路径。

**Usage** :

```sas
file = %str(~\分析数据)
```

---

#### xpt_dir

**Syntax** : _physical_path_

指定转换后生成的 `.xpt` 文件所在目录的路径。

**Usage** :

```sas
file = %str(~\分析数据\XPT)
```

> [!NOTE]
>
> 如果 `xpt_dir` 指定的目录不存在，则该目录将自动创建。

---

#### format

**Syntax** : `v5` | `v8` | `auto`

指定 XPT 格式的版本，默认为 `v8`。

**Usage** :

```sas
format = v5
```

---

#### debug

指定是否开启调试模式。

> [!NOTE]
>
> 这是一个用于开发者调试的参数，通常不需要关注。

**Syntax** : `true` | `false`

**Default** : `false`

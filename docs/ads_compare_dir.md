# %ads_compare_dir

## 简介

比较两个文件夹下的 analysis datasets。

## 语法

### 参数

#### 必选参数

- [base_dir](#base_dir)
- [compare_dir](#compare_dir)
- [outdata](#outdata)

#### 调试参数

- [debug](#debug)

### 参数说明

#### base_dir

**Syntax** : _physical_path_

指定基准文件夹的路径。

**Usage** :

```sas
file = %str(~\分析数据\draft\20260303)
```

---

#### compare_dir

**Syntax** : _physical_path_

指定比较文件夹的路径。

**Usage** :

```sas
file = %str(~\分析数据)
```

---

#### outdata

**Syntax** : _data-set-name_<(_data-set-option_)>

指定一个 SAS 数据集名称，这个数据集将用于保存比较结果。

`outdata` 数据集包含的变量及其含义如下：

| 变量名称    | 含义                     |
| ----------- | ------------------------ |
| `b_memname` | 基准文件夹下的数据集名称 |
| `c_memname` | 比较文件夹下的数据集名称 |
| `result`    | 比较结果的文字描述       |
| `comment`   | 备注                     |

变量 `result` 的取值是以下其中之一：

- _空值_
- `base 中不存在`
- `compare 中不存在`
- _以下描述中的一种或几种组合：_
  - `数据集标签不一致`
  - `数据集类型不一致`
  - `存在输入格式不同的变量`
  - `存在输出格式不同的变量`
  - `存在长度不同的变量`
  - `存在标签不同的变量`
  - `base 数据集存在 compare 数据集没有的观测`
  - `compare 数据集存在 base 数据集没有的观测`
  - `base 数据集存在 compare 数据集没有的 BY 组`
  - `compare 数据集存在 base 数据集没有的 BY 组`
  - `base 数据集存在 compare 数据集没有的变量`
  - `compare 数据集存在 base 数据集没有的变量`
  - `存在不同值`
  - `变量类型冲突`
  - `BY 变量不匹配`
  - `致命错误：未进行比较`

变量 `comment` 的取值是以下其中之一：

- _空值_
- `base 数据集为空`
- `compare 数据集为空`
- `base 和 compare 数据集均为空`

**Usage** :

```sas
outdata = diff_ads
outdata = diff_ads(where = (not missing(result)))
```

---

#### debug

指定是否开启调试模式。

> [!NOTE]
>
> 这是一个用于开发者调试的参数，通常不需要关注。

**Syntax** : `true` | `false`

**Default** : `false`

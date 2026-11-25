# pads-logic-eco

从 ECN 变更表和原理图 BOM 生成 PADS Logic ECO，通过校验才写出文件。

## 环境

- Windows，Python 3.12
- `pip install -r requirements.txt`
- 亿赛通加密的 Excel 需要本机安装 Excel（经 COM 读取，不改原文件）

器件库放在 `library/Electronic_components_library.xlsx`。

## 运行

```
python tool/pads_eco_app.py
```

界面五步：更新器件库 → 导入 ECN → 导入原理图 BOM → 生成 ECO → 反向检查。

命令行：

```
python tool/pads_eco_app.py gen 你的ECN.xlsx --bom 原理图BOM.xlsx
python tool/pads_eco_app.py check 你的.eco --bom 原理图BOM.xlsx
```

ECO 写在 ECN 同目录、同名 `.eco`，并附 `.map.csv` 对照表。

## 输入

| 文件 | 要求 |
|---|---|
| ECN | 文件名一般含 `Bom&ECN-changelist`，工作表 `RF_ECN_Generation`，Change Comments 为 物料修改 / 物料删除 / 物料新增 |
| 原理图 BOM | 从 PADS Logic 导出，必须带 Part Type 和 Value。采购 BOM / EBOM 不能用 |
| 器件库 | 表头要有「元件名称」（或「元件类型」）和 Part Number。第 1 步用 CSV 更新 |

当前 Part Type 只从原理图 BOM 读。

## 生成规则

| ECN 动作 | ECO |
|---|---|
| 物料修改 | `*CHGPART*`  BOM 当前类型 → After 料号对应的库元件名称 |
| 物料删除 | `*SET_ATTRIBUTE*` PART 位号，Value 前加 `@`，Part Type 不变 |
| 物料新增 | `*CHGPART*` 经同族 NC 两跳到 After |

Value 已是 `@` 的修改同样经 NC 两跳。封装前后不同会弹窗，点「ECN 变更表」可打开原表核对；确认后才继续生成。

## 导入 PADS Logic

1. 备份原理图
2. 断开 Layout
3. 关闭在线 DRC
4. 文件 → 导入 → 文件类型 .ECO → 选择此次生成的 ECO 文件

## 反向检查

原理图导入并改完后，重新导出 BOM，点第 5 步选这份新 BOM。工具会复制 ECN 为「原名.反向检查.xlsm」再跑宏，原文件不改。变更表没有数据行即通过。

## 打包

```
pip install pyinstaller openpyxl pywin32 pillow
powershell -ExecutionPolicy Bypass -File tool/build.ps1
```

产物：`dist/pads-eco/pads-eco.exe`

## 测试

```
python -m pytest tests -q
```

## 目录

```
scripts/          生成、校验、反向检查
tool/pads_eco_app.py   界面
tool/build.ps1    打包
tests/            回归测试
library/          器件库 xlsx（本机准备，不进仓库）
```

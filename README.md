# 非小酋转 Paimon.moe / UIGF 抽卡记录转换器

[English](README.en.md) | [中文](#)

本项目用于处理「非小酋」导出的原神抽卡记录 Excel 文件。它读取非小酋的抽卡记录 xlsx，并将其转换为两种更通用的数据格式：

- **Paimon.moe Excel 格式**：可直接用于 Paimon.moe 的抽卡记录导入/导出。
- **UIGF JSON 格式**：遵循 UIGF 统一可交换抽卡记录标准，默认输出最新 v4.2，同时提供 v3.0 选项，便于需要旧版格式的场景使用。

转换过程会自动处理中英文物品名、卡池匹配、保底/抽数/分组等字段，并尽量保留原有数据完整性。

## 项目结构

```text
feixiaoqiu-to-paimonmoe/
├── converter.py              # 转 Paimon.moe Excel
├── gui.py                    # 图形界面
├── to_uigf.py                # 转 UIGF JSON（v4.2 / v3.0）
├── create_demo_data.py       # 生成演示数据
├── update_data.py            # 刷新内置数据
├── build_exe.bat             # 打包 Windows exe
├── requirements.txt
├── VERSION
├── README.md
├── README.en.md
├── data/
│   ├── zh_to_en_items.json   # 中文物品名 -> 英文物品名
│   ├── item_ids.json         # Paimon.moe 物品 ID
│   ├── uigf_numeric_item_ids.json  # UIGF 数字 item_id 映射
│   └── banners.json          # 卡池时间表 + 双卡池信息
├── pic/
│   ├── 非小酋_导出excel示例.png
│   └── paimonmoe_导入excel示例.png
└── examples/
    ├── input_demo.xlsx       # 随机生成的「非小酋」格式演示输入
    ├── output_demo.xlsx      # 转换后的 Paimon.moe 格式演示输出
    └── output_demo.json      # 转换后的 UIGF 格式演示输出
```

## 安装依赖

```bash
pip install -r requirements.txt
```

## 转 Paimon.moe Excel

```bash
python converter.py "非小酋抽卡记录.xlsx" -o paimonmoe_wish_history_converted.xlsx
```

也可以省略输入文件，自动查找当前目录下的 `非小酋抽卡记录*.xlsx`：

```bash
python converter.py
```

### 可选参数

```bash
# 禁用双卡池 Wish 2 启发式识别
python converter.py --no-heuristic -o output.xlsx

# 指定数据目录
python converter.py --data-dir data -o output.xlsx
```

## 转 UIGF JSON

默认输出 UIGF v4.2：

```bash
python to_uigf.py "非小酋抽卡记录.xlsx" --uid 123456789 -o uigf_wish_history.json
```

如果需要输出 UIGF v3.0：

```bash
python to_uigf.py "非小酋抽卡记录.xlsx" --uid 123456789 --uigf-version v3.0 -o uigf_v3.json
```

- `--uid` 是必填项，因为非小酋 Excel 里没有 UID。
- `--uigf-version` 默认 `v4.2`，可改为 `v3.0`。
- `--lang` 默认 `zh-cn`，可改为 `en` 输出英文物品名。
- `--timezone` 默认 `8`（Asia 时区）。

```bash
# 英文输出
python to_uigf.py "非小酋抽卡记录.xlsx" --uid 123456789 --lang en -o uigf_en.json

# 禁用双卡池 400 启发式识别
python to_uigf.py "非小酋抽卡记录.xlsx" --uid 123456789 --no-heuristic -o uigf.json
```

## 演示数据

仓库内置了随机生成的演示数据，可直接查看效果：

```bash
python converter.py examples/input_demo.xlsx -o examples/output_demo.xlsx
python to_uigf.py examples/input_demo.xlsx --uid 123456789 -o examples/output_demo.json
```

重新生成演示数据：

```bash
python create_demo_data.py
```

## 输出说明：Paimon.moe Excel

生成文件包含：

| Sheet | 说明 |
|---|---|
| Character Event | 角色活动祈愿 |
| Weapon Event | 武器活动祈愿 |
| Standard | 常驻祈愿 |
| Beginners' Wish | 新手祈愿 |
| Chronicled Wish | 集录祈愿（额外保留，Paimon.moe Excel 导入会忽略此 sheet） |
| Banner List | 卡池时间表 |
| Information | 导出信息 |

## 更新数据

当新版本/新卡池/新物品出现后，可以刷新内置数据：

```bash
python update_data.py
```

该命令会从公开数据源重新拉取：

- 中文/英文物品名映射
- UIGF 数字 item_id 映射
- 卡池时间表与双卡池信息

## 无需 Python 环境的封装版本

仓库提供 `build_exe.bat`，会自动创建隔离环境并打包成 Windows 可执行文件：

```bash
build_exe.bat
```

打包后会在 `dist/` 下生成：

```text
feixiaoqiu-to-paimonmoe-gui.exe   # 图形界面版，双击即可使用（推荐）
feixiaoqiu-to-paimonmoe.exe       # 命令行版：转 Paimon.moe Excel
feixiaoqiu-to-uigf.exe            # 命令行版：转 UIGF JSON
```

终端用户不需要安装 Python，直接运行 exe 即可。

## 已知限制

1. 非小酋没有提供具体卡池名，转换器通过卡池时间表自动匹配。
2. 非小酋没有提供 `301/400` 双卡池标记，转换器使用启发式识别 `Wish 2`：
   - 同一时间戳的十连内出现第二个卡池的 5 星 UP 角色时，整组标记为 `Wish 2`；
   - 无法判断的记录默认按第一个卡池处理。
3. Paimon.moe 当前 Excel 导入/导出不支持集录祈愿，因此 `Chronicled Wish` sheet 仅用于保留数据，不会被 Paimon.moe Excel 导入读取。
4. UIGF 导出需要手动提供 UID。

## 界面示意

### 非小酋导出 Excel

![非小酋导出Excel示例](pic/非小酋_导出excel示例.png)

### Paimon.moe 导入 Excel

![Paimon.moe导入Excel示例](pic/paimonmoe_导入excel示例.png)

## 相关资源

- UIGF 标准：<https://uigf.org/zh/standards/uigf.html>
- 非小酋：<https://feixiaoqiu.com/>
- Paimon.moe：<https://paimon.moe/>
- 提瓦特指南（Teyvat Guide）：<https://github.com/BTMuli/TeyvatGuide>
- genshin-wish-export：<https://github.com/biuuu/genshin-wish-export>

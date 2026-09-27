# ClassIsland.ExcelTool · 档案 Excel 导入导出

把 **ClassIsland** 的档案（时间表、科目、课程表、全局配置）导出成一份 Excel 工作簿，
改完之后再从 Excel 反向导入回 ClassIsland。

适用于批量排课的场合：在 Excel 里整列复制、公式计算、查找替换，比在设置窗口里一个个点快得多。

---

## 功能

### 导出

一次性把档案写成同一个 `.xlsx` 的多个工作表：

| 工作表 | 内容 |
| --- | --- |
| `时间表` | 时间表名称 / ID、是否叠放、叠放源 ID、是否启用，以及各时间点的起止时间 |
| `科目` | 科目名称、缩写、ID |
| `课程表` | 课表名称 / ID、所属时间表、行号、各时段对应的科目与教师 |
| `课表清单` | 课表 ID、名称、时间表 ID / 名称、是否启用、是否叠放 |
| `全局配置` | 主界面、提醒、天气等全局设置项 |

导出时可勾选要写入哪些工作表，也可以选择「加表头样式并冻结首行」，
方便在 Excel 里直接筛选和排序。

### 导入

导入前先用 **「预览差异」** 看清楚会发生什么改动，确认无误再点 **「应用导入」**。

三种导入策略：

| 策略 | 行为 |
| --- | --- |
| **合并**（推荐） | Excel 里有的就更新，本地多出来的条目保留 |
| **覆盖** | 以 Excel 为准，本地多出来的条目会被删除 |
| **跳过已存在** | 只新增，已存在的条目一律不动 |

导入按**表头名字**定位列：缺哪一列不影响其它列，表里出现了未识别的列会原样保留，
不会因为多加一列备注就把整份档案读坏。

---

## 安装

1. 下载 Release 里的压缩包（或自行编译）。
2. 解压到 `<ClassIsland 程序目录>\data\Plugins\classisland.exceltool\`。
   目录里应当能看到 `manifest.yml`、`ClassIsland.ExcelTool.dll`、`icon.png`。
3. 重启 ClassIsland。

## 使用

打开 **设置 → 外部 → 档案 Excel 导入导出**：

1. **导出**：点「📤 导出为 Excel」会直接弹出保存对话框，选个位置即可。
   也可以先点「仅选择路径…」定好路径再导出。
2. **导入**：点「📥 预览差异」选文件 → 看日志里的差异报告 → 点「✅ 应用导入」。
3. 建议保持 **「导入前自动备份档案」** 勾选。备份默认落在插件配置目录的 `Backups` 下，
   也可以在界面里指定别的目录。

> ⚠️ 导入会写回 ClassIsland 的档案文件。动手前请确认已勾选自动备份，
> 或者自己手动复制一份 `data\Profiles`。

---

## 从源码编译

```powershell
git clone https://github.com/你的用户名/ClassIsland.ExcelTool.git
cd ClassIsland.ExcelTool
dotnet build -c Release
```

产物在 `bin\Release\`，把 `manifest.yml`、`ClassIsland.ExcelTool.dll`、`icon.png`
以及依赖的 dll 一起丢进 `data\Plugins\classisland.exceltool\` 即可。

### 依赖

- .NET 8.0
- [`ClassIsland.PluginSdk`](https://www.nuget.org/packages/ClassIsland.PluginSdk) 2.0.1（`PrivateAssets=all`，不打进产物）
- [`ClosedXML`](https://www.nuget.org/packages/ClosedXML) 0.104.2 —— 纯托管 Excel 读写库，无本地依赖，可随插件分发

---

## 实现说明

导出和导入都没有手写 ClassIsland 的模型映射，而是以 `System.Text.Json` 的 `JsonNode`
作为中间表示直接读写档案 JSON。好处是 ClassIsland 升级增删字段时，插件不会被模型类绑死 ——
认得的字段照常读写，认不得的字段原样透传保留。

- 读档案时按**字段的序列化名**取值（例如时间表的激活字段是 `IsActive` 而不是 `IsActivated`）。
- 空时间表也会占一行写出，否则导入时无从得知它的存在。

---

## 许可

未指定。作者：鲸娘。

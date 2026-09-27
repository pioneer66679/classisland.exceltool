using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Nodes;
using ClosedXML.Excel;

namespace ClassIsland.ExcelTool.Core;

/// <summary>
/// Excel → 档案 JSON 的导入引擎。
///
/// 核心策略（保证不写坏档案）：
///   1. 优先读取每个 Sheet 里的「原始行(JSON)」列 —— 这是导出时逐行存下的完整对象，
///      用它可以直接重建对象，字段零丢失。
///   2. 原始行缺失或不存在该列（用户手写的 Excel）时，退化为按可见列拼装必要字段。
///   3. 无论哪条路径，都在内存里构造出一个新的 JsonNode 树，由调用方决定覆盖/合并/跳过，
///      并且整个过程先算差异、后落盘，绝不做「边读边写」。
/// </summary>
public static class ImportEngine
{
    /// <summary>导入解析的产物：一份待应用的档案 + 差异报告。</summary>
    public sealed class Result
    {
        public JsonObject Profile { get; init; } = new();
        public JsonObject Settings { get; init; } = new();
        public DiffReport Diff { get; init; } = new();
        public List<string> Warnings { get; } = new();
    }

    /// <summary>
    /// 解析 Excel 得到「Excel 想要的目标状态」，并与当前档案比对出差异。
    /// </summary>
    public static Result Parse(
        string excelPath,
        JsonNode currentProfile,
        JsonNode currentSettings,
        IEnumerable<string> enabledSheets)
    {
        var result0 = new Result();
        var on = new HashSet<string>(enabledSheets, StringComparer.Ordinal);

        using var wb = new XLWorkbook(excelPath);

        var profile = new JsonObject();
        // 先把当前档案里、Excel 不负责的部分整体带过来，避免导入把无关配置抹掉
        var cur = currentProfile as JsonObject ?? new JsonObject();
        foreach (var kv in cur)
            profile[kv.Key] = kv.Value?.DeepClone();

        var settings = new JsonObject();
        var curSettings = currentSettings as JsonObject ?? new JsonObject();
        foreach (var kv in curSettings)
            settings[kv.Key] = kv.Value?.DeepClone();

        if (on.Contains(SheetNames.TimeLayouts) && wb.TryGetWorksheet(SheetNames.TimeLayouts, out var wsTL))
            ParseTimeLayouts(wsTL, profile, result0);
        else if (on.Contains(SheetNames.TimeLayouts))
            result0.Warnings.Add($"工作簿里没有「{SheetNames.TimeLayouts}」工作表，该部分跳过。");

        if (on.Contains(SheetNames.Subjects) && wb.TryGetWorksheet(SheetNames.Subjects, out var wsSub))
            ParseSubjects(wsSub, profile, result0);
        else if (on.Contains(SheetNames.Subjects))
            result0.Warnings.Add($"工作簿里没有「{SheetNames.Subjects}」工作表，该部分跳过。");

        if (on.Contains(SheetNames.ClassPlans) && wb.TryGetWorksheet(SheetNames.ClassPlans, out var wsCP))
            ParseClassPlans(wsCP, profile, result0);
        else if (on.Contains(SheetNames.ClassPlans))
            result0.Warnings.Add($"工作簿里没有「{SheetNames.ClassPlans}」工作表，该部分跳过。");

        if (on.Contains(SheetNames.Settings) && wb.TryGetWorksheet(SheetNames.Settings, out var wsSt))
            ParseSettings(wsSt, settings, result0);
        else if (on.Contains(SheetNames.Settings))
            result0.Warnings.Add($"工作簿里没有「{SheetNames.Settings}」工作表，该部分跳过。");

        // 用最终目标状态对当前档案做差异比对
        CompareProfile(currentProfile, profile, result0.Diff);
        CompareSettings(currentSettings, settings, result0.Diff);

        return new Result
        {
            Profile = profile,
            Settings = settings,
            Diff = result0.Diff
        }.AlsoMergeWarnings(result0.Warnings);
    }

    private static Result AlsoMergeWarnings(this Result r, List<string> warnings)
    {
        r.Warnings.AddRange(warnings);
        return r;
    }

    // ---------- 时间表 ----------

    private static void ParseTimeLayouts(IXLWorksheet ws, JsonObject profile, Result result)
    {
        var header = ReadHeader(ws);
        if (!header.TryGetValue("时间表名称", out var cName))
        {
            result.Warnings.Add($"「{ws.Name}」缺少「时间表名称」列，已跳过。");
            return;
        }
        header.TryGetValue("时间表ID", out var cId);
        header.TryGetValue("是否叠放", out var cOverlay);
        header.TryGetValue("叠放源ID", out var cOverlaySrc);
        header.TryGetValue("是否启用", out var cActive);
        header.TryGetValue("行号", out var cIdx);
        header.TryGetValue("原始行(JSON)", out var cRaw);

        var layouts = profile["TimeLayouts"] as JsonObject;
        if (layouts is null)
        {
            layouts = new JsonObject();
            profile["TimeLayouts"] = layouts;
        }

        // 按 时间表ID（无则按名称）分组，重建每个时间表
        var groups = new Dictionary<string, List<int>>(StringComparer.Ordinal);
        var order = new List<string>();
        var keyById = new Dictionary<string, string>(StringComparer.Ordinal);

        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        for (var r = 2; r <= lastRow; r++)
        {
            var name = Cell(ws, r, cName);
            var idCell = cId > 0 ? Cell(ws, r, cId) : "";

            if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(idCell)) continue;

            var key = !string.IsNullOrWhiteSpace(idCell) ? idCell : name;
            if (string.IsNullOrWhiteSpace(key)) continue;

            if (!groups.TryGetValue(key, out var list))
            {
                list = new List<int>();
                groups[key] = list;
                order.Add(key);
                keyById[key] = !string.IsNullOrWhiteSpace(idCell) ? idCell : "";
            }
            list.Add(r);
        }

        foreach (var key in order)
        {
            var rows = groups[key];
            var first = rows[0];

            var name = Cell(ws, first, cName);
            var id = keyById[key];
            if (string.IsNullOrWhiteSpace(id))
            {
                // 没有 ID —— 尝试按名称匹配已有的，否则新建
                id = FindKeyByName(layouts, name) ?? Guid.NewGuid().ToString();
            }

            var existing = layouts[id] as JsonObject;
            var tl = existing is null ? new JsonObject() : (JsonObject)existing.DeepClone();

            tl["Name"] = name;
            if (cOverlay > 0) tl["IsOverlay"] = Decode(Cell(ws, first, cOverlay));
            if (cOverlaySrc > 0) tl["OverlaySourceId"] = Decode(Cell(ws, first, cOverlaySrc));
            // 写激活状态时沿用原档案里已有的字段名，避免凭空造出 CL 不认识的字段。
            // 原档案用 IsActive；万一某版本用 IsActivated 就跟着用 IsActivated。
            if (cActive > 0)
            {
                var activeKey = tl.ContainsKey("IsActive") ? "IsActive"
                    : tl.ContainsKey("IsActivated") ? "IsActivated"
                    : "IsActive";
                tl[activeKey] = Decode(Cell(ws, first, cActive));
            }

            var items = new JsonArray();

            // 依「行号」排序，保证顺序与 Excel 里看到的一致
            var ordered = cIdx > 0
                ? rows.OrderBy(x => ParseInt(Cell(ws, x, cIdx)) ?? int.MaxValue).ToList()
                : rows;

            foreach (var rowNo in ordered)
            {
                var raw = cRaw > 0 ? Cell(ws, rowNo, cRaw) : "";
                if (string.IsNullOrWhiteSpace(raw))
                {
                    // 该行没有原始 JSON（用户手写的表）：用可见列拼一个
                    var it = new JsonObject
                    {
                        ["StartTime"] = Cell(ws, rowNo, header.GetValueOrDefault("开始时间")),
                        ["EndTime"] = Cell(ws, rowNo, header.GetValueOrDefault("结束时间")),
                    };
                    var tt = Decode(Cell(ws, rowNo, header.GetValueOrDefault("类型")));
                    it["TimeType"] = tt ?? JsonValue.Create(0);
                    items.Add(it);
                    continue;
                }

                var parsed = ExcelEngine.DecodeValue(raw);
                if (parsed is JsonObject obj)
                    items.Add(obj);
            }

            tl["Layouts"] = items;
            layouts[id] = tl;
        }
    }

    // ---------- 科目 ----------

    private static void ParseSubjects(IXLWorksheet ws, JsonObject profile, Result result)
    {
        var header = ReadHeader(ws);
        if (!header.TryGetValue("名称", out var cName))
        {
            result.Warnings.Add($"「{ws.Name}」缺少「名称」列，已跳过。");
            return;
        }
        header.TryGetValue("科目ID", out var cId);
        header.TryGetValue("简称", out var cInit);
        header.TryGetValue("教师", out var cTeacher);
        header.TryGetValue("是否室外", out var cOut);

        var subjects = profile["Subjects"] as JsonObject;
        if (subjects is null)
        {
            subjects = new JsonObject();
            profile["Subjects"] = subjects;
        }

        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        for (var r = 2; r <= lastRow; r++)
        {
            var name = Cell(ws, r, cName);
            var id = cId > 0 ? Cell(ws, r, cId) : "";
            if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(id)) continue;

            if (string.IsNullOrWhiteSpace(id))
                id = FindKeyByName(subjects, name) ?? Guid.NewGuid().ToString();

            var s = subjects[id] as JsonObject;
            var subj = s is null ? new JsonObject() : (JsonObject)s.DeepClone();

            subj["Name"] = name;
            if (cInit > 0) subj["Initial"] = Cell(ws, r, cInit);
            if (cTeacher > 0) subj["TeacherName"] = Cell(ws, r, cTeacher);
            if (cOut > 0) subj["IsOutDoor"] = Decode(Cell(ws, r, cOut));

            subjects[id] = subj;
        }
    }

    // ---------- 课程表 ----------

    private static void ParseClassPlans(IXLWorksheet ws, JsonObject profile, Result result)
    {
        var header = ReadHeader(ws);
        if (!header.TryGetValue("课表名称", out var cName))
        {
            result.Warnings.Add($"「{ws.Name}」缺少「课表名称」列，已跳过。");
            return;
        }
        header.TryGetValue("课表ID", out var cId);
        header.TryGetValue("时间表ID", out var cTl);
        header.TryGetValue("是否叠放", out var cOverlay);
        header.TryGetValue("行号", out var cIdx);
        header.TryGetValue("原始行(JSON)", out var cRaw);

        var plans = profile["ClassPlans"] as JsonObject;
        if (plans is null)
        {
            plans = new JsonObject();
            profile["ClassPlans"] = plans;
        }

        var groups = new Dictionary<string, List<int>>(StringComparer.Ordinal);
        var order = new List<string>();
        var idOf = new Dictionary<string, string>(StringComparer.Ordinal);

        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        for (var r = 2; r <= lastRow; r++)
        {
            var name = Cell(ws, r, cName);
            var idCell = cId > 0 ? Cell(ws, r, cId) : "";
            if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(idCell)) continue;

            var key = !string.IsNullOrWhiteSpace(idCell) ? idCell : name;
            if (string.IsNullOrWhiteSpace(key)) continue;

            if (!groups.TryGetValue(key, out var list))
            {
                list = new List<int>();
                groups[key] = list;
                order.Add(key);
                idOf[key] = idCell;
            }
            list.Add(r);
        }

        foreach (var key in order)
        {
            var rows = groups[key];
            var first = rows[0];

            var name = Cell(ws, first, cName);
            var id = idOf[key];
            if (string.IsNullOrWhiteSpace(id))
                id = FindKeyByName(plans, name) ?? Guid.NewGuid().ToString();

            var existing = plans[id] as JsonObject;
            var cp = existing is null ? new JsonObject() : (JsonObject)existing.DeepClone();

            cp["Name"] = name;
            if (cTl > 0)
            {
                var tl = Cell(ws, first, cTl);
                if (!string.IsNullOrWhiteSpace(tl)) cp["TimeLayoutId"] = tl;
            }
            if (cOverlay > 0) cp["IsOverlay"] = Decode(Cell(ws, first, cOverlay));

            var classes = new JsonArray();

            var ordered = cIdx > 0
                ? rows.OrderBy(x => ParseInt(Cell(ws, x, cIdx)) ?? int.MaxValue).ToList()
                : rows;

            foreach (var rowNo in ordered)
            {
                var raw = cRaw > 0 ? Cell(ws, rowNo, cRaw) : "";
                if (!string.IsNullOrWhiteSpace(raw))
                {
                    var parsed = ExcelEngine.DecodeValue(raw);
                    if (parsed is JsonObject obj) { classes.Add(obj); continue; }
                }

                // 手写表退化路径：靠 科目ID / 索引 拼
                var subId = Cell(ws, rowNo, header.GetValueOrDefault("科目ID"));
                if (string.IsNullOrWhiteSpace(subId)) continue;
                var idxNode = Decode(Cell(ws, rowNo, header.GetValueOrDefault("索引")));
                classes.Add(new JsonObject
                {
                    ["SubjectId"] = subId,
                    ["Index"] = idxNode ?? JsonValue.Create(classes.Count)
                });
            }

            cp["Classes"] = classes;
            plans[id] = cp;
        }
    }

    // ---------- 全局配置 ----------

    private static void ParseSettings(IXLWorksheet ws, JsonObject settings, Result result)
    {
        var header = ReadHeader(ws);
        if (!header.TryGetValue("配置项", out var cKey))
        {
            result.Warnings.Add($"「{ws.Name}」缺少「配置项」列，已跳过。");
            return;
        }
        header.TryGetValue("值", out var cVal);

        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        for (var r = 2; r <= lastRow; r++)
        {
            var key = Cell(ws, r, cKey);
            if (string.IsNullOrWhiteSpace(key)) continue;

            var raw = cVal > 0 ? Cell(ws, r, cVal) : "";
            // 注意：这里不能写 ?? JsonValue.Create("")，否则配置项里的 null 会被改成空串。
            settings[key.Trim()] = ExcelEngine.DecodeValue(raw);
        }
    }

    // ---------- 差异比对 ----------

    private static void CompareProfile(JsonNode? curNode, JsonObject tgt, DiffReport diff)
    {
        var cur = curNode as JsonObject ?? new JsonObject();

        CompareSection(cur["TimeLayouts"] as JsonObject, tgt["TimeLayouts"] as JsonObject,
            SheetNames.TimeLayouts, "时间表", diff, keyField: "Name");

        CompareSection(cur["Subjects"] as JsonObject, tgt["Subjects"] as JsonObject,
            SheetNames.Subjects, "科目", diff, keyField: "Name");

        CompareSection(cur["ClassPlans"] as JsonObject, tgt["ClassPlans"] as JsonObject,
            SheetNames.ClassPlans, "课表", diff, keyField: "Name");
    }

    private static void CompareSection(
        JsonObject? cur, JsonObject? tgt, string sheet, string kind,
        DiffReport diff, string keyField)
    {
        if (tgt is null) return;

        cur ??= new JsonObject();

        foreach (var kv in tgt)
        {
            var name = (kv.Value as JsonObject)?[keyField]?.GetValue<string>() ?? kv.Key;
            var label = string.IsNullOrWhiteSpace(name) ? kv.Key : name;

            if (cur[kv.Key] is null)
            {
                diff.Entries.Add(new DiffEntry
                {
                    Kind = DiffKind.Added, Section = sheet, Target = label,
                    Detail = $"新增{kind}"
                });
            }
            else if (!JsonNode.DeepEquals(cur[kv.Key], kv.Value))
            {
                diff.Entries.Add(new DiffEntry
                {
                    Kind = DiffKind.Modified, Section = sheet, Target = label,
                    Detail = $"{kind}内容有变化"
                });
            }
            else
            {
                diff.Entries.Add(new DiffEntry
                {
                    Kind = DiffKind.Unchanged, Section = sheet, Target = label,
                    Detail = "无变化"
                });
            }
        }

        foreach (var kv in cur)
        {
            if (tgt[kv.Key] is not null) continue;
            var name = (kv.Value as JsonObject)?[keyField]?.GetValue<string>() ?? kv.Key;
            diff.Entries.Add(new DiffEntry
            {
                Kind = DiffKind.Removed, Section = sheet,
                Target = string.IsNullOrWhiteSpace(name) ? kv.Key : name,
                Detail = $"Excel 中不存在（覆盖模式下将被删除）"
            });
        }
    }

    private static void CompareSettings(JsonNode? curNode, JsonObject tgt, DiffReport diff)
    {
        var cur = curNode as JsonObject ?? new JsonObject();

        foreach (var kv in tgt)
        {
            if (!cur.ContainsKey(kv.Key))
            {
                diff.Entries.Add(new DiffEntry
                {
                    Kind = DiffKind.Added, Section = SheetNames.Settings,
                    Target = kv.Key, Detail = $"新增配置项 = {ExcelEngine.EncodeValue(kv.Value)}"
                });
            }
            else if (!JsonNode.DeepEquals(cur[kv.Key], kv.Value))
            {
                diff.Entries.Add(new DiffEntry
                {
                    Kind = DiffKind.Modified, Section = SheetNames.Settings, Target = kv.Key,
                    Detail = $"{ExcelEngine.EncodeValue(cur[kv.Key])} → {ExcelEngine.EncodeValue(kv.Value)}"
                });
            }
            else
            {
                diff.Entries.Add(new DiffEntry
                {
                    Kind = DiffKind.Unchanged, Section = SheetNames.Settings,
                    Target = kv.Key, Detail = "无变化"
                });
            }
        }
    }

    // ---------- 应用策略 ----------

    /// <summary>
    /// 按策略把 Excel 的目标状态合到当前档案上，产出最终要落盘的 JSON。
    /// </summary>
    public static (JsonObject Profile, JsonObject Settings) ApplyStrategy(
        JsonNode currentProfile, JsonNode currentSettings,
        JsonObject targetProfile, JsonObject targetSettings,
        ImportStrategy strategy)
    {
        var finalProfile = strategy switch
        {
            ImportStrategy.Overwrite => (JsonObject)targetProfile.DeepClone(),
            ImportStrategy.Merge => MergeProfile(currentProfile as JsonObject ?? new JsonObject(), targetProfile),
            ImportStrategy.SkipExisting => SkipProfile(currentProfile as JsonObject ?? new JsonObject(), targetProfile),
            _ => (JsonObject)targetProfile.DeepClone()
        };

        var finalSettings = strategy switch
        {
            ImportStrategy.Overwrite => (JsonObject)targetSettings.DeepClone(),
            _ => MergeSettings(currentSettings as JsonObject ?? new JsonObject(), targetSettings)
        };

        return (finalProfile, finalSettings);
    }

    private static JsonObject MergeProfile(JsonObject cur, JsonObject tgt)
    {
        var result = (JsonObject)cur.DeepClone();

        MergeDict(result, tgt, "TimeLayouts");
        MergeDict(result, tgt, "Subjects");
        MergeDict(result, tgt, "ClassPlans");

        foreach (var kv in tgt)
        {
            if (kv.Key is "TimeLayouts" or "Subjects" or "ClassPlans") continue;
            result[kv.Key] = kv.Value?.DeepClone();
        }

        return result;
    }

    private static JsonObject SkipProfile(JsonObject cur, JsonObject tgt)
    {
        var result = (JsonObject)cur.DeepClone();

        SkipDict(result, tgt, "TimeLayouts");
        SkipDict(result, tgt, "Subjects");
        SkipDict(result, tgt, "ClassPlans");

        return result;
    }

    private static void MergeDict(JsonObject cur, JsonObject tgt, string section)
    {
        if (tgt[section] is not JsonObject tgtDict) return;

        if (cur[section] is not JsonObject curDict)
        {
            cur[section] = tgtDict.DeepClone();
            return;
        }

        foreach (var kv in tgtDict)
            curDict[kv.Key] = kv.Value?.DeepClone();
    }

    private static void SkipDict(JsonObject cur, JsonObject tgt, string section)
    {
        if (tgt[section] is not JsonObject tgtDict) return;

        if (cur[section] is not JsonObject curDict)
        {
            cur[section] = tgtDict.DeepClone();
            return;
        }

        foreach (var kv in tgtDict)
            if (!curDict.ContainsKey(kv.Key))
                curDict[kv.Key] = kv.Value?.DeepClone();
    }

    private static JsonObject MergeSettings(JsonObject cur, JsonObject tgt)
    {
        var result = (JsonObject)cur.DeepClone();
        foreach (var kv in tgt)
            result[kv.Key] = kv.Value?.DeepClone();
        return result;
    }

    // ---------- 小工具 ----------

    private static Dictionary<string, int> ReadHeader(IXLWorksheet ws)
    {
        var map = new Dictionary<string, int>(StringComparer.Ordinal);
        var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
        for (var c = 1; c <= lastCol; c++)
        {
            var v = ws.Cell(1, c).GetString().Trim();
            if (!string.IsNullOrEmpty(v) && !map.ContainsKey(v))
                map[v] = c;
        }
        return map;
    }

    private static string Cell(IXLWorksheet ws, int rowNumber, int col)
        => (col <= 0 || rowNumber <= 0) ? "" : ws.Cell(rowNumber, col).GetString();

    private static string? FindKeyByName(JsonObject dict, string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return null;
        foreach (var kv in dict)
            if (kv.Value is JsonObject o &&
                string.Equals(o["Name"]?.GetValue<string>(), name, StringComparison.Ordinal))
                return kv.Key;
        return null;
    }

    private static JsonNode? Decode(string s) => ExcelEngine.DecodeValue(s);

    private static int? ParseInt(string s)
        => int.TryParse(s, out var v) ? v : null;
}

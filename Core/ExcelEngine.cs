using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using ClosedXML.Excel;

namespace ClassIsland.ExcelTool.Core;

/// <summary>
/// Excel 读写引擎。
///
/// 设计要点（关键）：不手写 CL 的模型映射，而是直接以 <see cref="JsonNode"/> 为中间表示。
/// 理由：CL 的 Profile / Settings 结构字段多且版本间会变，逐字段硬编码映射极易漏字段、
/// 一旦漏了就会在导入时把用户档案写坏。这里把「档案 JSON ↔ Excel 单元格」做成无损往返：
///   - 导出的每个单元格都能原样解析回 JSON 值（数字/布尔/字符串/对象/数组都带类型标记）
///   - 导入时按表头定位列，缺列不影响，未知列原样保留
/// 这样即使 CL 以后加了新字段，也不需要改这个文件。
/// </summary>
public static class ExcelEngine
{
    // ---------- 单元格 ↔ JSON 的无损编码 ----------
    // 约定：对象/数组用 "json:" 前缀做类型标记；布尔与数字用原生类型；其余当字符串。

    private const string JsonPrefix = "json:";

    /// <summary>JSON null 的单元格标记，与空字符串 "" 区分。</summary>
    private const string NullMarker = "⊘null";

    internal static string EncodeValue(JsonNode? node)
    {
        // 关键：JSON 的 null 必须与「空字符串」区分开。
        // 早期版本把两者都写成 ""，导入时一律还原成 ""，导致 OverlaySourceId
        // 这类可空字段从 null 被悄悄改成空串（往返比对时被发现）。
        if (node is null) return NullMarker;

        switch (node)
        {
            case JsonObject or JsonArray:
                return JsonPrefix + node.ToJsonString();

            case JsonValue v:
                if (v.TryGetValue<bool>(out var b)) return b ? "TRUE" : "FALSE";
                if (v.TryGetValue<double>(out var d))
                {
                    // 整数不要写成 1.0，避免导入时类型漂移
                    if (Math.Abs(d % 1) < double.Epsilon && Math.Abs(d) < 1e15)
                        return ((long)d).ToString(System.Globalization.CultureInfo.InvariantCulture);
                    return d.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }
                if (v.TryGetValue<string>(out var s)) return s;
                return v.ToJsonString();

            default:
                return node.ToJsonString();
        }
    }

    internal static JsonNode? DecodeValue(string? text)
    {
        // 与 EncodeValue 对称：null 标记还原成 JSON null，真正的空串仍是空串。
        if (text is null) return null;
        if (text == NullMarker) return null;

        if (string.IsNullOrEmpty(text)) return JsonValue.Create("");

        var t = text.Trim();

        if (t.StartsWith(JsonPrefix, StringComparison.Ordinal))
        {
            try { return JsonNode.Parse(t.Substring(JsonPrefix.Length)); }
            catch { return JsonValue.Create(text); }
        }

        if (string.Equals(t, "TRUE", StringComparison.OrdinalIgnoreCase)) return JsonValue.Create(true);
        if (string.Equals(t, "FALSE", StringComparison.OrdinalIgnoreCase)) return JsonValue.Create(false);

        if (long.TryParse(t, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var l))
            return JsonValue.Create(l);

        if (double.TryParse(t, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var d))
            return JsonValue.Create(d);

        return JsonValue.Create(text);
    }

    /// <summary>把 Guid 之类的字符串安全转成 JsonValue。</summary>
    internal static JsonNode? AsNode(object? o) => o switch
    {
        null => null,
        JsonNode n => n,
        string s => JsonValue.Create(s),
        bool b => JsonValue.Create(b),
        int i => JsonValue.Create(i),
        long l => JsonValue.Create(l),
        double d => JsonValue.Create(d),
        Guid g => JsonValue.Create(g.ToString()),
        _ => JsonValue.Create(o.ToString())
    };

    // ---------- 导出 ----------

    /// <summary>
    /// 把档案导出为 Excel 工作簿。
    /// </summary>
    /// <param name="profileJson">Profiles\*.json 的内容</param>
    /// <param name="settingsJson">Settings.json 的内容</param>
    /// <param name="enabledSheets">要包含的 Sheet</param>
    /// <param name="styleHeader">是否加表头样式</param>
    /// <param name="showIds">
    /// 是否写入 ID 列。默认 false —— 那些 GUID 是档案内部主键，导入时按名称匹配即可，
    /// 写出来只会把表格糊满乱码。打开后 ID 会作为「隐藏列」排在所有可见列右边。
    /// </param>
    /// <returns>写好的 XLWorkbook</returns>
    public static XLWorkbook BuildWorkbook(
        JsonNode profileJson,
        JsonNode settingsJson,
        IEnumerable<string> enabledSheets,
        bool styleHeader = true,
        bool showIds = false)
    {
        var wb = new XLWorkbook();
        var on = new HashSet<string>(enabledSheets, StringComparer.Ordinal);

        if (on.Contains(SheetNames.TimeLayouts))
            WriteTimeLayouts(wb.Worksheets.Add(SheetNames.TimeLayouts), profileJson, showIds);

        if (on.Contains(SheetNames.Subjects))
            WriteSubjects(wb.Worksheets.Add(SheetNames.Subjects), profileJson, showIds);

        if (on.Contains(SheetNames.ClassPlans))
            WriteClassPlans(wb.Worksheets.Add(SheetNames.ClassPlans), profileJson, showIds);

        if (on.Contains(SheetNames.ClassPlanMeta))
            WriteClassPlanMeta(wb.Worksheets.Add(SheetNames.ClassPlanMeta), profileJson, showIds);

        if (on.Contains(SheetNames.Settings))
            WriteSettings(wb.Worksheets.Add(SheetNames.Settings), settingsJson);

        if (styleHeader)
            foreach (var ws in wb.Worksheets)
            {
                if (ws.LastRowUsed() is null) continue;
                var head = ws.Row(1);
                head.Style.Font.Bold = true;
                head.Style.Fill.BackgroundColor = XLColor.FromHtml("#F2F2F2");
                ws.SheetView.FreezeRows(1);
                ws.Columns().AdjustToContents(1, 60);
            }

        return wb;
    }

    private static void WriteTimeLayouts(IXLWorksheet ws, JsonNode profile, bool showIds)
    {
        // 可见列：只留人看得懂的内容。
        //   - ID 是 36 位 GUID，且每行重复，只会糊满表格；
        //   - 「原始行(JSON)」是导入用的机械载荷，人不需要读。
        // 两者都放到右侧隐藏列：导入按表头名找列，隐藏不影响读取。
        string[] cols =
        {
            "时间表名称", "是否叠放", "是否启用",
            "行号", "开始时间", "结束时间", "时长(秒)", "类型",
            "隐藏默认", "课间名称", "结束秒", "附加对象(JSON)", "行动集(JSON)"
        };
        for (var i = 0; i < cols.Length; i++) ws.Cell(1, i + 1).Value = cols[i];

        // 隐藏列：原始行永远写；ID 三列仅在 showIds 时写
        const int C_RAW = 14;            // 原始行(JSON)
        const int C_ID = 15;             // 时间表ID
        const int C_OVERLAY_SRC = 16;    // 叠放源ID
        const int C_DEFAULT_CLASS = 17;  // 默认课程ID

        var layouts = profile["TimeLayouts"] as JsonObject;
        if (layouts is null) return;

        var r = 2;
        foreach (var kv in layouts)
        {
            var tl = kv.Value as JsonObject;
            if (tl is null) continue;

            var tlName = tl["Name"]?.GetValue<string>() ?? "";
            var tlId = kv.Key;
            var isOverlay = tl["IsOverlay"];
            var overlaySrc = tl["OverlaySourceId"];
            // 注意：档案里时间表的激活字段序列化名是 IsActive（不是 CL 模型属性名 IsActivated）。
            // 早期版本误用了 IsActivated，导致每次导入都会凭空多出一个字段。
            var isActive = tl["IsActive"] ?? tl["IsActivated"];

            var items = tl["Layouts"] as JsonArray;
            var idx = 0;

            if (items is null || items.Count == 0)
            {
                // 空时间表也要占一行，导入时才知道它存在
                ws.Cell(r, 1).Value = tlName;
                ws.Cell(r, 2).Value = EncodeValue(isOverlay);
                ws.Cell(r, 3).Value = EncodeValue(isActive);
                if (showIds)
                {
                    ws.Cell(r, C_ID).Value = tlId;
                    ws.Cell(r, C_OVERLAY_SRC).Value = EncodeValue(overlaySrc);
                }
                r++;
                continue;
            }

            foreach (var item in items)
            {
                var it = item as JsonObject;
                ws.Cell(r, 1).Value = tlName;
                ws.Cell(r, 2).Value = EncodeValue(isOverlay);
                ws.Cell(r, 3).Value = EncodeValue(isActive);

                ws.Cell(r, 4).Value = idx;
                ws.Cell(r, 5).Value = it?["StartTime"]?.GetValue<string>() ?? "";
                ws.Cell(r, 6).Value = it?["EndTime"]?.GetValue<string>() ?? "";
                ws.Cell(r, 7).Value = it?["Last"]?.GetValue<string>() ?? "";
                ws.Cell(r, 8).Value = EncodeValue(it?["TimeType"]);
                ws.Cell(r, 9).Value = EncodeValue(it?["IsHideDefault"]);
                ws.Cell(r, 10).Value = it?["BreakName"]?.GetValue<string>() ?? "";
                ws.Cell(r, 11).Value = it?["EndSecond"]?.GetValue<string>() ?? "";
                ws.Cell(r, 12).Value = EncodeValue(it?["AttachedObjects"]);
                ws.Cell(r, 13).Value = EncodeValue(it?["ActionSet"]);

                // 原始行：导入的主路径载荷，永远写入但放在隐藏列
                ws.Cell(r, C_RAW).Value = it is null ? "" : JsonPrefix + it.ToJsonString();

                if (showIds)
                {
                    ws.Cell(r, C_ID).Value = tlId;
                    ws.Cell(r, C_OVERLAY_SRC).Value = EncodeValue(overlaySrc);
                    ws.Cell(r, C_DEFAULT_CLASS).Value = EncodeValue(it?["DefaultClassId"]);
                }

                r++;
                idx++;
            }
        }

        // 原始行列：永远存在，隐藏
        ws.Cell(1, C_RAW).Value = "原始行(JSON)";
        HideIdColumn(ws, C_RAW);

        if (showIds)
        {
            ws.Cell(1, C_ID).Value = "时间表ID";
            ws.Cell(1, C_OVERLAY_SRC).Value = "叠放源ID";
            ws.Cell(1, C_DEFAULT_CLASS).Value = "默认课程ID";
            HideIdColumn(ws, C_ID);
            HideIdColumn(ws, C_OVERLAY_SRC);
            HideIdColumn(ws, C_DEFAULT_CLASS);
        }
    }

    /// <summary>把 ID 列整列隐藏并做视觉弱化（灰色表头），避免误导主人去编辑。</summary>
    private static void HideIdColumn(IXLWorksheet ws, int col)
    {
        try { ws.Column(col).Hide(); } catch { /* 某些写出器可能不支持，忽略 */ }
        ws.Cell(1, col).Style.Font.Bold = true;
        ws.Cell(1, col).Style.Font.FontColor = XLColor.FromHtml("#888888");
        ws.Cell(1, col).Style.Fill.BackgroundColor = XLColor.FromHtml("#EFEFEF");
    }

    private static void WriteSubjects(IXLWorksheet ws, JsonNode profile, bool showIds)
    {
        // 可见列只有 4 个，ID 挪到隐藏列
        string[] cols = { "名称", "简称", "教师", "是否室外" };
        for (var i = 0; i < cols.Length; i++) ws.Cell(1, i + 1).Value = cols[i];

        const int C_ID = 5;

        var subjects = profile["Subjects"] as JsonObject;
        if (subjects is null) return;

        var r = 2;
        foreach (var kv in subjects)
        {
            var s = kv.Value as JsonObject;
            ws.Cell(r, 1).Value = s?["Name"]?.GetValue<string>() ?? "";
            ws.Cell(r, 2).Value = s?["Initial"]?.GetValue<string>() ?? "";
            ws.Cell(r, 3).Value = s?["TeacherName"]?.GetValue<string>() ?? "";
            ws.Cell(r, 4).Value = EncodeValue(s?["IsOutDoor"]);
            if (showIds) ws.Cell(r, C_ID).Value = kv.Key;
            r++;
        }

        if (showIds)
        {
            ws.Cell(1, C_ID).Value = "科目ID";
            HideIdColumn(ws, C_ID);
        }
    }

    private static void WriteClassPlans(IXLWorksheet ws, JsonNode profile, bool showIds)
    {
        // 可见列：课表名 / 科目名 在前，方便直接看；ID 与原始行都挪到隐藏列
        string[] cols =
        {
            "课表名称", "科目名", "索引", "是否启用",
            "是否换课", "是否叠放", "行号"
        };
        for (var i = 0; i < cols.Length; i++) ws.Cell(1, i + 1).Value = cols[i];

        const int C_RAW = 8;        // 原始行(JSON)
        const int C_CP_ID = 9;      // 课表ID
        const int C_TL_ID = 10;     // 时间表ID
        const int C_SUB_ID = 11;    // 科目ID

        var plans = profile["ClassPlans"] as JsonObject;
        if (plans is null) return;

        var subjects = profile["Subjects"] as JsonObject;

        var r = 2;
        foreach (var kv in plans)
        {
            var cp = kv.Value as JsonObject;
            if (cp is null) continue;

            var cpName = cp["Name"]?.GetValue<string>() ?? "";
            var tlId = cp["TimeLayoutId"] is JsonValue tv && tv.TryGetValue<string>(out var ts) ? ts : "";
            var isOverlay = cp["IsOverlay"];

            var classes = cp["Classes"] as JsonArray;
            if (classes is null || classes.Count == 0)
            {
                ws.Cell(r, 1).Value = cpName;
                ws.Cell(r, 6).Value = EncodeValue(isOverlay);
                if (showIds)
                {
                    ws.Cell(r, C_CP_ID).Value = kv.Key;
                    ws.Cell(r, C_TL_ID).Value = tlId;
                }
                r++;
                continue;
            }

            var idx = 0;
            foreach (var item in classes)
            {
                var ci = item as JsonObject;
                var subId = ci?["SubjectId"]?.GetValue<string>() ?? "";

                ws.Cell(r, 1).Value = cpName;
                ws.Cell(r, 2).Value = LookupSubjectName(subjects, subId);
                ws.Cell(r, 3).Value = EncodeValue(ci?["Index"]);
                ws.Cell(r, 4).Value = EncodeValue(ci?["IsEnabled"]);
                ws.Cell(r, 5).Value = EncodeValue(ci?["IsChangedClass"]);
                ws.Cell(r, 6).Value = EncodeValue(isOverlay);
                ws.Cell(r, 7).Value = idx;
                ws.Cell(r, C_RAW).Value = ci is null ? "" : JsonPrefix + ci.ToJsonString();

                if (showIds)
                {
                    ws.Cell(r, C_CP_ID).Value = kv.Key;
                    ws.Cell(r, C_TL_ID).Value = tlId;
                    ws.Cell(r, C_SUB_ID).Value = subId;
                }

                r++;
                idx++;
            }
        }

        // 原始行列：永远存在，隐藏
        ws.Cell(1, C_RAW).Value = "原始行(JSON)";
        HideIdColumn(ws, C_RAW);

        if (showIds)
        {
            ws.Cell(1, C_CP_ID).Value = "课表ID";
            ws.Cell(1, C_TL_ID).Value = "时间表ID";
            ws.Cell(1, C_SUB_ID).Value = "科目ID";
            HideIdColumn(ws, C_CP_ID);
            HideIdColumn(ws, C_TL_ID);
            HideIdColumn(ws, C_SUB_ID);
        }
    }

    private static string LookupSubjectName(JsonObject? subjects, string id)
    {
        if (subjects is null || string.IsNullOrEmpty(id)) return "";
        if (subjects[id] is JsonObject s)
            return s["Name"]?.GetValue<string>() ?? "";
        return "";
    }

    private static void WriteClassPlanMeta(IXLWorksheet ws, JsonNode profile, bool showIds)
    {
        // 可见列：全是人能看懂的东西；「关联分组」实际是 GUID，原始行是机械载荷，都移到隐藏区
        string[] cols =
        {
            "名称", "时间表名称", "是否启用", "是否叠放", "课时数"
        };
        for (var i = 0; i < cols.Length; i++) ws.Cell(1, i + 1).Value = cols[i];

        const int C_RAW = 6;          // 原始行(JSON)
        const int C_CP_ID = 7;        // 课表ID
        const int C_TL_ID = 8;        // 时间表ID
        const int C_OVERLAY_SRC = 9;  // 叠放源ID
        const int C_GROUP = 10;       // 关联分组（实际是 GUID）

        var plans = profile["ClassPlans"] as JsonObject;
        var layouts = profile["TimeLayouts"] as JsonObject;
        if (plans is null) return;

        var r = 2;
        foreach (var kv in plans)
        {
            var cp = kv.Value as JsonObject;
            if (cp is null) continue;

            var tlId = cp["TimeLayoutId"] is JsonValue tv && tv.TryGetValue<string>(out var ts) ? ts : "";
            var tlName = "";
            if (layouts?[tlId] is JsonObject tlo)
                tlName = tlo["Name"]?.GetValue<string>() ?? "";

            var count = (cp["Classes"] as JsonArray)?.Count ?? 0;

            ws.Cell(r, 1).Value = cp["Name"]?.GetValue<string>() ?? "";
            ws.Cell(r, 2).Value = tlName;
            ws.Cell(r, 3).Value = EncodeValue(cp["IsEnabled"]);
            ws.Cell(r, 4).Value = EncodeValue(cp["IsOverlay"]);
            ws.Cell(r, 5).Value = count;
            ws.Cell(r, C_RAW).Value = JsonPrefix + cp.ToJsonString();

            if (showIds)
            {
                ws.Cell(r, C_CP_ID).Value = kv.Key;
                ws.Cell(r, C_TL_ID).Value = tlId;
                ws.Cell(r, C_OVERLAY_SRC).Value = EncodeValue(cp["OverlaySourceId"]);
                ws.Cell(r, C_GROUP).Value = EncodeValue(cp["AssociatedGroup"]);
            }
            r++;
        }

        // 原始行列：永远存在，隐藏
        ws.Cell(1, C_RAW).Value = "原始行(JSON)";
        HideIdColumn(ws, C_RAW);

        if (showIds)
        {
            ws.Cell(1, C_CP_ID).Value = "课表ID";
            ws.Cell(1, C_TL_ID).Value = "时间表ID";
            ws.Cell(1, C_OVERLAY_SRC).Value = "叠放源ID";
            ws.Cell(1, C_GROUP).Value = "关联分组";
            HideIdColumn(ws, C_CP_ID);
            HideIdColumn(ws, C_TL_ID);
            HideIdColumn(ws, C_OVERLAY_SRC);
            HideIdColumn(ws, C_GROUP);
        }
    }

    private static void WriteSettings(IXLWorksheet ws, JsonNode settings)
    {
        string[] cols = { "配置项", "值", "类型", "说明" };
        for (var i = 0; i < cols.Length; i++) ws.Cell(1, i + 1).Value = cols[i];

        if (settings is not JsonObject obj) return;

        var r = 2;
        foreach (var kv in obj)
        {
            var kind = kv.Value switch
            {
                JsonObject or JsonArray => "复合",
                JsonValue v when v.TryGetValue<bool>(out _) => "布尔",
                JsonValue v when v.TryGetValue<double>(out _) => "数字",
                null => "空",
                _ => "文本"
            };

            ws.Cell(r, 1).Value = kv.Key;
            ws.Cell(r, 2).Value = EncodeValue(kv.Value);
            ws.Cell(r, 3).Value = kind;
            ws.Cell(r, 4).Value = SettingHints.GetValueOrDefault(kv.Key, "");
            r++;
        }
    }

    /// <summary>常见全局配置项的中文说明，让 Excel 里看得懂。</summary>
    private static readonly Dictionary<string, string> SettingHints = new()
    {
        ["SelectedProfile"] = "当前使用的档案文件名",
        ["IsMainWindowVisible"] = "主窗口是否可见",
        ["SingleWeekStartTime"] = "单双周起始时间",
        ["MultiWeekRotationOffset"] = "多周轮换偏移",
        ["MultiWeekRotationMaxCycle"] = "多周轮换最大周期",
        ["ClassPrepareNotifySeconds"] = "上课预备提醒秒数",
        ["IsClassPrepareNotificationEnabled"] = "是否启用上课预备提醒",
        ["ShowDate"] = "是否显示日期",
        ["HideOnClass"] = "上课时自动隐藏",
        ["IsClassChangingNotificationEnabled"] = "是否启用换课提醒",
        ["IsClassOffNotificationEnabled"] = "是否启用下课提醒",
        ["HideMode"] = "隐藏模式",
        ["HideOnFullscreen"] = "全屏时自动隐藏",
        ["HideOnMaxWindow"] = "最大化时自动隐藏",
        ["IsCountdownEnabled"] = "是否启用倒计时",
        ["CountdownSeconds"] = "倒计时秒数",
        ["ExactTimeServer"] = "对时服务器",
        ["IsExactTimeEnabled"] = "是否启用精确对时",
        ["AnimationLevel"] = "动画级别",
        ["IsSplashEnabled"] = "是否启用启动画面",
        ["SplashCustomText"] = "启动画面自定义文字",
        ["WeatherLongitude"] = "天气经度",
        ["WeatherLatitude"] = "天气纬度",
    };
}

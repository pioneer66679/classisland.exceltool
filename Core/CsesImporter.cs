using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Text.Json.Nodes;

namespace ClassIsland.ExcelTool.Core;

/// <summary>
/// CSES（Course Schedule Exchange Schema）导入引擎。
///
/// CSES 是社区通用的课程表交换格式，ClassIsland 官方 README 点名支持。
/// 规范：https://github.com/SmartTeachCN/CSES
///
/// 结构（YAML）：
///   version: 1
///   subjects:
///     - name: 数学
///       simplified_name: 数        # 可选
///       teacher: 李梅              # 可选
///       room: "101"                # 可选
///   schedules:
///     - name: 星期一
///       enable_day: 1              # 1-7 = 周一到周日
///       weeks: all                 # all / odd / even
///       classes:
///         - subject: 数学
///           start_time: "08:00:00" # HH:MM:SS
///           end_time: "09:00:00"
///
/// 映射到 ClassIsland：
///   subjects[].name              -> Subject.Name
///   subjects[].simplified_name   -> Subject.Initial
///   schedules[].enable_day       -> ClassPlan.TimeRule.WeekDay
///   schedules[].weeks            -> ClassPlan.TimeRule.WeekCountDiv（0=all, 1=odd, 2=even）
///   classes[].start_time/end_time-> 按时间匹配到 TimeLayout 的时间点，定位这是第几节课
///   classes[].subject            -> ClassPlan.Classes[i].SubjectId
/// </summary>
public static class CsesImporter
{
    /// <summary>CSES 里的一天对应 ClassPlan 的时间规则。</summary>
    private sealed class DaySchedule
    {
        public string Name { get; init; } = "";
        public int EnableDay { get; init; }
        public string Weeks { get; init; } = "all";
        public List<CsesClass> Classes { get; } = new();
    }

    private sealed class CsesClass
    {
        public string Subject { get; init; } = "";
        public string StartTime { get; init; } = "";
        public string EndTime { get; init; } = "";
    }

    private sealed class CsesSubject
    {
        public string Name { get; init; } = "";
        public string SimplifiedName { get; init; } = "";
        public string Teacher { get; init; } = "";
    }

    /// <summary>
    /// 解析 CSES 文件内容，得到一份可直接合并进档案的 Profile 片段。
    /// </summary>
    /// <param name="yamlText">CSES 文件（.yml/.yaml）的文本内容</param>
    /// <param name="currentProfile">当前档案，用于复用已有的科目/时间表/课表 ID</param>
    /// <param name="warnings">解析过程中的提示信息</param>
    public static JsonObject Parse(string yamlText, JsonObject currentProfile, List<string> warnings)
    {
        var root = MiniYaml.Parse(yamlText);
        if (root is not Dictionary<string, object?> doc)
        {
            warnings.Add("CSES 解析失败：根节点不是对象。");
            return new JsonObject();
        }

        // ---- version 校验 ----
        var version = GetInt(doc, "version");
        if (version != 1)
            warnings.Add($"CSES 版本号为 {version?.ToString() ?? "(缺失)"}，本插件按 v1 解析，可能不兼容。");

        // ---- subjects ----
        var csesSubjects = new List<CsesSubject>();
        foreach (var item in GetList(doc, "subjects"))
        {
            if (item is not Dictionary<string, object?> s) continue;
            var name = GetStr(s, "name");
            if (string.IsNullOrWhiteSpace(name)) continue;
            csesSubjects.Add(new CsesSubject
            {
                Name = name,
                SimplifiedName = GetStr(s, "simplified_name"),
                Teacher = GetStr(s, "teacher")
            });
        }

        // ---- schedules ----
        var schedules = new List<DaySchedule>();
        foreach (var item in GetList(doc, "schedules"))
        {
            if (item is not Dictionary<string, object?> s) continue;
            var ds = new DaySchedule
            {
                Name = GetStr(s, "name"),
                EnableDay = GetInt(s, "enable_day") ?? 0,
                Weeks = (GetStr(s, "weeks") ?? "all").Trim().ToLowerInvariant()
            };
            if (ds.EnableDay < 1 || ds.EnableDay > 7)
            {
                warnings.Add($"跳过一节课程表：enable_day={ds.EnableDay} 不在 1-7 范围内。");
                continue;
            }
            foreach (var c in GetList(s, "classes"))
            {
                if (c is not Dictionary<string, object?> co) continue;
                ds.Classes.Add(new CsesClass
                {
                    Subject = GetStr(co, "subject"),
                    StartTime = NormalizeTime(GetRaw(co, "start_time"), warnings),
                    EndTime = NormalizeTime(GetRaw(co, "end_time"), warnings)
                });
            }
            schedules.Add(ds);
        }

        if (schedules.Count == 0)
            warnings.Add("CSES 里没有任何 schedules，无事可做。");

        // ---- 写出 Profile 片段 ----
        var profile = new JsonObject();
        var outSubjects = new JsonObject();
        profile["Subjects"] = outSubjects;
        var outLayouts = new JsonObject();
        profile["TimeLayouts"] = outLayouts;
        var outPlans = new JsonObject();
        profile["ClassPlans"] = outPlans;

        // 现有档案里的同名科目，方便复用 ID
        var curSubjects = currentProfile["Subjects"] as JsonObject;
        var curLayouts = currentProfile["TimeLayouts"] as JsonObject;
        var curPlans = currentProfile["ClassPlans"] as JsonObject;

        // 1) 科目：先按 name 找已有的，没有就新建
        var subjectIdByName = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var s in csesSubjects)
        {
            var existing = FindKeyByName(curSubjects, s.Name);
            var id = existing ?? Guid.NewGuid().ToString();
            subjectIdByName[s.Name] = id;

            var node = new JsonObject
            {
                ["Name"] = s.Name,
                ["Initial"] = string.IsNullOrEmpty(s.SimplifiedName) ? s.Name : s.SimplifiedName,
                ["TeacherName"] = s.Teacher,
                ["IsOutDoor"] = false
            };
            outSubjects[id] = node;
        }

        // 2) 时间表：从所有 schedule 的 classes 里归纳出时间点
        //
        // CSES 只给了每节课的起止时间，而 ClassIsland 的时间表是「第 N 节 + 该节的起止」。
        // 做法：把所有 (start,end) 去重排序，就是这套课表的节次序列。
        // 若档案里已有同名时间表且节次一致，则复用其 ID；否则新建一个。
        var allSlots = schedules
            .SelectMany(x => x.Classes)
            .Where(c => !string.IsNullOrEmpty(c.StartTime) && !string.IsNullOrEmpty(c.EndTime))
            .Select(c => (Start: c.StartTime, End: c.EndTime))
            .Distinct()
            .OrderBy(x => x.Start, StringComparer.Ordinal)
            .ToList();

        // 用第一个课表的名称作为时间表名（CSES 没有单独的时间表概念）
        var layoutName = schedules.Count > 0 ? schedules[0].Name : "CSES 时间表";
        if (string.IsNullOrWhiteSpace(layoutName)) layoutName = "CSES 时间表";

        var existingLayout = FindKeyByName(curLayouts, layoutName);
        var layoutId = existingLayout ?? Guid.NewGuid().ToString();

        var layoutItems = new JsonArray();
        foreach (var (start, end) in allSlots)
        {
            layoutItems.Add(new JsonObject
            {
                ["StartSecond"] = "",
                ["EndSecond"] = "",
                ["StartTime"] = start,
                ["EndTime"] = end,
                ["TimeType"] = 0,
                ["IsHideDefault"] = false,
                ["DefaultClassId"] = null,
                ["BreakName"] = "",
                ["ActionSet"] = null,
                ["AttachedObjects"] = new JsonObject(),
                ["IsActive"] = false
            });
        }

        outLayouts[layoutId] = new JsonObject
        {
            ["IsOverlay"] = false,
            ["OverlaySourceId"] = null,
            ["Name"] = layoutName,
            ["Layouts"] = layoutItems,
            ["IsActive"] = schedules.Any(x => x.Weeks == "all"),
            ["IsActivated"] = schedules.Any(x => x.Weeks == "all")
        };

        // 3) 课表：一个 schedule -> 一个 ClassPlan
        foreach (var ds in schedules)
        {
            var planName = string.IsNullOrWhiteSpace(ds.Name) ? "CSES 课表" : ds.Name;

            // 同名同规则的已有课表 -> 复用 ID
            var planId = FindPlanByRule(curPlans, planName, ds.EnableDay, ds.Weeks) ?? Guid.NewGuid().ToString();

            var classes = new JsonArray();
            foreach (var c in ds.Classes)
            {
                if (!subjectIdByName.TryGetValue(c.Subject, out var sid))
                {
                    // CSES 允许 classes 里出现 subjects 中未声明的科目 —— 补一个
                    var fallbackId = FindKeyByName(curSubjects, c.Subject) ?? Guid.NewGuid().ToString();
                    subjectIdByName[c.Subject] = fallbackId;
                    outSubjects[fallbackId] = new JsonObject
                    {
                        ["Name"] = c.Subject,
                        ["Initial"] = c.Subject.Length > 0 ? c.Subject.Substring(0, 1) : "",
                        ["TeacherName"] = "",
                        ["IsOutDoor"] = false
                    };
                    warnings.Add($"CSES 的 classes 里出现了未在 subjects 声明的科目「{c.Subject}」，已自动补上。");
                    sid = fallbackId;
                }

                classes.Add(new JsonObject
                {
                    ["SubjectId"] = sid,
                    ["IsChangedClass"] = false,
                    ["IsEnabled"] = true,
                    ["AttachedObjects"] = new JsonObject(),
                    ["IsActive"] = false
                });
            }

            var timeRule = new JsonObject
            {
                ["WeekDay"] = ds.EnableDay,
                ["WeekCountDiv"] = ds.Weeks switch { "odd" => 1, "even" => 2, _ => 0 },
                ["WeekCountDivTotal"] = 2,
                ["IsActive"] = false
            };

            outPlans[planId] = new JsonObject
            {
                ["TimeLayoutId"] = layoutId,
                ["TimeRule"] = timeRule.ToJsonString(),
                ["Classes"] = classes,
                ["Name"] = planName,
                ["IsOverlay"] = false,
                ["OverlaySourceId"] = null,
                ["OverlaySetupTime"] = DateTime.Now.ToString("o"),
                ["IsEnabled"] = true,
                ["AssociatedGroup"] = null,
                ["AttachedObjects"] = new JsonObject(),
                ["IsActive"] = false
            };
        }

        return profile;
    }

    // ---------- 辅助 ----------

    /// <summary>
    /// 把 CSES 的时间值规范成 "HH:MM:SS"。
    ///
    /// 规范要求是字符串，但实际文件里常见两种偏差：
    ///   1. YAML 解析成整数（例如 36600）—— 那是「当日秒数」，换算回时间；
    ///   2. 形如 "08:00"（少了秒）。
    /// 两种都兼容，并给出提示。
    /// </summary>
    private static string NormalizeTime(object? raw, List<string> warnings)
    {
        switch (raw)
        {
            case null:
                return "";

            case string s:
                s = s.Trim();
                if (s.Length == 0) return "";
                if (Regex.IsMatch(s, @"^\d{2}:\d{2}:\d{2}$")) return s;
                if (Regex.IsMatch(s, @"^\d{1,2}:\d{2}$"))
                {
                    var parts = s.Split(':');
                    return $"{int.Parse(parts[0]):D2}:{int.Parse(parts[1]):D2}:00";
                }
                if (int.TryParse(s, out var sec)) return SecondsToTime(sec);
                warnings.Add($"无法解析的时间值「{s}」，已跳过。");
                return "";

            case int i:
                return SecondsToTime(i);

            case long l:
                return SecondsToTime((int)l);

            case double d:
                return SecondsToTime((int)d);

            default:
                var str = raw.ToString() ?? "";
                warnings.Add($"无法解析的时间值「{str}」，已跳过。");
                return "";
        }
    }

    private static string SecondsToTime(int seconds)
    {
        if (seconds < 0 || seconds >= 86400) return "";
        return TimeSpan.FromSeconds(seconds).ToString(@"hh\:mm\:ss");
    }

    /// <summary>在 JsonObject 里按 Name 找已有的键（GUID）。</summary>
    private static string? FindKeyByName(JsonObject? obj, string name)
    {
        if (obj is null || string.IsNullOrWhiteSpace(name)) return null;
        foreach (var kv in obj)
        {
            if (kv.Value is JsonObject o &&
                string.Equals(o["Name"]?.GetValue<string>(), name, StringComparison.Ordinal))
                return kv.Key;
        }
        return null;
    }

    /// <summary>按「名称 + 星期 + 周次」找已有的课表键，用于复用 ID。</summary>
    private static string? FindPlanByRule(JsonObject? plans, string name, int weekDay, string weeks)
    {
        if (plans is null) return null;
        var wantDiv = weeks switch { "odd" => 1, "even" => 2, _ => 0 };

        foreach (var kv in plans)
        {
            if (kv.Value is not JsonObject p) continue;
            if (!string.Equals(AsString(p["Name"]), name, StringComparison.Ordinal)) continue;

            // 档案里的 TimeRule 可能是字符串（CL 序列化时转成 JSON 文本），也可能是对象
            var rule = AsObject(p["TimeRule"]);
            if (rule is null) continue;

            if (AsInt(rule["WeekDay"]) == weekDay && AsInt(rule["WeekCountDiv"]) == wantDiv)
                return kv.Key;
        }
        return null;
    }

    /// <summary>安全地把 JsonNode 读成字符串（不是 JsonValue 或类型不符时返回 null）。</summary>
    private static string? AsString(JsonNode? node)
    {
        if (node is not JsonValue v) return null;
        try { return v.TryGetValue<string>(out var s) ? s : null; }
        catch { return null; }
    }

    /// <summary>安全地把 JsonNode 读成整数。</summary>
    private static int? AsInt(JsonNode? node)
    {
        if (node is not JsonValue v) return null;
        try
        {
            if (v.TryGetValue<int>(out var i)) return i;
            if (v.TryGetValue<long>(out var l)) return (int)l;
            if (v.TryGetValue<string>(out var s) && int.TryParse(s, out var p)) return p;
            if (v.TryGetValue<double>(out var d)) return (int)d;
        }
        catch { /* 类型不符 */ }
        return null;
    }

    /// <summary>把 JsonNode 读成对象：本身是对象就直接用，是字符串则尝试再解析一次。</summary>
    private static JsonObject? AsObject(JsonNode? node)
    {
        if (node is JsonObject o) return o;
        var text = AsString(node);
        if (string.IsNullOrWhiteSpace(text)) return null;
        try { return JsonNode.Parse(text) as JsonObject; }
        catch { return null; }
    }

    // ---------- YAML 取值 ----------

    private static object? GetRaw(Dictionary<string, object?> d, string key)
        => d.TryGetValue(key, out var v) ? v : null;

    private static string GetStr(Dictionary<string, object?> d, string key)
    {
        var v = GetRaw(d, key);
        return v switch
        {
            null => "",
            string s => s.Trim(),
            _ => v.ToString()?.Trim() ?? ""
        };
    }

    private static int? GetInt(Dictionary<string, object?> d, string key)
    {
        var v = GetRaw(d, key);
        return v switch
        {
            int i => i,
            long l => (int)l,
            string s when int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var r) => r,
            _ => null
        };
    }

    private static IEnumerable<object?> GetList(Dictionary<string, object?> d, string key)
    {
        var v = GetRaw(d, key);
        return v as List<object?> ?? new List<object?>();
    }
}

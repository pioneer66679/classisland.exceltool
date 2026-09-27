using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ClassIsland.ExcelTool.Core;

/// <summary>
/// 极简 YAML 解析器 —— 只覆盖 CSES 用得到的那一小撮语法。
///
/// 为什么自己写而不引 YamlDotNet：
///   插件依赖越少越好，多一个包就多一份与主程序冲突的可能；
///   而 CSES 的结构非常规整（映射 + 列表 + 标量），用不着完整 YAML 实现。
///
/// 支持：
///   key: value              标量
///   key:                    嵌套映射
///   - item                  列表项
///   - key: value            列表里的映射
///   # 注释                  整行或行尾
///   'text' / "text"         引号字符串
///   数字（含负数）、true/false/null
///
/// 不支持（CSES 也用不到）：
///   多文档（---）、锚点与别名（&/*）、折叠/字面块（| >）、流式写法（{} []）
/// </summary>
internal static class MiniYaml
{
    public static object? Parse(string text)
    {
        var lines = new List<(int Indent, string Content)>();
        foreach (var raw in text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'))
        {
            var line = StripComment(raw);
            if (string.IsNullOrWhiteSpace(line)) continue;
            // 制表符会让缩进计算错乱，先换成空格
            line = line.Replace("\t", "    ");
            var indent = 0;
            while (indent < line.Length && line[indent] == ' ') indent++;
            lines.Add((indent, line.Substring(indent).TrimEnd()));
        }

        var pos = 0;
        return ParseBlock(lines, ref pos, 0);
    }

    /// <summary>去掉行尾注释（引号内的 # 不算）。</summary>
    private static string StripComment(string line)
    {
        var inSingle = false;
        var inDouble = false;
        for (var i = 0; i < line.Length; i++)
        {
            var c = line[i];
            if (c == '\'' && !inDouble) inSingle = !inSingle;
            else if (c == '"' && !inSingle) inDouble = !inDouble;
            else if (c == '#' && !inSingle && !inDouble)
            {
                // # 前面必须是行首或空格，才算注释
                if (i == 0 || line[i - 1] == ' ') return line.Substring(0, i);
            }
        }
        return line;
    }

    private static object? ParseBlock(List<(int Indent, string Content)> lines, ref int pos, int minIndent)
    {
        if (pos >= lines.Count) return null;

        var firstIndent = lines[pos].Indent;
        if (firstIndent < minIndent) return null;

        // 列表
        if (lines[pos].Content.StartsWith("- ", StringComparison.Ordinal) ||
            lines[pos].Content == "-")
            return ParseList(lines, ref pos, firstIndent);

        // 映射
        return ParseMap(lines, ref pos, firstIndent);
    }

    private static Dictionary<string, object?> ParseMap(
        List<(int Indent, string Content)> lines, ref int pos, int indent)
    {
        var map = new Dictionary<string, object?>(StringComparer.Ordinal);

        while (pos < lines.Count)
        {
            var (ind, content) = lines[pos];
            if (ind < indent) break;
            if (ind > indent) break;                 // 交给上层处理
            if (content.StartsWith("- ", StringComparison.Ordinal)) break;

            var colon = FindColon(content);
            if (colon < 0)
            {
                pos++;
                continue;                            // 不是键值对，跳过
            }

            var key = content.Substring(0, colon).Trim();
            var rest = content.Substring(colon + 1).Trim();
            pos++;

            if (rest.Length > 0)
            {
                map[key] = ParseScalar(rest);
            }
            else
            {
                // 值在下面的块里。
                //
                // 注意缩进有两种合法写法，都要认：
                //   subjects:            subjects:
                //     - name: 语文         - name: 语文        ← 列表与父键【同缩进】
                //     - name: 数学         - name: 数学
                // 后者（同缩进）在 CSES 的实际文件里很常见。
                if (pos < lines.Count && lines[pos].Indent > indent)
                {
                    map[key] = ParseBlock(lines, ref pos, lines[pos].Indent);
                }
                else if (pos < lines.Count && lines[pos].Indent == indent &&
                         IsListStart(lines[pos].Content))
                {
                    map[key] = ParseList(lines, ref pos, indent);
                }
                else
                {
                    map[key] = null;
                }
            }
        }

        return map;
    }

    private static bool IsListStart(string content)
        => content.StartsWith("- ", StringComparison.Ordinal) || content == "-";

    private static List<object?> ParseList(
        List<(int Indent, string Content)> lines, ref int pos, int indent)
    {
        var list = new List<object?>();

        while (pos < lines.Count)
        {
            var (ind, content) = lines[pos];
            if (ind < indent) break;
            if (ind > indent) break;
            if (!content.StartsWith("- ", StringComparison.Ordinal) && content != "-") break;

            var body = content.Length > 1 ? content.Substring(2).Trim() : "";
            pos++;

            if (body.Length == 0)
            {
                // "-" 独占一行，值在下面
                list.Add(pos < lines.Count && lines[pos].Indent > indent
                    ? ParseBlock(lines, ref pos, lines[pos].Indent)
                    : null);
                continue;
            }

            var colon = FindColon(body);
            if (colon < 0)
            {
                // 纯标量项
                list.Add(ParseScalar(body));
                continue;
            }

            // 形如 "- key: value"。CSS/YAML 里这种列表项【本身就是一个映射】，
            // 后续凡是缩进 >= 列表项内容缩进（即 indent+2）的行，都属于这个映射 ——
            // 不管是排在首个键前面还是后面。
            // 例：
            //   - name: 新课表      <- 列表项，内容缩进 = indent+2
            //     classes:          <- 同一个映射的键
            //     - subject: 早读   <- classes 的列表项（与其父键同缩进）
            //     enable_day: 1     <- 仍是同一个映射的键（可以出现在 classes 之后）
            var map = new Dictionary<string, object?>(StringComparer.Ordinal);
            var memIndent = indent + 2;

            var key = body.Substring(0, colon).Trim();
            var rest = body.Substring(colon + 1).Trim();
            map[key] = rest.Length > 0 ? ParseScalar(rest) : null;

            // 依次吃掉本列表项的其余键
            while (pos < lines.Count && lines[pos].Indent >= memIndent)
            {
                var (ind2, c2) = lines[pos];

                // 遇到下一个列表项（缩进回到 indent 且以 "- " 开头）就结束本项
                if (ind2 <= indent && IsListStart(c2)) break;

                var col2 = FindColon(c2);
                if (col2 < 0) { pos++; continue; }   // 不是键值对，跳过

                var k2 = c2.Substring(0, col2).Trim();
                var v2 = c2.Substring(col2 + 1).Trim();
                pos++;

                if (v2.Length > 0)
                {
                    map[k2] = ParseScalar(v2);
                }
                else
                {
                    // 键后面是子块。两种缩进都要认：严格缩进，或与父键同缩进的列表。
                    if (pos < lines.Count && lines[pos].Indent > ind2)
                        map[k2] = ParseBlock(lines, ref pos, lines[pos].Indent);
                    else if (pos < lines.Count && lines[pos].Indent == ind2 && IsListStart(lines[pos].Content))
                        map[k2] = ParseList(lines, ref pos, ind2);
                    else
                        map[k2] = null;
                }
            }

            list.Add(map);
        }

        return list;
    }

    /// <summary>找键值分隔的冒号（跳过引号内的冒号，且要求冒号后是空格或行尾）。</summary>
    private static int FindColon(string s)
    {
        var inSingle = false;
        var inDouble = false;
        for (var i = 0; i < s.Length; i++)
        {
            var c = s[i];
            if (c == '\'' && !inDouble) inSingle = !inSingle;
            else if (c == '"' && !inSingle) inDouble = !inDouble;
            else if (c == ':' && !inSingle && !inDouble)
            {
                if (i + 1 >= s.Length || s[i + 1] == ' ') return i;
            }
        }
        return -1;
    }

    private static object? ParseScalar(string s)
    {
        s = s.Trim();
        if (s.Length == 0) return "";

        // 字符串字面量也可以带引号
        if (s.Length >= 2 &&
            ((s[0] == '"' && s[^1] == '"') || (s[0] == '\'' && s[^1] == '\'')))
            return s.Substring(1, s.Length - 2);

        if (s is "null" or "~" or "Null" or "NULL") return null;

        var low = s.ToLowerInvariant();
        if (low == "true") return true;
        if (low == "false") return false;

        // 整数（CSES 里 start_time 可能被写成秒数）
        if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var i)) return i;
        if (long.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var l)) return l;
        if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var d)) return d;

        return s;
    }
}

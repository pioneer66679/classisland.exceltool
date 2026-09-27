using System;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

namespace ClassIsland.ExcelTool.Core;

/// <summary>
/// 配置序列化的薄封装，隔离 System.Text.Json 的版本差异与源生成依赖。
/// </summary>
internal static class JsonSerializerCompat
{
    private static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.Never,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        PropertyNameCaseInsensitive = true,
        Converters = { new JsonStringEnumConverter() }
    };

    public static T? Deserialize<T>(JsonNode? node) where T : class
    {
        if (node is null) return null;
        try
        {
            return node.Deserialize<T>(Options);
        }
        catch
        {
            return null;
        }
    }

    public static string Serialize<T>(T value) => JsonSerializer.Serialize(value, Options);
}

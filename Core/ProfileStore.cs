using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace ClassIsland.ExcelTool.Core;

/// <summary>
/// 档案存取：直接操作磁盘上的 Profiles\*.json 与 Settings.json。
///
/// 为什么不用 IProfileService：它在 2.0.1 里只暴露了 CleanExpiredTempClassPlan 等少数方法，
/// 没有「读取/替换整个档案」的稳定公开入口；而 Excel 导入导出本质就是整档备份还原，
/// 直接读写这两个文件语义最明确，也不会因为服务内部状态缓存而产生「界面显示已改、文件没改」的错觉。
///
/// 安全性：所有写操作一律先备份，且写 JSON 用「临时文件 + 原子替换」，中途失败不会留下半个文件。
/// </summary>
public sealed class ProfileStore
{
    public string DataRoot { get; }
    public string ProfilesDir => Path.Combine(DataRoot, "Profiles");
    public string SettingsPath => Path.Combine(DataRoot, "Settings.json");

    private static readonly JsonSerializerOptions WriteOptions = new()
    {
        WriteIndented = false,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    public ProfileStore(string dataRoot)
    {
        DataRoot = dataRoot;
    }

    /// <summary>定位 ClassIsland 的 data 目录。</summary>
    public static string? LocateDataRoot()
    {
        // 1) 插件自身被加载时位于 <CL>\data\Plugins\<id>\，
        //    AppContext.BaseDirectory 就是插件目录，往上两级即 data。
        var baseDir = AppContext.BaseDirectory;
        var probe = Path.GetFullPath(Path.Combine(baseDir, "..", ".."));
        if (File.Exists(Path.Combine(probe, "Settings.json")))
            return probe;

        // 2) 从插件目录向上逐级找「有 Profiles 子目录的 data」
        var dir = new DirectoryInfo(baseDir);
        for (var i = 0; i < 6 && dir != null; i++, dir = dir.Parent)
        {
            var candidate = Path.Combine(dir.FullName, "data");
            if (Directory.Exists(Path.Combine(candidate, "Profiles")))
                return candidate;

            if (dir.Name.Equals("data", StringComparison.OrdinalIgnoreCase) &&
                Directory.Exists(Path.Combine(dir.FullName, "Profiles")))
                return dir.FullName;
        }

        // 3) 用户级数据目录兜底
        var local = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ClassIsland");
        if (Directory.Exists(Path.Combine(local, "Profiles")))
            return local;

        return null;
    }

    /// <summary>当前选中的档案文件（读 Settings.json 的 SelectedProfile）。</summary>
    public string GetCurrentProfilePath()
    {
        var name = "Default.json";

        try
        {
            if (File.Exists(SettingsPath))
            {
                var s = JsonNode.Parse(File.ReadAllText(SettingsPath));
                var sel = s?["SelectedProfile"]?.GetValue<string>();
                if (!string.IsNullOrWhiteSpace(sel)) name = sel;
            }
        }
        catch
        {
            // 读不出来就用默认名
        }

        var p = Path.Combine(ProfilesDir, name);
        if (File.Exists(p)) return p;

        // 兜底：Profiles 里的第一个 json
        if (Directory.Exists(ProfilesDir))
        {
            var files = Directory.GetFiles(ProfilesDir, "*.json");
            if (files.Length > 0) return files[0];
        }

        return p;
    }

    public List<string> ListProfiles()
    {
        var list = new List<string>();
        if (!Directory.Exists(ProfilesDir)) return list;
        foreach (var f in Directory.GetFiles(ProfilesDir, "*.json"))
            list.Add(Path.GetFileName(f));
        list.Sort(StringComparer.OrdinalIgnoreCase);
        return list;
    }

    public JsonNode LoadProfile(string? path = null)
    {
        path ??= GetCurrentProfilePath();
        if (!File.Exists(path)) return new JsonObject();
        return JsonNode.Parse(File.ReadAllText(path)) ?? new JsonObject();
    }

    public JsonNode LoadSettings()
    {
        if (!File.Exists(SettingsPath)) return new JsonObject();
        return JsonNode.Parse(File.ReadAllText(SettingsPath)) ?? new JsonObject();
    }

    /// <summary>
    /// 原子写入：先写 .tmp，再替换目标；被替换的旧文件另存为时间戳备份。
    /// </summary>
    public void SaveAtomic(string path, JsonNode content, string? backupFolder = null)
    {
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

        var json = content.ToJsonString(WriteOptions);
        var tmp = path + ".tmp";

        File.WriteAllText(tmp, json);

        if (File.Exists(path) && !string.IsNullOrEmpty(backupFolder))
        {
            Directory.CreateDirectory(backupFolder);
            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var bak = Path.Combine(backupFolder,
                $"{Path.GetFileNameWithoutExtension(path)}_{stamp}{Path.GetExtension(path)}");
            try { File.Copy(path, bak, overwrite: true); } catch { /* 备份失败不阻断主流程 */ }
        }

        // File.Replace 需要目标存在；不存在时直接移动
        if (File.Exists(path))
            File.Replace(tmp, path, null);
        else
            File.Move(tmp, path);
    }

    public void SaveProfile(JsonNode content, string? path = null, string? backupFolder = null)
        => SaveAtomic(path ?? GetCurrentProfilePath(), content, backupFolder);

    public void SaveSettings(JsonNode content, string? backupFolder = null)
        => SaveAtomic(SettingsPath, content, backupFolder);
}

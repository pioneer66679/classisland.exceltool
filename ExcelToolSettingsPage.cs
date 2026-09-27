using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json.Nodes;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using ClassIsland.Core.Abstractions.Controls;
using ClassIsland.Core.Abstractions.Services;
using ClassIsland.Core.Attributes;
using ClassIsland.Core.Enums.SettingsWindow;
using ClassIsland.ExcelTool.Core;
using ClosedXML.Excel;
using Microsoft.Extensions.Logging;

namespace ClassIsland.ExcelTool;

/// <summary>
/// 独立配置界面（设置 → 档案 Excel 导入导出）。
/// 界面用代码构建 Avalonia 控件树，不依赖 XAML 资源字典，避免样式缺失导致页面打不开。
/// </summary>
[SettingsPageInfo("classisland.exceltool.settings", "档案 Excel 导入导出", SettingsPageCategory.External)]
public class ExcelToolSettingsPage : SettingsPageBase
{
    private readonly ILogger<ExcelToolSettingsPage>? _logger;
    private readonly ProfileStore _store;

    private readonly TextBox _exportPathBox = new() { Watermark = "导出到哪个 .xlsx 文件" };
    private readonly TextBox _importPathBox = new() { Watermark = "要导入的 .xlsx 文件路径" };
    private readonly TextBox _backupBox = new() { Watermark = "留空则用插件配置目录\\Backups" };
    private readonly CheckBox _autoBackup = new() { Content = "导入前自动备份档案（建议开启）" };
    private readonly CheckBox _styleHeader = new() { Content = "导出时加表头样式并冻结首行" };
    private readonly CheckBox _skipEmpty = new() { Content = "导入时跳过空行" };
    private readonly ComboBox _strategyBox = new() { Width = 220 };
    private readonly Dictionary<string, CheckBox> _sheetBoxes = new();
    private readonly TextBox _logBox = new()
    {
        IsReadOnly = true,
        AcceptsReturn = true,
        TextWrapping = TextWrapping.NoWrap,
        Height = 200,
        FontFamily = new FontFamily("Consolas, monospace")
    };

    private ExcelToolConfig _config = new();
    private string _configPath = "";
    private ProfileStore _profileStore = null!;

    /// <summary>
    /// 插件配置目录。SettingsPageBase 不继承 PluginBase，拿不到 PluginConfigFolder，
    /// 所以由 Plugin.Initialize 在注册时通过 DI 传进来（见 Plugin.cs 的 AddSettingsPage 重载配合）。
    /// 兜底：没传时退回插件程序集所在目录。
    /// </summary>
    private static string ConfigFolderOverride { get; set; } = "";

    public ExcelToolSettingsPage(ILogger<ExcelToolSettingsPage>? logger = null)
    {
        _logger = logger;
        _store = new ProfileStore(ProfileStore.LocateDataRoot() ?? AppContext.BaseDirectory);
        _profileStore = _store;

        _strategyBox.Items.Add("合并（推荐）：Excel 有就更新，本地多余的保留");
        _strategyBox.Items.Add("覆盖：以 Excel 为准，本地多余的会被删除");
        _strategyBox.Items.Add("跳过已存在：只新增，已存在的一律不动");
        _strategyBox.SelectedIndex = 0;

        Content = BuildUi();
        LoadConfig();
        RefreshStatus();
    }

    /// <summary>由插件入口在注册配置页之前调用，告诉界面插件配置目录在哪。</summary>
    public static void SetConfigFolder(string folder) => ConfigFolderOverride = folder ?? "";

    // ---------------- UI 构建 ----------------

    private Control BuildUi()
    {
        var panel = new StackPanel
        {
            Spacing = 10,
            Margin = new Avalonia.Thickness(16),
            HorizontalAlignment = HorizontalAlignment.Stretch
        };

        panel.Children.Add(Header("档案转 Excel / Excel 转档案"));
        panel.Children.Add(Hint(
            "导出会把课表（时间表）、科目、课程表与全局配置写成同一个 Excel 工作簿的多个工作表；" +
            "导入支持「合并 / 覆盖 / 跳过」三种策略，并在导入前显示差异预览。"));

        panel.Children.Add(Header("当前位置"));
        panel.Children.Add(_statusText);

        panel.Children.Add(Header("导出"));
        panel.Children.Add(Hint("点「导出为 Excel」会直接弹出保存对话框，路径可留空。也可以先选好路径再导出。"));
        panel.Children.Add(_exportPathBox);
        var exportRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        exportRow.Children.Add(MakeButton("📤 导出为 Excel", OnExport));
        exportRow.Children.Add(MakeButton("仅选择路径…", OnPickExport));
        panel.Children.Add(exportRow);

        panel.Children.Add(Header("导出内容（勾选要写入的工作表）"));
        var sheetRow = new WrapPanel { Orientation = Orientation.Horizontal };
        foreach (var s in SheetNames.All)
        {
            var cb = new CheckBox { Content = s, IsChecked = true, Margin = new Avalonia.Thickness(0, 0, 16, 0) };
            _sheetBoxes[s] = cb;
            sheetRow.Children.Add(cb);
        }
        panel.Children.Add(sheetRow);
        panel.Children.Add(_styleHeader);

        panel.Children.Add(Header("导入"));
        panel.Children.Add(Hint("点「预览差异」或「应用导入」都会先弹出文件选择框，路径可留空。"));
        panel.Children.Add(_importPathBox);
        var importRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        importRow.Children.Add(MakeButton("📥 预览差异", OnPreview));
        importRow.Children.Add(MakeButton("✅ 应用导入", OnApplyImport));
        importRow.Children.Add(MakeButton("仅选择文件…", OnPickImport));
        panel.Children.Add(importRow);

        var stratRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        stratRow.Children.Add(new TextBlock
        {
            Text = "导入策略：",
            VerticalAlignment = VerticalAlignment.Center
        });
        stratRow.Children.Add(_strategyBox);
        panel.Children.Add(stratRow);
        panel.Children.Add(_autoBackup);
        panel.Children.Add(_skipEmpty);

        panel.Children.Add(Header("备份目录"));
        panel.Children.Add(_backupBox);

        panel.Children.Add(Header("运行日志"));
        var logScroll = new ScrollViewer
        {
            Content = _logBox,
            Height = 200,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto
        };
        panel.Children.Add(logScroll);

        var footer = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        footer.Children.Add(MakeButton("清空日志", (_, _) => _logBox.Text = ""));
        footer.Children.Add(MakeButton("保存设置", (_, _) => { SaveConfig(); Log("设置已保存。"); }));
        panel.Children.Add(footer);

        return new ScrollViewer { Content = panel };
    }

    private readonly TextBlock _statusText = new()
    {
        TextWrapping = TextWrapping.Wrap,
        Foreground = new SolidColorBrush(Color.FromRgb(0x44, 0x44, 0x44))
    };

    private static TextBlock Header(string s) => new()
    {
        Text = s,
        FontWeight = FontWeight.SemiBold,
        FontSize = 15,
        Margin = new Avalonia.Thickness(0, 8, 0, 0)
    };

    private static TextBlock Hint(string s) => new()
    {
        Text = s,
        TextWrapping = TextWrapping.Wrap,
        Opacity = 0.75,
        FontSize = 12
    };

    private static Button MakeButton(string text, EventHandler<Avalonia.Interactivity.RoutedEventArgs> handler)
    {
        var b = new Button { Content = text, Padding = new Avalonia.Thickness(14, 6) };
        b.Click += handler;
        return b;
    }

    // ---------------- 配置读写 ----------------

    private string PluginConfigFolder => string.IsNullOrWhiteSpace(ConfigFolderOverride)
        ? AppContext.BaseDirectory
        : ConfigFolderOverride;

    private void LoadConfig()
    {
        _configPath = Path.Combine(PluginConfigFolder, "Settings.json");

        try
        {
            if (File.Exists(_configPath))
            {
                var node = JsonNode.Parse(File.ReadAllText(_configPath));
                _config = JsonSerializerCompat.Deserialize<ExcelToolConfig>(node) ?? new ExcelToolConfig();
            }
        }
        catch (Exception ex)
        {
            Log("读取配置失败，使用默认值：" + ex.Message);
            _config = new ExcelToolConfig();
        }

        _exportPathBox.Text = _config.LastExportPath;
        _importPathBox.Text = _config.LastImportPath;
        _backupBox.Text = _config.BackupFolder;
        _autoBackup.IsChecked = _config.AutoBackupBeforeImport;
        _styleHeader.IsChecked = _config.StyleHeader;
        _skipEmpty.IsChecked = _config.SkipEmptyRows;
        _strategyBox.SelectedIndex = _config.DefaultStrategy switch
        {
            ImportStrategy.Overwrite => 1,
            ImportStrategy.SkipExisting => 2,
            _ => 0
        };

        foreach (var (name, cb) in _sheetBoxes)
            cb.IsChecked = _config.EnabledSheets.Contains(name);
    }

    private void CollectConfig()
    {
        _config.LastExportPath = _exportPathBox.Text ?? "";
        _config.LastImportPath = _importPathBox.Text ?? "";
        _config.BackupFolder = _backupBox.Text ?? "";
        _config.AutoBackupBeforeImport = _autoBackup.IsChecked == true;
        _config.StyleHeader = _styleHeader.IsChecked == true;
        _config.SkipEmptyRows = _skipEmpty.IsChecked == true;
        _config.DefaultStrategy = _strategyBox.SelectedIndex switch
        {
            1 => ImportStrategy.Overwrite,
            2 => ImportStrategy.SkipExisting,
            _ => ImportStrategy.Merge
        };

        _config.EnabledSheets = _sheetBoxes.Where(kv => kv.Value.IsChecked == true)
            .Select(kv => kv.Key).ToList();
        if (_config.EnabledSheets.Count == 0)
            _config.EnabledSheets = new List<string>(SheetNames.All);
    }

    private void SaveConfig()
    {
        try
        {
            CollectConfig();
            var dir = Path.GetDirectoryName(_configPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
            File.WriteAllText(_configPath,
                JsonSerializerCompat.Serialize(_config));
        }
        catch (Exception ex)
        {
            Log("保存配置失败：" + ex.Message);
        }
    }

    private string EffectiveBackupFolder()
        => string.IsNullOrWhiteSpace(_config.BackupFolder)
            ? Path.Combine(Path.GetDirectoryName(_configPath) ?? AppContext.BaseDirectory, "Backups")
            : _config.BackupFolder;

    // ---------------- 行为 ----------------

    private void RefreshStatus()
    {
        var root = ProfileStore.LocateDataRoot();
        var sb = new StringBuilder();

        if (root is null)
        {
            _statusText.Text = "⚠ 没找到 ClassIsland 的 data 目录，导出导入都会失败。请确认插件装在 <CL>\\data\\Plugins 下。";
            return;
        }

        sb.AppendLine($"data 目录：{root}");
        sb.AppendLine($"当前档案：{_profileStore.GetCurrentProfilePath()}");
        var profiles = _profileStore.ListProfiles();
        sb.AppendLine($"可用档案（{profiles.Count}）：{string.Join("、", profiles)}");
        sb.AppendLine($"插件配置目录：{PluginConfigFolder}");
        _statusText.Text = sb.ToString();
    }

    private async void OnPickExport(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        try
        {
            var top = TopLevelResolver.Resolve(this);
            if (top is null)
            {
                // 不再静默失败：把诊断信息写进日志，并提示可手写路径
                Log("⚠ 无法打开保存对话框（拿不到窗口）。");
                foreach (var line in TopLevelResolver.Diagnose(this).Split('\n'))
                    Log("   " + line.TrimEnd());
                Log("   → 可以直接在上面的输入框里手写完整路径。");
                return;
            }

            var file = await top.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = "导出档案为 Excel",
                SuggestedFileName = $"ClassIsland档案_{DateTime.Now:yyyyMMdd_HHmm}.xlsx",
                DefaultExtension = "xlsx",
                ShowOverwritePrompt = true,
                FileTypeChoices = new[]
                {
                    new FilePickerFileType("Excel 工作簿") { Patterns = new[] { "*.xlsx" } }
                }
            });

            if (file is null) { Log("已取消选择。"); return; }

            var p = ToLocalPath(file);
            if (!string.IsNullOrWhiteSpace(p))
            {
                _exportPathBox.Text = p;
                Log("已选择导出路径：" + p);
            }
            else
            {
                Log("⚠ 选中的文件拿不到本地路径（可能是虚拟位置），请手动填写完整路径。");
            }
        }
        catch (Exception ex)
        {
            Log("✗ 选择导出路径失败：" + ex.Message);
        }
    }

    /// <summary>
    /// 从 IStorageItem 取本地路径，多路兜底：
    /// LocalPath → Path（Uri）→ TryGetLocalPath 风格转换。
    /// </summary>
    private static string ToLocalPath(Avalonia.Platform.Storage.IStorageItem? item)
    {
        if (item is null) return "";

        try
        {
            var lp = item.Path?.LocalPath;
            if (!string.IsNullOrWhiteSpace(lp)) return lp;
        }
        catch { }

        try
        {
            var uri = item.Path;
            if (uri is not null)
            {
                if (uri.IsFile) return uri.LocalPath;
                var s = uri.ToString();
                if (!string.IsNullOrWhiteSpace(s)) return s;
            }
        }
        catch { }

        return "";
    }

    private async void OnPickImport(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        try
        {
            var top = TopLevelResolver.Resolve(this);
            if (top is null)
            {
                Log("⚠ 无法打开文件对话框（拿不到窗口）。");
                foreach (var line in TopLevelResolver.Diagnose(this).Split('\n'))
                    Log("   " + line.TrimEnd());
                Log("   → 可以直接在上面的输入框里手写完整路径。");
                return;
            }

            var files = await top.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = "选择要导入的 Excel",
                AllowMultiple = false,
                FileTypeFilter = new[]
                {
                    new FilePickerFileType("Excel 工作簿")
                    {
                        Patterns = new[] { "*.xlsx", "*.xlsm" },
                        MimeTypes = new[]
                        {
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        }
                    }
                }
            });

            if (files is null || files.Count == 0) { Log("已取消选择。"); return; }

            var p = ToLocalPath(files[0]);
            if (!string.IsNullOrWhiteSpace(p))
            {
                _importPathBox.Text = p;
                Log("已选择导入文件：" + p);

                // 立刻检查文件是否真的存在，早暴露问题
                if (!File.Exists(p))
                    Log("⚠ 该路径当前不存在，导入前请确认文件还在。");
            }
            else
            {
                Log("⚠ 选中的文件拿不到本地路径，请手动填写完整路径。");
            }
        }
        catch (Exception ex)
        {
            Log("✗ 选择导入文件失败：" + ex.Message);
        }
    }

    private async void OnExport(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        try
        {
            CollectConfig();

            var path = _exportPathBox.Text?.Trim() ?? "";

            // 一键直达：没填路径就先弹保存对话框，选完直接导出，不需要主人先点「选择」。
            if (string.IsNullOrWhiteSpace(path))
            {
                path = await PickSavePath();
                if (string.IsNullOrWhiteSpace(path)) return;   // 主人取消了
                _exportPathBox.Text = path;
            }

            var profile = _profileStore.LoadProfile();
            var settings = _profileStore.LoadSettings();

            using var wb = ExcelEngine.BuildWorkbook(profile, settings, _config.EnabledSheets, _config.StyleHeader);

            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
            wb.SaveAs(path);

            var fi = new FileInfo(path);
            Log($"✓ 导出成功：{path}");
            Log($"  工作表：{string.Join("、", _config.EnabledSheets)}");
            Log($"  文件大小：{fi.Length / 1024.0:F1} KB");

            SaveConfig();
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "导出 Excel 失败");
            Log("✗ 导出失败：" + ex.Message);
        }
    }

    /// <summary>弹出保存对话框，返回选中的路径；取消或失败返回空串。</summary>
    private async Task<string> PickSavePath()
    {
        var top = TopLevelResolver.Resolve(this);
        if (top is null)
        {
            Log("⚠ 打不开保存对话框，可以手动在输入框里填完整路径（如 D:\\课表.xlsx）。");
            return "";
        }

        var file = await top.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "导出档案为 Excel",
            SuggestedFileName = $"ClassIsland档案_{DateTime.Now:yyyyMMdd_HHmm}.xlsx",
            DefaultExtension = "xlsx",
            ShowOverwritePrompt = true,
            FileTypeChoices = new[]
            {
                new FilePickerFileType("Excel 工作簿") { Patterns = new[] { "*.xlsx" } }
            }
        });

        if (file is null) { Log("已取消导出。"); return ""; }

        var p = ToLocalPath(file);
        if (string.IsNullOrWhiteSpace(p))
            Log("⚠ 选中的位置拿不到本地路径，请手动填写完整路径。");

        return p;
    }

    /// <summary>
    /// 确保有一个可用的导入文件路径：没填就弹打开对话框，选完立刻校验存在性。
    /// 返回空串表示主人取消或取不到路径，调用方直接结束即可。
    /// </summary>
    private async Task<string> EnsureImportPath()
    {
        var path = _importPathBox.Text?.Trim() ?? "";

        if (string.IsNullOrWhiteSpace(path))
        {
            path = await PickImportPath();
            if (string.IsNullOrWhiteSpace(path)) return "";
            _importPathBox.Text = path;
        }

        if (!File.Exists(path))
        {
            Log($"✗ 找不到这个文件：{path}");
            Log("   可以点输入框右边的「选择 Excel 文件…」重新选一个。");
            return "";
        }

        return path;
    }

    /// <summary>弹出打开对话框，返回选中的文件路径；取消或失败返回空串。</summary>
    private async Task<string> PickImportPath()
    {
        var top = TopLevelResolver.Resolve(this);
        if (top is null)
        {
            Log("⚠ 打不开文件对话框，可以手动在输入框里填完整路径（如 D:\\课表.xlsx）。");
            return "";
        }

        var files = await top.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "选择要导入的 Excel",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("Excel 工作簿")
                {
                    Patterns = new[] { "*.xlsx", "*.xlsm" },
                    MimeTypes = new[]
                    {
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    }
                }
            }
        });

        if (files is null || files.Count == 0) { Log("已取消选择。"); return ""; }

        var p = ToLocalPath(files[0]);
        if (string.IsNullOrWhiteSpace(p))
        {
            Log("⚠ 选中的文件拿不到本地路径，请手动填写完整路径。");
            return "";
        }

        Log("已选择导入文件：" + p);
        return p;
    }

    private async void OnPreview(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        try
        {
            CollectConfig();

            var path = await EnsureImportPath();
            if (string.IsNullOrWhiteSpace(path)) return;

            var curProfile = _profileStore.LoadProfile();
            var curSettings = _profileStore.LoadSettings();

            var result = ImportEngine.Parse(path, curProfile, curSettings, _config.EnabledSheets);

            Log("── 差异预览 ──");
            Log("  " + result.Diff.Summary);

            foreach (var w in result.Warnings) Log("  ⚠ " + w);

            foreach (var kind in new[] { DiffKind.Added, DiffKind.Modified, DiffKind.Removed })
            {
                var items = result.Diff.Entries.Where(x => x.Kind == kind).Take(30).ToList();
                if (items.Count == 0) continue;

                var label = kind switch
                {
                    DiffKind.Added => "新增",
                    DiffKind.Modified => "修改",
                    _ => "会删除（仅覆盖模式）"
                };
                Log($"  [{label}]");
                foreach (var it in items) Log($"    · {it.Section} / {it.Target} — {it.Detail}");
            }

            _config.LastReport = result.Diff.Summary;
            SaveConfig();
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "预览差异失败");
            Log("✗ 预览失败：" + ex.Message);
        }
    }

    private async void OnApplyImport(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        try
        {
            CollectConfig();

            var path = await EnsureImportPath();
            if (string.IsNullOrWhiteSpace(path)) return;

            var profilePath = _profileStore.GetCurrentProfilePath();
            var curProfile = _profileStore.LoadProfile();
            var curSettings = _profileStore.LoadSettings();

            var parsed = ImportEngine.Parse(path, curProfile, curSettings, _config.EnabledSheets);

            var (finalProfile, finalSettings) =
                ImportEngine.ApplyStrategy(curProfile, curSettings, parsed.Profile, parsed.Settings, _config.DefaultStrategy);

            var backupFolder = _config.AutoBackupBeforeImport ? EffectiveBackupFolder() : null;
            if (backupFolder is not null)
                Log($"  备份目录：{backupFolder}");

            _profileStore.SaveProfile(finalProfile, profilePath, backupFolder);
            _profileStore.SaveSettings(finalSettings, backupFolder);

            Log($"✓ 导入完成（策略：{_config.DefaultStrategy}）");
            Log($"  {parsed.Diff.Summary}");
            Log("  提示：回到 ClassIsland 主界面后重新打开设置或重启应用，即可看到新档案。");

            RefreshStatus();
            SaveConfig();
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "应用导入失败");
            Log("✗ 导入失败：" + ex.Message);
        }
    }

    private void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] {msg}";
        _logBox.Text = string.IsNullOrEmpty(_logBox.Text) ? line : _logBox.Text + Environment.NewLine + line;
    }
}

using System;
using System.Collections.Generic;

namespace ClassIsland.ExcelTool.Core;

/// <summary>
/// Excel 工作簿的 Sheet 约定名。导入导出共用，保证双向一致。
/// </summary>
public static class SheetNames
{
    /// <summary>时间表（TimeLayout）：每一行 = 一个时间段。</summary>
    public const string TimeLayouts = "时间表";

    /// <summary>科目（Subject）：每一行 = 一个科目。</summary>
    public const string Subjects = "科目";

    /// <summary>课程表（ClassPlan）：每一行 = 某课表某时间点的一节课。</summary>
    public const string ClassPlans = "课程表";

    /// <summary>课程表清单：每一行 = 一份课表的元数据。</summary>
    public const string ClassPlanMeta = "课表清单";

    /// <summary>全局配置（Settings.json）：每一行 = 一个配置项 key/value。</summary>
    public const string Settings = "全局配置";

    /// <summary>导入导出都要包含的 Sheet（按顺序）。</summary>
    public static readonly string[] All = { TimeLayouts, Subjects, ClassPlans, ClassPlanMeta, Settings };
}

/// <summary>
/// 导入时发现的差异级别。
/// </summary>
public enum DiffKind
{
    /// <summary>Excel 里有、CL 里没有 —— 会新增。</summary>
    Added,

    /// <summary>两边都有但内容不同 —— 会更新。</summary>
    Modified,

    /// <summary>两边都有且内容一致 —— 无需改动。</summary>
    Unchanged,

    /// <summary>CL 里有、Excel 里没有 —— 覆盖模式下会删除。</summary>
    Removed
}

/// <summary>
/// 一条差异记录，用于导入前的预览。
/// </summary>
public sealed class DiffEntry
{
    public DiffKind Kind { get; init; }

    /// <summary>所属 Sheet（如「时间表」）。</summary>
    public string Section { get; init; } = "";

    /// <summary>对象标识（如时间表名 + 行号，或配置项 key）。</summary>
    public string Target { get; init; } = "";

    /// <summary>人类可读的变更说明。</summary>
    public string Detail { get; init; } = "";

    public override string ToString() => $"[{Kind}] {Section} / {Target} — {Detail}";
}

/// <summary>
/// 差异报告：给配置界面预览用。
/// </summary>
public sealed class DiffReport
{
    public List<DiffEntry> Entries { get; } = new();

    public int Added => Count(DiffKind.Added);
    public int Modified => Count(DiffKind.Modified);
    public int Unchanged => Count(DiffKind.Unchanged);
    public int Removed => Count(DiffKind.Removed);

    private int Count(DiffKind k)
    {
        var n = 0;
        foreach (var e in Entries)
            if (e.Kind == k) n++;
        return n;
    }

    public bool HasChanges => Added > 0 || Modified > 0 || Removed > 0;

    public string Summary =>
        $"新增 {Added} 项，修改 {Modified} 项，删除 {Removed} 项，未变 {Unchanged} 项";
}

/// <summary>
/// 导入策略（主人在配置界面里选）。
/// </summary>
public enum ImportStrategy
{
    /// <summary>覆盖：以 Excel 为准，CL 中多余的同名内容被移除；未出现在 Excel 中的其它档案保留。</summary>
    Overwrite,

    /// <summary>合并：Excel 中的内容更新/新增，CL 中已有但 Excel 未提及的保留。</summary>
    Merge,

    /// <summary>跳过：只新增，已存在的对象一律不动。</summary>
    SkipExisting
}

/// <summary>
/// 插件配置（持久化到 PluginConfigFolder\Settings.json）。
/// </summary>
public sealed class ExcelToolConfig
{
    /// <summary>上次导出目录。</summary>
    public string LastExportPath { get; set; } = "";

    /// <summary>上次导入文件。</summary>
    public string LastImportPath { get; set; } = "";

    /// <summary>导出时包含哪些 Sheet。</summary>
    public List<string> EnabledSheets { get; set; } = new(SheetNames.All);

    /// <summary>导入前自动备份档案（强烈建议开启）。</summary>
    public bool AutoBackupBeforeImport { get; set; } = true;

    /// <summary>备份目录，留空则用 PluginConfigFolder\Backups。</summary>
    public string BackupFolder { get; set; } = "";

    /// <summary>默认导入策略。</summary>
    public ImportStrategy DefaultStrategy { get; set; } = ImportStrategy.Merge;

    /// <summary>导出时是否写入表头样式（加粗+冻结首行）。</summary>
    public bool StyleHeader { get; set; } = true;

    /// <summary>
    /// 是否在导出结果里写入 ID 列（时间表ID / 科目ID / 课表ID 等）。
    ///
    /// 默认 false：这些 ID 是 ClassIsland 档案内部的 GUID 主键（36 位乱码），
    /// 对人眼核对毫无帮助，且每行都要重复一遍，会把表格糊满。
    /// 导入时按「名称」匹配即可，不依赖这些 ID（见 ImportEngine 的分组逻辑）。
    ///
    /// 打开后 ID 会写成「隐藏列」，排在所有可见列的右边 —— 需要精确定位时展开看一眼，
    /// 平时不影响阅读，也不会被误改。
    /// </summary>
    public bool ShowIdColumns { get; set; } = false;

    /// <summary>导入时是否跳过空行。</summary>
    public bool SkipEmptyRows { get; set; } = true;

    /// <summary>最近一次导入的差异报告（文本，供查看）。</summary>
    public string LastReport { get; set; } = "";
}

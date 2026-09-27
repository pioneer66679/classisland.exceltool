using System;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.VisualTree;

namespace ClassIsland.ExcelTool.Core;

/// <summary>
/// 健壮地拿到 TopLevel（用于打开文件/文件夹选择对话框）。
///
/// 背景：CL 设置页里的控件是通过 ContentPresenter 动态装载的，
/// 在「页面刚构造完、还没挂到视觉树」或「窗口结构特殊」时，
/// TopLevel.GetTopLevel(control) 会返回 null，导致文件选择按钮点了没反应。
///
/// 这里按可靠性从高到低依次尝试多条路径，只要有一条成功就能弹出对话框：
///   1. TopLevel.GetTopLevel(control)      —— 标准路径
///   2. 沿视觉树向上找 TopLevel / Window    —— 处理控件未直接挂载到顶层的情况
///   3. control.VisualRoot as TopLevel     —— 兜底
///   4. 已打开的窗口列表里找活动窗口        —— 最后保险
/// </summary>
public static class TopLevelResolver
{
    /// <summary>
    /// 尝试解析出可用的 TopLevel。返回 null 表示确实拿不到（调用方应提示用户）。
    /// </summary>
    public static TopLevel? Resolve(Visual? anchor)
    {
        // 1) 标准路径
        if (anchor is not null)
        {
            try
            {
                var tl = TopLevel.GetTopLevel(anchor);
                if (HasStorage(tl)) return tl;
            }
            catch { /* 继续下一条路径 */ }

            // 2) 沿视觉树向上逐级找
            try
            {
                foreach (var v in anchor.GetSelfAndVisualAncestors())
                {
                    if (v is TopLevel t && HasStorage(t)) return t;
                }
            }
            catch { }

            // 3) VisualRoot 兜底
            try
            {
                if (anchor.GetVisualRoot() is TopLevel vt && HasStorage(vt)) return vt;
            }
            catch { }
        }

        // 4) 全窗口列表里找活动窗口（anchor 为 null 或前几条都失败时）
        try
        {
            var windows = (Application.Current?.ApplicationLifetime as
                Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime)?.Windows;
            if (windows is not null)
            {
                var active = windows.FirstOrDefault(w => w.IsActive && HasStorage(w));
                if (active is not null) return active;

                var any = windows.FirstOrDefault(w => HasStorage(w));
                if (any is not null) return any;
            }
        }
        catch { }

        return null;
    }

    /// <summary>该 TopLevel 是否真的具备文件选择能力。</summary>
    public static bool HasStorage(TopLevel? tl)
    {
        try
        {
            var sp = tl?.StorageProvider;
            return sp is not null;
        }
        catch { return false; }
    }

    /// <summary>诊断信息：解析失败时告诉用户到底卡在哪，而不是静默无反应。</summary>
    public static string Diagnose(Visual? anchor)
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine("无法打开文件选择对话框。诊断信息：");
        sb.AppendLine($"  锚点控件：{(anchor is null ? "null" : anchor.GetType().Name)}");

        if (anchor is not null)
        {
            var byStd = TopLevel.GetTopLevel(anchor);
            sb.AppendLine($"  TopLevel.GetTopLevel → {(byStd is null ? "null" : byStd.GetType().Name)}");
            sb.AppendLine($"  视觉树祖先数：{anchor.GetSelfAndVisualAncestors().Count()}");
        }

        var life = Application.Current?.ApplicationLifetime;
        sb.AppendLine($"  应用生命周期：{(life is null ? "null" : life.GetType().Name)}");

        if (life is Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime d)
            sb.AppendLine($"  已打开窗口数：{d.Windows.Count}");
        else
            sb.AppendLine("  （非桌面生命周期，无法枚举窗口）");

        sb.AppendLine("  建议：直接在输入框里手写完整路径（如 D:\\备份\\课表.xlsx），同样可用。");
        return sb.ToString();
    }
}

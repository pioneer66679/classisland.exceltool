using ClassIsland.Core.Abstractions;
using ClassIsland.Core.Extensions.Registry;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace ClassIsland.ExcelTool;

/// <summary>
/// 插件入口。ClassIsland 通过 manifest.yml 的 entranceAssembly 找到本程序集，
/// 再按 [PluginEntrance] 定位到本类并调用 Initialize 完成依赖注入。
/// </summary>
public class Plugin : PluginBase
{
    public override void Initialize(HostBuilderContext context, IServiceCollection services)
    {
        // 把插件配置目录传给配置界面（SettingsPageBase 拿不到 PluginConfigFolder）
        ExcelToolSettingsPage.SetConfigFolder(PluginConfigFolder);

        // 注册独立配置界面：设置窗口 → 外部 分类下会出现「档案 Excel 导入导出」
        services.AddSettingsPage<ExcelToolSettingsPage>();
    }
}

using System;
using System.ComponentModel.Composition;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.TextManager.Interop;
using Task = System.Threading.Tasks.Task;

namespace EncodingDisplayExtension
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(PackageGuidString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class EncodingDisplayPackage : AsyncPackage
    {
        public const string PackageGuidString = "a8b9c0d1-e2f3-4a5b-6c7d-8e9f0a1b2c3d"; // 注意：如果你创建了新项目，请保留原文件中的 GUID

        private IComponentModel _componentModel;
        private IVsTextManager _textManager;
        private IVsEditorAdaptersFactoryService _editorAdapter;

        // 这是我们要插入到状态栏的 WPF 控件
        private TextBlock _encodingTextBlock;

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            // 1. 获取服务
            _textManager = await GetServiceAsync(typeof(SVsTextManager)) as IVsTextManager;
            _componentModel = await GetServiceAsync(typeof(SComponentModel)) as IComponentModel;
            if (_componentModel != null)
            {
                _editorAdapter = _componentModel.GetService<IVsEditorAdaptersFactoryService>();
            }

            // 2. 初始化界面 (注入 WPF 控件)
            InjectStatusBarUI();

            // 3. 注册事件监听
            if (_textManager != null)
            {
                var dte = await GetServiceAsync(typeof(EnvDTE.DTE)) as EnvDTE.DTE;
                if (dte != null)
                {
                    dte.Events.WindowEvents.WindowActivated += OnWindowActivated;
                }
            }

            // 4. 初次运行更新
            UpdateEncodingDisplay();
        }

        private void InjectStatusBarUI()
        {
            // 确保在 UI 线程运行
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                // 获取 VS 主窗口
                var mainWindow = Application.Current.MainWindow;
                if (mainWindow == null) return;

                // 查找状态栏 (StatusBar)
                var statusBar = FindChild<StatusBar>(mainWindow);
                if (statusBar != null)
                {
                    // 创建我们要显示的文本块
                    _encodingTextBlock = new TextBlock
                    {
                        Text = "Loading...",
                        Margin = new Thickness(10, 0, 10, 0), // 左右留点空隙
                        VerticalAlignment = VerticalAlignment.Center,
                        Foreground = System.Windows.SystemColors.WindowTextBrush, // 简单设置白色，适配深色主题；更完美的做法是绑定 VS Theme Key
                        ToolTip = "Current File Encoding"
                    };

                    // 包装在 StatusBarItem 中
                    var newItem = new StatusBarItem
                    {
                        Content = _encodingTextBlock,
                        Name = "EncodingDisplayItem" // 给个名字防止重复添加
                    };

                    // 关键布局设置：停靠在右侧
                    DockPanel.SetDock(newItem, Dock.Right);

                    // 检查是否已经添加过，防止重复
                    bool exists = false;
                    foreach (var child in statusBar.Items)
                    {
                        if (child is StatusBarItem item && item.Name == "EncodingDisplayItem")
                        {
                            exists = true;
                            // 如果找到了旧的，更新引用即可
                            if (item.Content is TextBlock tb) _encodingTextBlock = tb;
                            break;
                        }
                    }

                    if (!exists)
                    {
                        // 插入到 Items 集合中。
                        // 在 DockPanel 中，对于 Dock.Right 的元素：
                        // 列表前面的元素会被推到最右边，列表后面的元素会往左排。
                        // 为了让它显示在 WakaTime (通常是最右侧) 的左边，我们直接 Add 到末尾即可。
                        statusBar.Items.Add(newItem);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error injecting UI: {ex.Message}");
            }
        }

        // 辅助方法：递归查找子控件
        private T FindChild<T>(DependencyObject parent) where T : DependencyObject
        {
            if (parent == null) return null;

            int childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T typedChild)
                {
                    return typedChild;
                }

                var result = FindChild<T>(child);
                if (result != null)
                {
                    return result;
                }
            }
            return null;
        }

        private void OnWindowActivated(EnvDTE.Window gotFocus, EnvDTE.Window lostFocus)
        {
            _ = this.JoinableTaskFactory.RunAsync(async () =>
            {
                await this.JoinableTaskFactory.SwitchToMainThreadAsync();
                UpdateEncodingDisplay();
            });
        }

        private void UpdateEncodingDisplay()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // 如果 UI 还没初始化成功，就不做任何事
            if (_encodingTextBlock == null || _textManager == null || _editorAdapter == null)
                return;

            try
            {
                IVsTextView activeView = null;
                _textManager.GetActiveView(1, null, out activeView);

                if (activeView == null)
                {
                    _encodingTextBlock.Text = "";
                    return;
                }

                var wpfTextView = _editorAdapter.GetWpfTextView(activeView);
                if (wpfTextView == null) return;

                ITextBuffer buffer = wpfTextView.TextBuffer;

                if (buffer.Properties.TryGetProperty(typeof(ITextDocument), out ITextDocument document))
                {
                    if (document != null && document.Encoding != null)
                    {
                        // 更新 WPF 控件的文字
                        string encodingName = document.Encoding.WebName.ToUpper();
                        int codePage = document.Encoding.CodePage;

                        // 简单的显示格式，比如 "UTF-8"
                        _encodingTextBlock.Text = $"{encodingName}";

                        // 简单的颜色映射逻辑
                        // 注意：这里使用的 Brushes 颜色是针对深色主题优化的
                        if (encodingName == "UTF-8")
                        {
                            _encodingTextBlock.Foreground = Brushes.LightGray; // 默认
                        }
                        else if (encodingName.Contains("GB") || codePage == 936) // GB2312, GBK, GB18030
                        {
                            _encodingTextBlock.Foreground = Brushes.SpringGreen; // 醒目的绿色
                        }
                        else if (encodingName.Contains("ASCII"))
                        {
                            _encodingTextBlock.Foreground = Brushes.SkyBlue; // 蓝色
                        }
                        else if (encodingName.Contains("UTF-16") || encodingName.Contains("UNICODE"))
                        {
                            _encodingTextBlock.Foreground = Brushes.Violet; // 紫色
                        }
                        else
                        {
                            _encodingTextBlock.Foreground = Brushes.Orange; // 其他生僻编码用橙色警告
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Update Error: {ex.Message}");
            }
        }
    }
}
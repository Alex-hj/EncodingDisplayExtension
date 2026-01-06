using System;
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
        private EnvDTE.DTE _dte;
        private EnvDTE.WindowEvents _windowEvents;

        // 当前跟踪的文档（用于监听编码变化）
        private ITextDocument _currentDocument;

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
                _dte = await GetServiceAsync(typeof(EnvDTE.DTE)) as EnvDTE.DTE;
                if (_dte != null)
                {
                    _windowEvents = _dte.Events.WindowEvents;
                    _windowEvents.WindowActivated += OnWindowActivated;
                }
            }

            // 4. 监听 VS 主窗口获得焦点事件（从其他应用切回时刷新）
            var mainWindow = Application.Current.MainWindow;
            if (mainWindow != null)
            {
                mainWindow.Activated += OnMainWindowActivated;
            }

            // 5. 初次运行更新
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
                await Task.Delay(50); // 等待 TextManager 状态同步
                UpdateEncodingDisplay();
            });
        }

        private void OnMainWindowActivated(object sender, EventArgs e)
        {
            // 从其他应用切回 VS 时刷新编码显示
            _ = this.JoinableTaskFactory.RunAsync(async () =>
            {
                await this.JoinableTaskFactory.SwitchToMainThreadAsync();
                await Task.Delay(50); // 等待状态同步
                UpdateEncodingDisplay();
            });
        }

        private void OnEncodingChanged(object sender, EncodingChangedEventArgs e)
        {
            // 文档编码发生变化时刷新显示
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

                if (buffer.Properties.TryGetProperty(typeof(ITextDocument), out ITextDocument document) &&
                    document != null && document.Encoding != null)
                {
                    // 如果文档发生变化，更新编码变化事件的订阅
                    if (_currentDocument != document)
                    {
                        // 取消旧文档的订阅
                        if (_currentDocument != null)
                        {
                            _currentDocument.EncodingChanged -= OnEncodingChanged;
                        }
                        // 订阅新文档的编码变化事件
                        _currentDocument = document;
                        _currentDocument.EncodingChanged += OnEncodingChanged;
                    }

                    // 更新 WPF 控件的文字
                    string encodingName = document.Encoding.WebName.ToUpper();
                    int codePage = document.Encoding.CodePage;
                    bool hasBom = document.Encoding.GetPreamble().Length > 0;

                    // 区分 UTF-8 和 UTF-8 with BOM
                    if (encodingName == "UTF-8" && hasBom)
                    {
                        _encodingTextBlock.Text = "UTF-8 BOM";
                    }
                    else
                    {
                        _encodingTextBlock.Text = encodingName;
                    }

                    // 颜色映射逻辑（使用在深色和浅色主题下都较清晰的颜色）
                    if (encodingName == "UTF-8")
                    {
                        _encodingTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(0x60, 0xA0, 0x60)); // 柔和绿色
                    }
                    else if (encodingName.Contains("GB") || codePage == 936) // GB2312, GBK, GB18030
                    {
                        _encodingTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(0xE0, 0xA0, 0x30)); // 琥珀色警告
                    }
                    else if (encodingName.Contains("ASCII"))
                    {
                        _encodingTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(0x50, 0x90, 0xD0)); // 天蓝色
                    }
                    else if (encodingName.Contains("UTF-16") || encodingName.Contains("UNICODE"))
                    {
                        _encodingTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(0xA0, 0x70, 0xD0)); // 紫色
                    }
                    else
                    {
                        _encodingTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(0xE0, 0x60, 0x40)); // 橙红色警告
                    }
                }
                else
                {
                    // 非文本文件或无法获取编码信息时清空显示
                    _encodingTextBlock.Text = "";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Update Error: {ex.Message}");
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // 取消事件订阅，防止内存泄漏
                if (_windowEvents != null)
                {
                    _windowEvents.WindowActivated -= OnWindowActivated;
                    _windowEvents = null;
                }

                // 取消主窗口激活事件订阅
                var mainWindow = Application.Current?.MainWindow;
                if (mainWindow != null)
                {
                    mainWindow.Activated -= OnMainWindowActivated;
                }

                // 取消当前文档编码变化事件订阅
                if (_currentDocument != null)
                {
                    _currentDocument.EncodingChanged -= OnEncodingChanged;
                    _currentDocument = null;
                }

                _dte = null;
            }

            base.Dispose(disposing);
        }
    }
}
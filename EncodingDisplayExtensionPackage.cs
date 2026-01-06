using System;
using System.ComponentModel.Composition;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
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
    /// <summary>
    /// 这是插件的主入口类。
    /// PackageRegistration 属性告诉 VS 这个类存在。
    /// ProvideAutoLoad 属性确保插件在打开解决方案或没有解决方案时也会自动加载，
    /// 这样我们才能一直监听文件切换事件。
    /// </summary>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(PackageGuidString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class EncodingDisplayPackage : AsyncPackage
    {
        public const string PackageGuidString = "a8b9c0d1-e2f3-4a5b-6c7d-8e9f0a1b2c3d"; // 注意：如果你创建了新项目，请保留原文件中的 GUID

        // 状态栏服务
        private IVsStatusbar _statusBar;
        // MEF 组件宿主，用于获取现代编辑器服务
        private IComponentModel _componentModel;
        // 传统的文本管理器服务，用于监听视图变化
        private IVsTextManager _textManager;
        // 用于将 IVsTextView 转换为 WPF TextView
        private IVsEditorAdaptersFactoryService _editorAdapter;

        /// <summary>
        /// 初始化步骤。当 VS 加载插件时会调用此方法。
        /// </summary>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            // 切换到主线程，因为我们需要访问 UI 服务
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            // 获取基础服务
            _statusBar = await GetServiceAsync(typeof(SVsStatusbar)) as IVsStatusbar;
            _textManager = await GetServiceAsync(typeof(SVsTextManager)) as IVsTextManager;
            _componentModel = await GetServiceAsync(typeof(SComponentModel)) as IComponentModel;

            if (_componentModel != null)
            {
                _editorAdapter = _componentModel.GetService<IVsEditorAdaptersFactoryService>();
            }

            // 开始监听活动视图的变化
            if (_textManager != null)
            {
                // 创建一个监听器
                var sink = new ViewNotificationSink(this);
                // 注册监听器，当活动视图发生变化时通知我们
                // 注意：RegisterViewNotificationSink 是 IVsTextManager2 的方法
                // 这里我们简化处理，使用即时轮询或通过 DTE 事件也可以，但通过 WindowEvents 更稳健
                // 为了简化代码并确保不用复杂的 ConnectionPoint，我们这里使用 DTE 事件作为触发器

                // 重新获取 DTE 服务用于事件监听（这是最简单的方式）
                var dte = await GetServiceAsync(typeof(EnvDTE.DTE)) as EnvDTE.DTE;
                if (dte != null)
                {
                    // 订阅窗口切换事件
                    dte.Events.WindowEvents.WindowActivated += OnWindowActivated;
                }
            }

            // 初始化时尝试显示一次
            UpdateEncodingDisplay();
        }

        private void OnWindowActivated(EnvDTE.Window gotFocus, EnvDTE.Window lostFocus)
        {
            // 当窗口切换时，更新编码显示
            // 使用 JoinableTaskFactory 确保在主线程执行
            _ = this.JoinableTaskFactory.RunAsync(async () =>
            {
                await this.JoinableTaskFactory.SwitchToMainThreadAsync();
                UpdateEncodingDisplay();
            });
        }

        /// <summary>
        /// 核心方法：获取当前文件的编码并显示在状态栏
        /// </summary>
        private void UpdateEncodingDisplay()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_textManager == null || _editorAdapter == null || _statusBar == null)
                return;

            try
            {
                // 获取当前活动的视图
                IVsTextView activeView = null;
                _textManager.GetActiveView(1, null, out activeView);

                if (activeView == null)
                {
                    _statusBar.SetText(""); // 没有活动视图，清空状态栏
                    return;
                }

                // 将 COM 的 IVsTextView 转换为现代的 IWpfTextView
                var wpfTextView = _editorAdapter.GetWpfTextView(activeView);
                if (wpfTextView == null) return;

                // 从 TextView 获取 TextBuffer
                ITextBuffer buffer = wpfTextView.TextBuffer;

                // 尝试从 Buffer 中获取 ITextDocument（这就包含了文件路径和编码信息）
                // ITextDocument 来源于 Microsoft.VisualStudio.Text.Data
                if (buffer.Properties.TryGetProperty(typeof(ITextDocument), out ITextDocument document))
                {
                    if (document != null && document.Encoding != null)
                    {
                        // 构造显示字符串，例如： "UTF-8"
                        string statusText = document.Encoding.WebName.ToUpper();

                        // 写入状态栏
                        // 0 = 暂时显示，1 = 稳定显示
                        // 这里我们使用 SetText 直接写入主区域
                        _statusBar.SetText(statusText);
                    }
                }
            }
            catch (Exception ex)
            {
                // 发生错误时不崩溃，只是在输出窗口打印（可选）
                System.Diagnostics.Debug.WriteLine($"Error updating encoding: {ex.Message}");
            }
        }

        // 简单的辅助类，用于接收视图通知（如果不想用 DTE，可以用这个，本例主要用了 DTE）
        private class ViewNotificationSink : IVsTextManagerEvents
        {
            private readonly EncodingDisplayPackage _package;
            public ViewNotificationSink(EncodingDisplayPackage package) { _package = package; }
            public void OnRegisterMarkerType(int iMarkerType) { }
            public void OnRegisterView(IVsTextView pView) { }
            public void OnUnregisterView(IVsTextView pView) { }
            public void OnUserPreferencesChanged(VIEWPREFERENCES[] pViewPrefs, FRAMEPREFERENCES[] pFramePrefs, LANGPREFERENCES[] pLangPrefs, FONTCOLORPREFERENCES[] pColorPrefs) { }
        }
    }
}
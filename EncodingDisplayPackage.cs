using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
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
        public const string PackageGuidString = "a8b9c0d1-e2f3-4a5b-6c7d-8e9f0a1b2c3d";

        // 窗口/主窗口激活后，等待 TextManager 状态同步再刷新显示
        private const int ActivationRefreshDelayMs = 50;

        private EncodingStatusBarItem _statusBarItem;
        private ActiveDocumentTracker _documentTracker;
        private EnvDTE.WindowEvents _windowEvents;
        private Window _mainWindow;

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            var textManager = await GetServiceAsync(typeof(SVsTextManager)) as IVsTextManager;
            var componentModel = await GetServiceAsync(typeof(SComponentModel)) as IComponentModel;
            var dte = await GetServiceAsync(typeof(EnvDTE.DTE)) as EnvDTE.DTE;

            _documentTracker = CreateDocumentTracker(textManager, componentModel);
            _mainWindow = Application.Current?.MainWindow;
            _statusBarItem = EncodingStatusBarItem.TryInject(_mainWindow);
            SubscribeActivationEvents(dte);

            Refresh();
        }

        private ActiveDocumentTracker CreateDocumentTracker(IVsTextManager textManager, IComponentModel componentModel)
        {
            var editorAdapter = componentModel?.GetService<IVsEditorAdaptersFactoryService>();
            if (textManager == null || editorAdapter == null)
            {
                return null;
            }

            return new ActiveDocumentTracker(textManager, editorAdapter, () => ScheduleRefresh());
        }

        private void SubscribeActivationEvents(EnvDTE.DTE dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (dte != null)
            {
                // 需要持有 WindowEvents 引用，否则会被 GC 回收导致事件不再触发
                _windowEvents = dte.Events.WindowEvents;
                _windowEvents.WindowActivated += OnWindowActivated;
            }

            // 从其他应用切回 VS 时也要刷新
            if (_mainWindow != null)
            {
                _mainWindow.Activated += OnMainWindowActivated;
            }
        }

        private void UnsubscribeActivationEvents()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_windowEvents != null)
            {
                _windowEvents.WindowActivated -= OnWindowActivated;
                _windowEvents = null;
            }

            if (_mainWindow != null)
            {
                _mainWindow.Activated -= OnMainWindowActivated;
                _mainWindow = null;
            }
        }

        private void OnWindowActivated(EnvDTE.Window gotFocus, EnvDTE.Window lostFocus)
        {
            ScheduleRefresh(ActivationRefreshDelayMs);
        }

        private void OnMainWindowActivated(object sender, EventArgs e)
        {
            ScheduleRefresh(ActivationRefreshDelayMs);
        }

        // 统一的刷新入口，可在任意线程调用
        private void ScheduleRefresh(int delayMs = 0)
        {
            _ = JoinableTaskFactory.RunAsync(() => RefreshAsync(delayMs));
        }

        // 切到 UI 线程，按需延时后刷新
        private async Task RefreshAsync(int delayMs)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync();
            if (delayMs > 0)
            {
                await Task.Delay(delayMs);
            }

            Refresh();
        }

        private void Refresh()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // 如果 UI 或文档跟踪器还没初始化成功，就不做任何事
            if (_statusBarItem == null || _documentTracker == null)
            {
                return;
            }

            try
            {
                var document = _documentTracker.GetActiveDocument();
                if (document?.Encoding == null)
                {
                    // 没有活动文本视图、非文本文件或无法获取编码信息时清空显示
                    _statusBarItem.Clear();
                    return;
                }

                _statusBarItem.Show(EncodingDisplayInfo.From(document.Encoding));
            }
            catch (Exception ex)
            {
                ActivityLog.TryLogError(nameof(EncodingDisplayPackage), $"Update Error: {ex}");
            }
        }

        protected override void Dispose(bool disposing)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (disposing)
            {
                // 取消事件订阅并释放当前文档引用，防止内存泄漏
                UnsubscribeActivationEvents();
                _documentTracker?.Dispose();
                _documentTracker = null;
            }

            base.Dispose(disposing);
        }
    }
}

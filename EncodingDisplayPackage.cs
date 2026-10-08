using System;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Task = System.Threading.Tasks.Task;

namespace EncodingDisplayExtension
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [ProvideMenuResource("Menus.ctmenu", 2)] // EncodingMenu.vsct 有改动时递增，VS 据此重新合并菜单
    [Guid(PackageGuidString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class EncodingDisplayPackage : AsyncPackage
    {
        public const string PackageGuidString = "a8b9c0d1-e2f3-4a5b-6c7d-8e9f0a1b2c3d";

        // 窗口/主窗口激活后，等待 TextManager 状态同步再刷新显示
        private const int ActivationRefreshDelayMs = 50;

        private const string MessageTitle = "Change File Encoding";

        private EncodingStatusBarItem _statusBarItem;
        private ActiveDocumentTracker _documentTracker;
        private EncodingConverter _encodingConverter;
        private EncodingMenu _encodingMenu;
        private EnvDTE.WindowEvents _windowEvents;
        private Window _mainWindow;

        // 转换过程中会连续触发 EncodingChanged，此时忽略刷新请求，转换结束后按被转换的文档统一刷新
        private bool _isConverting;

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            var monitorSelection = await GetServiceAsync(typeof(SVsShellMonitorSelection)) as IVsMonitorSelection;
            var componentModel = await GetServiceAsync(typeof(SComponentModel)) as IComponentModel;
            var runningDocumentTable = await GetServiceAsync(typeof(SVsRunningDocumentTable)) as IVsRunningDocumentTable;
            var commandService = await GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            var dte = await GetServiceAsync(typeof(EnvDTE.DTE)) as EnvDTE.DTE;

            _documentTracker = CreateDocumentTracker(monitorSelection, componentModel);
            _encodingConverter = new EncodingConverter(runningDocumentTable);
            _encodingMenu = EncodingMenu.Create(commandService, OnEncodingSelected);
            _mainWindow = Application.Current?.MainWindow;
            _statusBarItem = EncodingStatusBarItem.TryInject(_mainWindow, OnStatusBarItemClicked);
            SubscribeActivationEvents(dte);

            Refresh();
        }

        private ActiveDocumentTracker CreateDocumentTracker(IVsMonitorSelection monitorSelection, IComponentModel componentModel)
        {
            var editorAdapter = componentModel?.GetService<IVsEditorAdaptersFactoryService>();
            if (monitorSelection == null || editorAdapter == null)
            {
                return null;
            }

            return new ActiveDocumentTracker(monitorSelection, editorAdapter, () => ScheduleRefresh());
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

            // 如果 UI 或文档跟踪器还没初始化成功，或正在转换编码，就不做任何事
            if (_statusBarItem == null || _documentTracker == null || _isConverting)
            {
                return;
            }

            try
            {
                var document = _documentTracker.GetActiveDocument();
                if (document?.Encoding == null)
                {
                    // 没有打开文本文档、活动文档不是文本文件或无法获取编码信息时清空显示
                    _statusBarItem.Clear();
                    return;
                }

                ShowEncoding(document.Encoding);
            }
            catch (Exception ex)
            {
                ActivityLog.TryLogError(nameof(EncodingDisplayPackage), $"Update Error: {ex}");
            }
        }

        // 显示编码；属于支持列表（SupportedEncodings）的才允许点击修改
        private void ShowEncoding(Encoding encoding)
        {
            _statusBarItem.Show(EncodingDisplayInfo.From(encoding), SupportedEncodings.Contains(encoding));
        }

        private void OnStatusBarItemClicked()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // 用状态栏正在显示的文档：它就是当前活动文档窗口里的文件
            ITextDocument document = _documentTracker?.TrackedDocument;
            if (document != null && _encodingMenu != null)
            {
                _encodingMenu.Show(_statusBarItem.ScreenPosition, document.Encoding);
            }
        }

        private void OnEncodingSelected(Encoding target)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            ITextDocument document = _documentTracker?.TrackedDocument;
            if (document != null)
            {
                ConvertEncoding(document, target);
            }
        }

        private void ConvertEncoding(ITextDocument document, Encoding target)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            _isConverting = true;
            try
            {
                string message = _encodingConverter.TryConvert(document, target);
                if (message != null)
                {
                    ShowMessage(message);
                }
            }
            catch (Exception ex)
            {
                ActivityLog.TryLogError(nameof(EncodingDisplayPackage), $"Convert Error: {ex}");
                ShowMessage($"Failed to change the file encoding: {ex.Message}");
            }
            finally
            {
                _isConverting = false;
            }

            ShowEncoding(document.Encoding);
        }

        private void ShowMessage(string message)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            VsShellUtilities.ShowMessageBox(
                this, message, MessageTitle,
                OLEMSGICON.OLEMSGICON_WARNING, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
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

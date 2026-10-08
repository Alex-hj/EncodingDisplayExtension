using System;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.TextManager.Interop;

namespace EncodingDisplayExtension
{
    /// <summary>
    /// 解析当前活动文档窗口对应的文档，并且只订阅该文档的 EncodingChanged 事件。
    /// 活动文档窗口只会在切换文档时变化：焦点移到输出、调用堆栈、解决方案资源管理器等工具窗口时，
    /// 仍然是原来的文件；只有没有文本文档（没有打开文件、图片、设计器等）时才返回 null。
    /// </summary>
    internal sealed class ActiveDocumentTracker : IDisposable
    {
        // VSFPROPID_IsDocDataInitialized：VS 17.9 起，异步打开的文档在初始化完成前读取 DocData 可能阻塞。
        // 数值已在本机 VS 17.14 的 Microsoft.VisualStudio.Interop 中核实；当前引用的 SDK 版本还没有对应的枚举。
        private const int IsDocDataInitializedPropertyId = -5055;

        private readonly ActiveDocumentFrameWatcher _frameWatcher;
        private readonly IVsEditorAdaptersFactoryService _editorAdapter;
        private readonly Action _onChanged;

        // 当前跟踪的文档（用于监听编码变化）
        private ITextDocument _trackedDocument;

        /// <param name="onChanged">
        /// 状态栏显示可能需要更新时调用：活动文档窗口变化（含延迟加载的文档显示出来、窗口关闭），或当前文档的编码变化。
        /// 可能在 VS 的事件回调里同步调用。
        /// </param>
        public ActiveDocumentTracker(
            IVsMonitorSelection monitorSelection,
            IVsEditorAdaptersFactoryService editorAdapter,
            Action onChanged)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            _editorAdapter = editorAdapter;
            _onChanged = onChanged;
            _frameWatcher = new ActiveDocumentFrameWatcher(monitorSelection, onChanged);
        }

        /// <summary>
        /// 最近一次 GetActiveDocument 解析出的文档，也就是状态栏当前显示的那个；没有则为 null。
        /// </summary>
        public ITextDocument TrackedDocument => _trackedDocument;

        /// <summary>
        /// 返回当前活动文档窗口的文档（没有则为 null），并让编码变化订阅跟随该文档。
        /// 没有文档时会释放旧订阅，避免继续持有已关闭的文档。
        /// </summary>
        public ITextDocument GetActiveDocument()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            ITextDocument document = FindActiveDocument();
            Track(document);
            return document;
        }

        public void Dispose()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            Track(null);
            _frameWatcher.Dispose();
        }

        private ITextDocument FindActiveDocument()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            ITextBuffer buffer = FindActiveDocumentBuffer();
            if (buffer == null)
            {
                return null;
            }

            buffer.Properties.TryGetProperty(typeof(ITextDocument), out ITextDocument document);
            return document;
        }

        // 活动文档窗口 → DocData（IVsTextBuffer）→ 文档缓冲区；DocData 不是文本缓冲区时返回 null
        private ITextBuffer FindActiveDocumentBuffer()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            IVsWindowFrame frame = _frameWatcher.ActiveFrame;
            if (frame == null || !IsDocDataInitialized(frame))
            {
                return null;
            }

            if (ErrorHandler.Failed(frame.GetProperty((int)__VSFPROPID.VSFPROPID_DocData, out object docData)))
            {
                return null;
            }

            return docData is IVsTextBuffer textBuffer ? _editorAdapter.GetDocumentBuffer(textBuffer) : null;
        }

        // 旧版本 VS 没有这个属性（读取失败），按已初始化处理；读到的值不是 bool 时按未初始化处理
        private static bool IsDocDataInitialized(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (ErrorHandler.Failed(frame.GetProperty(IsDocDataInitializedPropertyId, out object value)))
            {
                return true;
            }

            return value is bool initialized && initialized;
        }

        private void Track(ITextDocument document)
        {
            if (_trackedDocument == document)
            {
                return;
            }

            if (_trackedDocument != null)
            {
                _trackedDocument.EncodingChanged -= OnEncodingChanged;
            }

            _trackedDocument = document;

            if (_trackedDocument != null)
            {
                _trackedDocument.EncodingChanged += OnEncodingChanged;
            }
        }

        private void OnEncodingChanged(object sender, EncodingChangedEventArgs e)
        {
            _onChanged();
        }
    }
}

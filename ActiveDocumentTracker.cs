using System;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;

namespace EncodingDisplayExtension
{
    /// <summary>
    /// 解析当前活动编辑器对应的文档，并且只订阅该文档的 EncodingChanged 事件。
    /// </summary>
    internal sealed class ActiveDocumentTracker : IDisposable
    {
        // GetActiveView 的 fMustHaveFocus 参数：只返回当前拥有焦点的文本视图
        private const int MustHaveFocus = 1;

        private readonly IVsTextManager _textManager;
        private readonly IVsEditorAdaptersFactoryService _editorAdapter;
        private readonly Action _onEncodingChanged;

        // 当前跟踪的文档（用于监听编码变化）
        private ITextDocument _trackedDocument;

        public ActiveDocumentTracker(
            IVsTextManager textManager,
            IVsEditorAdaptersFactoryService editorAdapter,
            Action onEncodingChanged)
        {
            _textManager = textManager;
            _editorAdapter = editorAdapter;
            _onEncodingChanged = onEncodingChanged;
        }

        /// <summary>
        /// 返回当前活动文本视图的文档（没有则为 null），并让编码变化订阅跟随该文档。
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
            Track(null);
        }

        private ITextDocument FindActiveDocument()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            _textManager.GetActiveView(MustHaveFocus, null, out IVsTextView activeView);
            if (activeView == null)
            {
                return null;
            }

            IWpfTextView wpfTextView = _editorAdapter.GetWpfTextView(activeView);
            if (wpfTextView == null)
            {
                return null;
            }

            wpfTextView.TextBuffer.Properties.TryGetProperty(typeof(ITextDocument), out ITextDocument document);
            return document;
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
            _onEncodingChanged();
        }
    }
}

using System;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace EncodingDisplayExtension
{
    /// <summary>
    /// 跟踪“活动文档窗口”。做法与 Roslyn 的 VisualStudioActiveDocumentTracker 一致：
    /// 用 IVsSelectionEvents 得知新激活的文档窗口，再对它登记 IVsWindowFrameNotify，
    /// 以便在窗口真正显示出来（延迟加载的文档此时数据才初始化）以及关闭、隐藏、切走标签页时得到通知。
    /// 焦点移到工具窗口不会改变活动文档窗口。
    /// </summary>
    internal sealed class ActiveDocumentFrameWatcher : IVsSelectionEvents, IVsWindowFrameNotify, IDisposable
    {
        private readonly IVsMonitorSelection _monitorSelection;
        private readonly Action _onChanged;
        private uint _selectionCookie;

        // 当前登记了窗口通知的活动文档窗口
        private IVsWindowFrame2 _watchedFrame;
        private uint _frameCookie;

        public ActiveDocumentFrameWatcher(IVsMonitorSelection monitorSelection, Action onChanged)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            _monitorSelection = monitorSelection;
            _onChanged = onChanged;

            // 先取启动时已经存在的活动文档窗口，再订阅之后的切换
            SetActiveFrame(GetCurrentDocumentFrame());
            _monitorSelection.AdviseSelectionEvents(this, out _selectionCookie);
        }

        /// <summary>
        /// 当前活动文档窗口；没有打开文档时为 null。活动窗口变化、显示出来或被关闭时通过构造函数传入的回调通知。
        /// </summary>
        public IVsWindowFrame ActiveFrame { get; private set; }

        public void Dispose()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            SetActiveFrame(null);
            if (_selectionCookie != VSConstants.VSCOOKIE_NIL)
            {
                _monitorSelection.UnadviseSelectionEvents(_selectionCookie);
                _selectionCookie = VSConstants.VSCOOKIE_NIL;
            }
        }

        int IVsSelectionEvents.OnElementValueChanged(uint elementid, object varValueOld, object varValueNew)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // 新的活动文档有时通过 SEID_WindowFrame 而不是 SEID_DocumentFrame 通知；
            // SEID_WindowFrame 对工具窗口也会触发，所以要用窗口类型过滤：工具窗口不改变活动文档
            if (IsFrameElement(elementid) && varValueNew is IVsWindowFrame frame && IsDocumentFrame(frame))
            {
                SetActiveFrame(frame);
                _onChanged();
            }

            return VSConstants.S_OK;
        }

        int IVsWindowFrameNotify.OnShow(int fShow)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            switch ((__FRAMESHOW)fShow)
            {
                case __FRAMESHOW.FRAMESHOW_WinShown:
                    // 窗口显示出来了：延迟加载的文档数据此时已初始化，通知重新解析
                    _onChanged();
                    break;
                case __FRAMESHOW.FRAMESHOW_WinClosed:
                case __FRAMESHOW.FRAMESHOW_WinHidden:
                case __FRAMESHOW.FRAMESHOW_TabDeactivated:
                    // 活动文档窗口被关闭、隐藏或切走标签页：不再有活动文档，等下一次激活通知
                    SetActiveFrame(null);
                    _onChanged();
                    break;
            }

            return VSConstants.S_OK;
        }

        int IVsSelectionEvents.OnSelectionChanged(
            IVsHierarchy pHierOld, uint itemidOld, IVsMultiItemSelect pMISOld, ISelectionContainer pSCOld,
            IVsHierarchy pHierNew, uint itemidNew, IVsMultiItemSelect pMISNew, ISelectionContainer pSCNew)
        {
            return VSConstants.S_OK;
        }

        int IVsSelectionEvents.OnCmdUIContextChanged(uint dwCmdUICookie, int fActive)
        {
            return VSConstants.S_OK;
        }

        int IVsWindowFrameNotify.OnMove()
        {
            return VSConstants.S_OK;
        }

        int IVsWindowFrameNotify.OnSize()
        {
            return VSConstants.S_OK;
        }

        int IVsWindowFrameNotify.OnDockableChange(int fDockable)
        {
            return VSConstants.S_OK;
        }

        private static bool IsFrameElement(uint elementid)
        {
            return elementid == (uint)VSConstants.VSSELELEMID.SEID_DocumentFrame
                || elementid == (uint)VSConstants.VSSELELEMID.SEID_WindowFrame;
        }

        private static bool IsDocumentFrame(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            return ErrorHandler.Succeeded(frame.GetProperty((int)__VSFPROPID.VSFPROPID_Type, out object frameType))
                && frameType is int type
                && type == (int)__WindowFrameTypeFlags.WINDOWFRAMETYPE_Document;
        }

        private IVsWindowFrame GetCurrentDocumentFrame()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            int hr = _monitorSelection.GetCurrentElementValue((uint)VSConstants.VSSELELEMID.SEID_DocumentFrame, out object value);
            return ErrorHandler.Succeeded(hr) ? value as IVsWindowFrame : null;
        }

        // 切换活动文档窗口：先取消对旧窗口的登记，再对新窗口登记（frame 为 null 表示没有活动文档窗口）
        private void SetActiveFrame(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            UnwatchFrame();
            ActiveFrame = frame;

            if (frame is IVsWindowFrame2 frame2 && ErrorHandler.Succeeded(frame2.Advise(this, out uint cookie)))
            {
                _watchedFrame = frame2;
                _frameCookie = cookie;
            }
        }

        private void UnwatchFrame()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_watchedFrame == null)
            {
                return;
            }

            _watchedFrame.Unadvise(_frameCookie);
            _watchedFrame = null;
            _frameCookie = VSConstants.VSCOOKIE_NIL;
        }
    }
}

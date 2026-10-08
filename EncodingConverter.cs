using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;

namespace EncodingDisplayExtension
{
    /// <summary>
    /// 把文档按目标编码保存到磁盘：文本内容不变，只改写文件字节。
    /// 做法与 VS 自带的“带编码保存”一致：先设置 ITextDocument.Encoding，再走 VS 标准保存链路。
    /// </summary>
    internal sealed class EncodingConverter
    {
        private readonly IVsRunningDocumentTable _runningDocumentTable;

        public EncodingConverter(IVsRunningDocumentTable runningDocumentTable)
        {
            _runningDocumentTable = runningDocumentTable;
        }

        /// <summary>
        /// 转换成功返回 null；被拦截或失败时返回给用户看的原因，且文档编码保持原样。
        /// </summary>
        public string TryConvert(ITextDocument document, Encoding target)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string targetText = EncodingDisplayInfo.From(target).Text;
            string blockReason = GetBlockReason(document, target, targetText);
            if (blockReason != null)
            {
                return blockReason;
            }

            Encoding original = document.Encoding;
            document.Encoding = target;

            int hr = SaveToDisk(document.FilePath);
            if (ErrorHandler.Failed(hr))
            {
                // 保存失败时还原编码，避免状态栏显示磁盘上并不存在的编码
                document.Encoding = original;
                return $"Failed to save the file as {targetText} (HRESULT 0x{hr:X8}). The encoding was not changed.";
            }

            return GetOverrideWarning(document, targetText);
        }

        // 不满足转换前提时返回原因：文件还没写到磁盘，或转换会丢字符
        private static string GetBlockReason(ITextDocument document, Encoding target, string targetText)
        {
            if (!File.Exists(document.FilePath))
            {
                return "The file has not been saved to disk yet. Save it first, then change its encoding.";
            }

            if (!target.CanEncode(document.TextBuffer.CurrentSnapshot.GetText()))
            {
                return $"The file contains characters that cannot be represented in {targetText}. The encoding was not changed.";
            }

            return null;
        }

        // 通过 RDT 强制保存：文档未修改（只改了编码）也会写盘，并保留 VS 标准保存流程（只读、源代码管理提示等）
        private int SaveToDisk(string filePath)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            int hr = _runningDocumentTable.FindAndLockDocument(
                (uint)_VSRDTFLAGS.RDT_NoLock, filePath,
                out IVsHierarchy hierarchy, out uint itemId, out IntPtr docData, out uint cookie);

            // RDT 返回的 docData 带一次引用计数，用完必须释放
            if (docData != IntPtr.Zero)
            {
                Marshal.Release(docData);
            }

            if (ErrorHandler.Failed(hr))
            {
                return hr;
            }

            if (cookie == VSConstants.VSCOOKIE_NIL)
            {
                return VSConstants.E_FAIL; // 文档不在 RDT 中（例如已关闭）
            }

            return _runningDocumentTable.SaveDocuments(
                (uint)__VSRDTSAVEOPTIONS.RDTSAVEOPT_ForceSave, hierarchy, itemId, cookie);
        }

        // 保存时 .editorconfig 的 charset 等会覆盖所选编码，实际结果与所选不一致时要告知用户
        private static string GetOverrideWarning(ITextDocument document, string targetText)
        {
            string actualText = EncodingDisplayInfo.From(document.Encoding).Text;
            if (actualText == targetText)
            {
                return null;
            }

            return $"The file was saved as {actualText} instead of {targetText}. "
                + "A 'charset' rule in .editorconfig (or a Visual Studio option that forces the save encoding) overrides the encoding on save.";
        }
    }
}

using System.Text;

namespace EncodingDisplayExtension
{
    internal static class EncodingExtensions
    {
        /// <summary>
        /// 返回文本中第一个无法用该编码表示的字符的下标（按 UTF-16 单元计，代理对指向第一个单元）；
        /// 全部都能表示时返回 -1。无法表示的字符在保存时会丢失或被替换。
        /// </summary>
        public static int IndexOfUnencodable(this Encoding encoding, string text)
        {
            Encoding strictEncoding = Encoding.GetEncoding(
                encoding.CodePage, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);

            try
            {
                strictEncoding.GetByteCount(text);
                return -1;
            }
            catch (EncoderFallbackException ex)
            {
                return ex.Index;
            }
        }
    }
}

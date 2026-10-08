using System.Text;

namespace EncodingDisplayExtension
{
    internal static class EncodingExtensions
    {
        /// <summary>
        /// 文本中的字符是否都能用该编码表示。无法表示的字符在保存时会丢失或被替换。
        /// </summary>
        public static bool CanEncode(this Encoding encoding, string text)
        {
            Encoding strictEncoding = Encoding.GetEncoding(
                encoding.CodePage, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);

            try
            {
                strictEncoding.GetByteCount(text);
                return true;
            }
            catch (EncoderFallbackException)
            {
                return false;
            }
        }
    }
}

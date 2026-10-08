using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EncodingDisplayExtension
{
    /// <summary>
    /// 插件支持显示、也支持互相转换的 5 种编码（与 README 中“支持的编码格式显示示例”一致）。
    /// </summary>
    internal static class SupportedEncodings
    {
        public static IReadOnlyList<Encoding> All { get; } = new Encoding[]
        {
            new UTF8Encoding(false),          // UTF-8
            new UTF8Encoding(true),           // UTF-8 BOM
            Encoding.GetEncoding(936),        // GB2312/GBK
            new UnicodeEncoding(false, true), // UTF-16 LE（带 BOM，与 Encoding.Unicode 一致）
            Encoding.ASCII                    // US-ASCII
        };

        /// <summary>
        /// 判断编码是否属于这 5 种。与状态栏使用同一套显示文字比较，保证“能显示的就是能转换的”。
        /// </summary>
        public static bool Contains(Encoding encoding)
        {
            string text = EncodingDisplayInfo.From(encoding).Text;
            return All.Any(supported => EncodingDisplayInfo.From(supported).Text == text);
        }
    }
}

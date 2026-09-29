using System.Text;
using System.Windows.Media;

namespace EncodingDisplayExtension
{
    /// <summary>
    /// 状态栏要显示的编码信息（文本 + 前景色）。只做纯计算，不依赖 VS 服务。
    /// </summary>
    internal sealed class EncodingDisplayInfo
    {
        // 颜色使用在深色和浅色主题下都较清晰的颜色；画刷冻结后可反复复用，避免每次刷新都新建
        private static readonly Brush Utf8Brush = CreateFrozenBrush(0x60, 0xA0, 0x60);  // 柔和绿色
        private static readonly Brush GbBrush = CreateFrozenBrush(0xE0, 0xA0, 0x30);    // 琥珀色警告
        private static readonly Brush AsciiBrush = CreateFrozenBrush(0x50, 0x90, 0xD0); // 天蓝色
        private static readonly Brush Utf16Brush = CreateFrozenBrush(0xA0, 0x70, 0xD0); // 紫色
        private static readonly Brush OtherBrush = CreateFrozenBrush(0xE0, 0x60, 0x40); // 橙红色警告

        private EncodingDisplayInfo(string text, Brush foreground)
        {
            Text = text;
            Foreground = foreground;
        }

        public string Text { get; }

        public Brush Foreground { get; }

        public static EncodingDisplayInfo From(Encoding encoding)
        {
            // 用固定区域性转大写，避免 tr-TR 等区域性下 "i" 被转成 "İ" 导致名称匹配失败
            string name = encoding.WebName.ToUpperInvariant();
            return new EncodingDisplayInfo(GetText(name, encoding), GetForeground(name, encoding.CodePage));
        }

        // 区分 UTF-8 和 UTF-8 with BOM
        private static string GetText(string name, Encoding encoding)
        {
            bool isUtf8WithBom = name == "UTF-8" && encoding.GetPreamble().Length > 0;
            return isUtf8WithBom ? "UTF-8 BOM" : name;
        }

        // 颜色映射规则，按顺序匹配
        private static Brush GetForeground(string name, int codePage)
        {
            if (name == "UTF-8")
            {
                return Utf8Brush;
            }

            if (name.Contains("GB") || codePage == 936) // GB2312, GBK, GB18030
            {
                return GbBrush;
            }

            if (name.Contains("ASCII"))
            {
                return AsciiBrush;
            }

            if (name.Contains("UTF-16") || name.Contains("UNICODE"))
            {
                return Utf16Brush;
            }

            return OtherBrush;
        }

        private static Brush CreateFrozenBrush(byte red, byte green, byte blue)
        {
            var brush = new SolidColorBrush(Color.FromRgb(red, green, blue));
            brush.Freeze();
            return brush;
        }
    }
}

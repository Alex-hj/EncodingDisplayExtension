using System.Windows;
using System.Windows.Media;

namespace EncodingDisplayExtension
{
    internal static class VisualTreeExtensions
    {
        /// <summary>
        /// 在可视化树中深度优先查找第一个 T 类型的子元素；找不到（或 parent 为 null）返回 null。
        /// </summary>
        public static T FindChild<T>(this DependencyObject parent) where T : DependencyObject
        {
            if (parent == null)
            {
                return null;
            }

            int childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T typedChild)
                {
                    return typedChild;
                }

                var result = child.FindChild<T>();
                if (result != null)
                {
                    return result;
                }
            }

            return null;
        }
    }
}

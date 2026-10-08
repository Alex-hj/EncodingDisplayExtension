using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using Microsoft.VisualStudio.Shell;

namespace EncodingDisplayExtension
{
    /// <summary>
    /// 注入到 VS 状态栏右侧的编码显示项。
    /// </summary>
    internal sealed class EncodingStatusBarItem
    {
        private const string ItemName = "EncodingDisplayItem";
        private const string DefaultToolTip = "Current File Encoding";
        private const string ClickableToolTip = "Current File Encoding (click to change)";

        private readonly TextBlock _textBlock;
        private readonly Action _onClicked;

        // 当前显示的编码是否允许点击修改
        private bool _canChange;

        private EncodingStatusBarItem(TextBlock textBlock, Action onClicked)
        {
            _textBlock = textBlock;
            _onClicked = onClicked;
            _textBlock.MouseLeftButtonUp += OnMouseLeftButtonUp;
        }

        /// <summary>
        /// 在主窗口状态栏注入（或复用已有的）编码显示项；找不到主窗口或状态栏时返回 null。
        /// onClicked 在用户点击一个允许修改的编码时调用。
        /// </summary>
        public static EncodingStatusBarItem TryInject(Window mainWindow, Action onClicked)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                var statusBar = mainWindow.FindChild<StatusBar>();
                return statusBar == null ? null : new EncodingStatusBarItem(FindOrAddTextBlock(statusBar), onClicked);
            }
            catch (Exception ex)
            {
                ActivityLog.TryLogError(nameof(EncodingStatusBarItem), $"Error injecting UI: {ex}");
                return null;
            }
        }

        /// <summary>
        /// 显示编码；canChange 为 true 表示该编码允许点击修改（显示手型光标和提示）。
        /// </summary>
        public void Show(EncodingDisplayInfo info, bool canChange)
        {
            _textBlock.Text = info.Text;
            _textBlock.Foreground = info.Foreground;
            SetClickable(canChange);
        }

        public void Clear()
        {
            _textBlock.Text = string.Empty;
            SetClickable(false);
        }

        /// <summary>
        /// 编码项左上角的屏幕坐标（设备像素），作为弹出菜单的位置。
        /// </summary>
        public Point ScreenPosition => _textBlock.PointToScreen(new Point(0, 0));

        private void SetClickable(bool canChange)
        {
            _canChange = canChange;
            _textBlock.Cursor = canChange ? Cursors.Hand : null;
            _textBlock.ToolTip = canChange ? ClickableToolTip : DefaultToolTip;
        }

        private void OnMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_canChange)
            {
                _onClicked();
            }
        }

        private static TextBlock FindOrAddTextBlock(StatusBar statusBar)
        {
            // 检查是否已经添加过，防止重复；找到旧的就复用它的控件
            foreach (var child in statusBar.Items)
            {
                if (child is StatusBarItem item && item.Name == ItemName && item.Content is TextBlock existing)
                {
                    return existing;
                }
            }

            var textBlock = CreateTextBlock();

            // 插入到 Items 集合中。
            // 在 DockPanel 中，对于 Dock.Right 的元素：
            // 列表前面的元素会被推到最右边，列表后面的元素会往左排。
            // 为了让它显示在 WakaTime (通常是最右侧) 的左边，我们直接 Add 到末尾即可。
            statusBar.Items.Add(CreateItem(textBlock));
            return textBlock;
        }

        private static TextBlock CreateTextBlock()
        {
            return new TextBlock
            {
                Text = "Loading...",
                Margin = new Thickness(10, 0, 10, 0), // 左右留点空隙
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = SystemColors.WindowTextBrush, // 首次刷新前的占位色，刷新后由编码对应的颜色覆盖
                ToolTip = DefaultToolTip
            };
        }

        private static StatusBarItem CreateItem(TextBlock textBlock)
        {
            var item = new StatusBarItem
            {
                Content = textBlock,
                Name = ItemName // 给个名字防止重复添加
            };

            // 关键布局设置：停靠在右侧
            DockPanel.SetDock(item, Dock.Right);
            return item;
        }
    }
}

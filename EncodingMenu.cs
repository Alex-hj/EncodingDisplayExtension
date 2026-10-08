using System;
using System.ComponentModel.Design;
using System.Text;
using System.Windows;
using Microsoft.VisualStudio.Shell;

namespace EncodingDisplayExtension
{
    /// <summary>
    /// 点击状态栏编码项后弹出的编码选择菜单：VS 原生上下文菜单（定义见 EncodingMenu.vsct），
    /// 列出 SupportedEncodings，当前编码打勾。
    /// </summary>
    internal sealed class EncodingMenu
    {
        // 以下 3 个值必须与 EncodingMenu.vsct 一致；编码按钮的 ID 从 FirstEncodingCommandId 起与 SupportedEncodings.All 一一对应
        private static readonly Guid CommandSetGuid = new Guid("35027143-ada8-4353-b0e2-7f159477159a");
        private const int ContextMenuId = 0x1000;
        private const int FirstEncodingCommandId = 0x0100;

        private readonly OleMenuCommandService _commandService;
        private readonly Action<Encoding> _onSelected;

        // 最近一次弹出菜单时的当前编码显示文字，用于打勾
        private string _currentText;

        private EncodingMenu(OleMenuCommandService commandService, Action<Encoding> onSelected)
        {
            _commandService = commandService;
            _onSelected = onSelected;
        }

        /// <summary>
        /// 注册菜单里的编码命令；命令服务不可用时返回 null。
        /// onSelected 在用户选中一个不同于当前编码的编码时调用。
        /// </summary>
        public static EncodingMenu Create(OleMenuCommandService commandService, Action<Encoding> onSelected)
        {
            if (commandService == null)
            {
                return null;
            }

            var menu = new EncodingMenu(commandService, onSelected);
            menu.RegisterCommands();
            return menu;
        }

        /// <summary>
        /// 在屏幕坐标（设备像素）处弹出菜单，current 为当前编码。
        /// </summary>
        public void Show(Point screenPosition, Encoding current)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            _currentText = EncodingDisplayInfo.From(current).Text;
            var menuId = new CommandID(CommandSetGuid, ContextMenuId);
            _commandService.ShowContextMenu(menuId, (int)screenPosition.X, (int)screenPosition.Y);
        }

        private void RegisterCommands()
        {
            for (int i = 0; i < SupportedEncodings.All.Count; i++)
            {
                var command = new OleMenuCommand(OnCommandInvoked, new CommandID(CommandSetGuid, FirstEncodingCommandId + i));
                command.BeforeQueryStatus += OnBeforeQueryStatus;
                _commandService.AddCommand(command);
            }
        }

        // 菜单显示前：文字取自 EncodingDisplayInfo（与状态栏一致），当前编码打勾
        private void OnBeforeQueryStatus(object sender, EventArgs e)
        {
            var command = (OleMenuCommand)sender;
            command.Text = GetText(command);
            command.Checked = IsCurrent(command);
        }

        private void OnCommandInvoked(object sender, EventArgs e)
        {
            var command = (OleMenuCommand)sender;
            if (!IsCurrent(command))
            {
                _onSelected(GetEncoding(command));
            }
        }

        private static Encoding GetEncoding(MenuCommand command)
        {
            return SupportedEncodings.All[command.CommandID.ID - FirstEncodingCommandId];
        }

        private static string GetText(MenuCommand command)
        {
            return EncodingDisplayInfo.From(GetEncoding(command)).Text;
        }

        private bool IsCurrent(MenuCommand command)
        {
            return GetText(command) == _currentText;
        }
    }
}

using System;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using System.Resources;
using System.Reflection;
using System.Globalization;
using System.Collections.Generic;

namespace ToAPI
{
    /// <summary>用于实现外接程序的对象。</summary>
    /// <seealso class='IDTExtensibility2' />
    public class Connect : IDTExtensibility2, IDTCommandTarget
    {
        /// <summary>实现外接程序对象的构造函数。请将您的初始化代码置于此方法内。</summary>
        public Connect()
        {
            _myBars = new List<CommandBarControl>();
        }

        /// <summary>实现 IDTExtensibility2 接口的 OnConnection 方法。接收正在加载外接程序的通知。</summary>
        /// <param term='application'>宿主应用程序的根对象。</param>
        /// <param term='connectMode'>描述外接程序的加载方式。</param>
        /// <param term='addInInst'>表示此外接程序的对象。</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _dte = (DTE2)application;
            _addIn = (AddIn)addInInst;

            switch (connectMode)
            {
                case ext_ConnectMode.ext_cm_UISetup:
                    break;
                case ext_ConnectMode.ext_cm_Startup:
                    break;
                case ext_ConnectMode.ext_cm_AfterStartup:
                    createToDefineBar();
                    createFindUsedBar();
                    break;
            }
        }

        /// <summary>实现 IDTExtensibility2 接口的 OnDisconnection 方法。接收正在卸载外接程序的通知。</summary>
        /// <param term='disconnectMode'>描述外接程序的卸载方式。</param>
        /// <param term='custom'>特定于宿主应用程序的参数数组。</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
            switch (disconnectMode)
            {
                case ext_DisconnectMode.ext_dm_HostShutdown:
                case ext_DisconnectMode.ext_dm_UserClosed:
                    removeBar();
                    break;
            }
        }

        /// <summary>实现 IDTExtensibility2 接口的 OnAddInsUpdate 方法。当外接程序集合已发生更改时接收通知。</summary>
        /// <param term='custom'>特定于宿主应用程序的参数数组。</param>
        /// <seealso class='IDTExtensibility2' />		
        public void OnAddInsUpdate(ref Array custom)
        {
        }

        /// <summary>实现 IDTExtensibility2 接口的 OnStartupComplete 方法。接收宿主应用程序已完成加载的通知。</summary>
        /// <param term='custom'>特定于宿主应用程序的参数数组。</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnStartupComplete(ref Array custom)
        {
            createToDefineBar();
            createFindUsedBar();
        }

        /// <summary>实现 IDTExtensibility2 接口的 OnBeginShutdown 方法。接收正在卸载宿主应用程序的通知。</summary>
        /// <param term='custom'>特定于宿主应用程序的参数数组。</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnBeginShutdown(ref Array custom)
        {
        }

        /// <summary>实现 IDTCommandTarget 接口的 QueryStatus 方法。此方法在更新该命令的可用性时调用</summary>
        /// <param term='commandName'>要确定其状态的命令的名称。</param>
        /// <param term='neededText'>该命令所需的文本。</param>
        /// <param term='status'>该命令在用户界面中的状态。</param>
        /// <param term='commandText'>neededText 参数所要求的文本。</param>
        /// <seealso class='Exec' />
        public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
        {
            if (neededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
            {
                switch (commandName)
                {
                    case "ToAPI.Connect.ToDefine":
                    case "ToAPI.Connect.FindUsedSimple":
                    case "ToAPI.Connect.FindUsedExact":
                        status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                        return;
                }
            }
        }

        /// <summary>实现 IDTCommandTarget 接口的 Exec 方法。此方法在调用该命令时调用。</summary>
        /// <param term='commandName'>要执行的命令的名称。</param>
        /// <param term='executeOption'>描述该命令应如何运行。</param>
        /// <param term='varIn'>从调用方传递到命令处理程序的参数。</param>
        /// <param term='varOut'>从命令处理程序传递到调用方的参数。</param>
        /// <param term='handled'>通知调用方此命令是否已被处理。</param>
        /// <seealso class='Exec' />
        public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
        {
            handled = false;
            if (executeOption == vsCommandExecOption.vsCommandExecOptionDoDefault)
            {
                switch(commandName)
                {
                    case "ToAPI.Connect.ToDefine":
                        new ToDefineBiz(_dte).ToDefine();
                        break;
                    case "ToAPI.Connect.FindUsedSimple":
                        new FindUsedBiz(_dte).FindUsed(false);
                        break;
                    case "ToAPI.Connect.FindUsedExact":
                        new FindUsedBiz(_dte).FindUsed(true);
                        break;
                }

                handled = true;
                return;
            }
        }

        private DTE2 _dte;
        private AddIn _addIn;

        #region

        private List<CommandBarControl> _myBars;

        private Command findCmd(string cmdName)
        {
            try
            {
                return ((Commands2)_dte.Commands).Item(_addIn.ProgID + "." + cmdName, -1);
            }
            catch
            {
                return null;
            }
        }

        private void createToDefineBar()
        {
            object[] contextGUIDS = new object[] { };
            var commands = (Commands2)_dte.Commands;
            var commandBars = (CommandBars)_dte.CommandBars;

            // js和aspx编辑器右键菜单
            var ctxJs = commandBars["Script Context"];
            var ctxAspx = commandBars["ASPX Context"];

            // 将命令添加到 Commands 集合
            Command toDefine = findCmd("ToDefine") ??
                commands.AddNamedCommand2(_addIn, "ToDefine", "转到ServiceAPI", "查看ServiceAPI",
                    true, 526, ref contextGUIDS,
                    (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled,
                    (int)vsCommandStyle.vsCommandStylePictAndText,
                    vsCommandControlType.vsCommandControlTypeButton);

            // 将对应于该命令的控件添加到菜单
            if (ctxJs != null)
            {
                _myBars.Add((CommandBarButton)toDefine.AddControl(ctxJs, 1));
            }
            if (ctxAspx != null)
            {
                _myBars.Add((CommandBarButton)toDefine.AddControl(ctxAspx, 1));
            }
        }

        private void createFindUsedBar()
        {
            object[] contextGUIDS = new object[] { };
            var commands = (Commands2)_dte.Commands;
            var commandBars = (CommandBars)_dte.CommandBars;

            // 代码编辑器右键菜单
            CommandBar ctxCode = commandBars["Code Window"];

            // 将命令添加到 Commands 集合
            Command findUsedSimple = findCmd("FindUsedSimple") ??
                commands.AddNamedCommand2(_addIn, "FindUsedSimple", "不那么准确的查找", "仅按方法名查找",
                    true, 109, ref contextGUIDS,
                    (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled,
                    (int)vsCommandStyle.vsCommandStylePictAndText,
                    vsCommandControlType.vsCommandControlTypeButton);

            Command findUsedExact = findCmd("FindUsedExact") ??
                commands.AddNamedCommand2(_addIn, "FindUsedExact", "比较准确的查找", "按类名+方法名查找",
                    true, 172, ref contextGUIDS,
                    (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled,
                    (int)vsCommandStyle.vsCommandStylePictAndText,
                    vsCommandControlType.vsCommandControlTypeButton);

            // 将对应于该命令的控件添加到菜单
            if (ctxCode != null)
            {
                // 杯具了，AddCommandBar是永久添加菜单。。。
                //var findCtx = (CommandBar)commands.AddCommandBar("查找ServiceAPI引用", vsCommandBarType.vsCommandBarTypeMenu, ctx, 1);

                // 添加一级菜单
                var findUsedBar = (CommandBarPopup)ctxCode.Controls.Add(MsoControlType.msoControlPopup, Type.Missing, Type.Missing, 1, true);
                findUsedBar.Caption = "查找ServiceAPI引用";
                findUsedBar.CommandBar.Name = "ToAPI FindUsed";

                // 添加二级菜单
                findUsedExact.AddControl(findUsedBar.CommandBar, 1);
                findUsedSimple.AddControl(findUsedBar.CommandBar, 2);

                _myBars.Add(findUsedBar);
            }
        }

        private void removeBar()
        {
            _myBars.ForEach(bar =>
            {
                if (bar != null)
                {
                    bar.Delete(true);
                }
            });
        }

        #endregion
    }
}
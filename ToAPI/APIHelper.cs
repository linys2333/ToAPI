using System;
using System.Collections.Generic;
using System.Text;
using EnvDTE;
using EnvDTE80;
using System.Text.RegularExpressions;
using System.IO;
using System.Windows.Forms;

namespace ToAPI
{
    /// <summary>
    /// Service相关信息
    /// </summary>
    public class ServiceFunc
    {
        /// <summary>
        /// 命名空间：Mysoft.Tzsy.Services.ProjCompile
        /// </summary>
        public string Namespace { get; set; }

        /// <summary>
        /// 类名：SaleService
        /// </summary>
        public string Class { get; set; }

        /// <summary>
        /// 方法名：Save
        /// </summary>
        public string Func { get; set; }

        /// <summary>
        /// 参数个数
        /// </summary>
        public int ParamNum { get; set; }

        /// <summary>
        /// 返回值类型
        /// </summary>
        public bool IsPublic { get; set; }

        /// <summary>
        /// 方法完全限定名：Mysoft.Tzsy.Services.ProjCompile.SaleService.Save
        /// </summary>
        public string FullName
        {
            get
            {
                if (string.IsNullOrEmpty(Namespace) ||
                    string.IsNullOrEmpty(Class) ||
                    string.IsNullOrEmpty(Func))
                {
                    return "";
                }

                return string.Format(@"{0}.{1}.{2}", Namespace, Class, Func);
            }
        }

        /// <summary>
        /// 文件相对路径：\Mysoft.Tzsy.Services\ProjCompile\SaleService.cs
        /// </summary>
        public string FilePath
        {
            get
            {
                if (string.IsNullOrEmpty(Namespace) ||
                    string.IsNullOrEmpty(Class) ||
                    string.IsNullOrEmpty(Func))
                {
                    return "";
                }

                return string.Format(@"\{0}\{1}", 
                    Regex.Replace(Namespace, @"(?<=Services)\..+", m => m.Value.Replace(".", @"\")), 
                    Class);
            }
        }
    }

    /// <summary>
    /// ServiceAPI相关操作
    /// </summary>
    public class APIHelper
    {
        private ServiceFunc _serviceFunc;
        private DTE2 _dte;

        public APIHelper(DTE2 dte)
        {
            _dte = dte;
            _serviceFunc = new ServiceFunc();
        }

        /// <summary>
        /// 转到定义
        /// </summary>
        public void ToDefine()
        {
            // 解析
            analyzeInfo(GetSelection());

            // 校验
            if (string.IsNullOrEmpty(_serviceFunc.Class) ||
                string.IsNullOrEmpty(_serviceFunc.Func))
            {
                return;
            }

            if (string.IsNullOrEmpty(_serviceFunc.Namespace))
            {
                MessageBox.Show("该方法缺少js引用！");
                return;
            }

            // 查找
            string path = getPath();
            if (string.IsNullOrEmpty(path))
            {
                MessageBox.Show("未找到后端文件！");
                return;
            }

            Window win = _dte.ItemOperations.OpenFile(path, Constants.vsViewKindPrimary);
            if (!toFunc(win))
            {
                MessageBox.Show("未找到匹配方法！");
                return;
            }
        }

        /// <summary>
        /// 查找引用
        /// </summary>
        /// <param name="isMatchWhole">是否完全匹配</param>
        public void FindUsed(bool isMatchWhole)
        {
            // 解析
            setInfoByCodeModel();

            // 校验
            if (string.IsNullOrEmpty(_serviceFunc.FullName))
            {
                return;
            }

            if (!_serviceFunc.Namespace.Contains("Services") ||
                !_serviceFunc.Class.EndsWith("Service") ||
                !_serviceFunc.IsPublic)
            {
                return;
            }

            // 查找
            execFind(isMatchWhole);
        }

        /// <summary>
        /// 获取触发行的代码
        /// </summary>
        /// <returns></returns>
        protected virtual string GetSelection()
        {
            // 触发起始点
            var selection = (TextSelection)_dte.ActiveDocument.Selection;
            EditPoint editPoint = selection.AnchorPoint.CreateEditPoint();

            // 选中触发行
            // TODO：匹配多行
            string code = editPoint.GetLines(editPoint.Line, editPoint.Line + 1);

            return code;
        }

        /// <summary>
        /// 解析代码设置目标对象
        /// </summary>
        /// <param name="code"></param>
        private void analyzeInfo(string code)
        {
            string service = Regex.Match(code, @"\b\w+(?=Service\.)").Value;
            string js = getJs(service);

            _serviceFunc.Namespace = Regex.Match(js, string.Format(@"(?<=/service/)[\w\.]+(?=\.{0}/)", service)).Value;
            _serviceFunc.Class = service + "Service";
            _serviceFunc.Func = Regex.Match(code, @"(?<=Service\.)\w+\b").Value;

            // 采用正则平衡组匹配嵌套参数
            string param = Regex.Match(code,
                string.Format(@"(?<={0}\()(?>[^\(\)]+|\((?<sign>)|\)(?<-sign>))*(?(sign)(?!))(?=\))", _serviceFunc.Func)).Value;

            if(string.IsNullOrEmpty(param))
            {
                _serviceFunc.ParamNum = 0;
            }
            else
            {
                param = filterParam(param);
                _serviceFunc.ParamNum = param.Split(',').Length;
            }
        }

        /// <summary>
        /// 根据service名称获取js引用代码
        /// </summary>
        /// <param name="service"></param>
        /// <returns></returns>
        private string getJs(string service)
        {
            string path = _dte.ActiveDocument.Path + _dte.ActiveDocument.Name.Split('.')[0] + ".aspx";

            using (var sw = new StreamReader(path, Encoding.UTF8))
            {
                string text = sw.ReadToEnd();
                string pattern = string.Format(@"src=\S+\.{0}/Scripts.aspx", service);
                return Regex.Match(text, pattern, RegexOptions.IgnoreCase).Value;
            }
        }

        /// <summary>
        /// 根据触发文件的路径获取代码根路径
        /// </summary>
        /// <returns></returns>
        private string getPath()
        {
            string rootPath = Regex.Match(_dte.ActiveDocument.Path, @"^.+\\ERP\\明源整体解决方案(?=\\Map)", RegexOptions.IgnoreCase).Value;
            string path = rootPath + _serviceFunc.FilePath;
            string extName = ".cs";

            if (!File.Exists(path + extName))
            {
                path += "s";

                if (!File.Exists(path + extName))
                {
                    return "";
                }
            }
            return path + extName;
        }

        /// <summary>
        /// 递归过滤掉嵌套参数
        /// </summary>
        /// <param name="param"></param>
        /// <returns></returns>
        private string filterParam(string param)
        {
            string filter = Regex.Replace(param, @"\([^\(\)]*\)", m => "");
            if (param != filter)
            {
                return filterParam(filter);
            }
            return filter;
        }

        /// <summary>
        /// 匹配方法
        /// </summary>
        /// <returns></returns>
        private bool toFunc(Window win)
        {
            // 匹配方法定义代码（vs的正则表达式有点特殊）
            string pattern = "";                
            switch (_serviceFunc.ParamNum)
            { 
                case 0:
                    pattern = string.Format(@"public JsonResult {0}\(\)", _serviceFunc.Func);
                    break;
                case 1:
                    pattern = string.Format(@"public JsonResult {0}\([:a \t\n]+\)", _serviceFunc.Func);
                    break;
                default:
                    pattern = string.Format(@"public JsonResult {0}\(([:a \t\n]+,)^{1}[:a \t\n]+\)", _serviceFunc.Func, _serviceFunc.ParamNum - 1);
                    break;
            }

            var doc = (TextDocument)win.Document.Object("TextDocument");

            EditPoint editPoint = doc.StartPoint.CreateEditPoint();
            EditPoint endPoint = null;
            TextRanges tags = null;

            bool find = editPoint.FindPattern(pattern, (int)(vsFindOptions.vsFindOptionsFromStart |
                vsFindOptions.vsFindOptionsRegularExpression), ref endPoint, ref tags);

            if (find && tags != null)
            {
                doc.Selection.MoveToPoint(endPoint, false);
                doc.Selection.MoveToPoint(editPoint, true);
            }

            return find;
        }

        /// <summary>
        /// 利用CodeModel设置目标对象
        /// </summary>
        /// <param name="code"></param>
        private void setInfoByCodeModel()
        {
            var codeModel = (FileCodeModel2)_dte.ActiveDocument.ProjectItem.FileCodeModel;
            var selection = (TextSelection)_dte.ActiveDocument.Selection;

            try
            {
                // 触发的ServiceAPI
                var codeFunc = (CodeFunction2)codeModel.CodeElementFromPoint(selection.AnchorPoint, vsCMElement.vsCMElementFunction);
                var codeClass = (CodeClass2)codeFunc.Parent;

                _serviceFunc.Namespace = codeClass.Namespace.Name;
                _serviceFunc.Class = codeClass.Name;
                _serviceFunc.Func = codeFunc.Name;
                _serviceFunc.ParamNum = codeFunc.Parameters.Count;
                _serviceFunc.IsPublic = codeFunc.Access == vsCMAccess.vsCMAccessPublic;
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                // 没找到对应元素
                throw;
            }
        }

        /// <summary>
        /// 执行查找
        /// </summary>
        /// <param name="isMatchWhole">是否完全匹配</param>
        private void execFind(bool isMatchWhole)
        {
            string pattern = string.Format("{0}.{1}", _serviceFunc.Class, _serviceFunc.Func);

            var findWin = (Find2)_dte.Find;

            findWin.Action = vsFindAction.vsFindActionFindAll;
            findWin.Backwards = false;
            findWin.FilesOfType = "*.aspx;*.js";
            findWin.FindWhat = pattern;
            findWin.MatchCase = false;
            findWin.MatchInHiddenText = true;
            findWin.MatchWholeWord = true;
            findWin.PatternSyntax = vsFindPatternSyntax.vsFindPatternSyntaxLiteral;
            findWin.ResultsLocation = vsFindResultsLocation.vsFindResults2;
            findWin.SearchSubfolders = true;
            findWin.Target = vsFindTarget.vsFindTargetSolution;
            findWin.WaitForFindToComplete = false;

            if (isMatchWhole)
            {
                pattern = pattern.Replace(".", @"\.");
                switch (_serviceFunc.ParamNum)
                {
                    case 0:
                        pattern = string.Format(@"{0}\(\)", pattern);
                        break;
                    case 1:
                        pattern = string.Format(@"{0}\([:a \t\n]+\)", pattern);
                        break;
                    default:
                        pattern = string.Format(@"{0}\(([:a \t\n]+,)^{1}[:a \t\n]+\)", pattern, _serviceFunc.ParamNum - 1);
                        break;
                }

                findWin.FilesOfType = "*.js";
                findWin.FindWhat = pattern;
                findWin.MatchCase = true;
                findWin.PatternSyntax = vsFindPatternSyntax.vsFindPatternSyntaxRegExpr;
            }

            findWin.Execute();

            // 还原部分设置，防止后续误操作导致卡顿
            findWin.PatternSyntax = vsFindPatternSyntax.vsFindPatternSyntaxLiteral;
            findWin.Target = vsFindTarget.vsFindTargetCurrentDocument;
        }
    }
}

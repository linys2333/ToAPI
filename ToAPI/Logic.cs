using System;
using System.Collections.Generic;
using System.Text;
using EnvDTE;
using EnvDTE80;
using System.Text.RegularExpressions;
using System.IO;

namespace ToAPI
{
    /// <summary>
    /// Service相关信息
    /// </summary>
    public class ServiceInfo
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
        /// 方法名
        /// </summary>
        public string Func { get; set; }

        /// <summary>
        /// 参数个数
        /// </summary>
        public int ParamNum { get; set; }

        /// <summary>
        /// 方法完全限定名
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

                return string.Format(@"\{0}\{1}.cs", 
                    Regex.Replace(Namespace, @"(?<=Services)\..+", m => m.Value.Replace(".", @"\")), 
                    Class);
            }
        }
    }

    /// <summary>
    /// 功能实现
    /// </summary>
    public class Logic
    {
        private ServiceInfo _serviceInfo;
        private DTE2 _dte;

        public Logic(DTE2 dte)
        {
            _dte = dte;
            _serviceInfo = new ServiceInfo();
        }

        /// <summary>
        /// 转到定义
        /// </summary>
        public void ToAPI()
        {
            setInfo(getCode());

            string path = _serviceInfo.FilePath;
            if (string.IsNullOrEmpty(path)) 
            {
                return;
            }

            // 在默认查看器中打开指定文件
            Window win = _dte.ItemOperations.OpenFile(getRootPath() + path, Constants.vsViewKindPrimary);
            toFunc(win);
        }

        /// <summary>
        /// 获取触发行的代码
        /// </summary>
        /// <returns></returns>
        private string getCode()
        {
            // 触发对象
            var selection = (TextSelection)_dte.ActiveDocument.Selection;
            int startLine = selection.AnchorPoint.Line;
            int startOffset = selection.AnchorPoint.LineCharOffset;
            
            // 选中触发行
            // TODO：匹配多行
            selection.SelectLine();
            string code = selection.Text;

            // 还原触发点
            selection.MoveToLineAndOffset(startLine, startOffset, false);

            return code;
        }

        /// <summary>
        /// 根据代码设置目标对象
        /// </summary>
        /// <param name="code"></param>
        private void setInfo(string code)
        {
            string service = Regex.Match(code, @"(?<=\= ?)\w+(?=Service\.)").Value;
            string js = getJs(service);

            _serviceInfo.Namespace = Regex.Match(js, string.Format(@"(?<=/service/)[\w\.]+(?=\.{0}/)", service)).Value;
            _serviceInfo.Class = service + "Service";
            _serviceInfo.Func = Regex.Match(code, @"(?<=Service\.)\w+(?=\()").Value;

            // 采用正则平衡组匹配嵌套参数
            string param = Regex.Match(code,
                string.Format(@"(?<={0}\()(?>[^\(\)]+|\((?<sign>)|\)(?<-sign>))*(?(sign)(?!))(?=\))", _serviceInfo.Func)).Value;

            if(string.IsNullOrEmpty(param))
            {
                _serviceInfo.ParamNum = 0;
            }
            else
            {
                param = filterParam(param);
                _serviceInfo.ParamNum = param.Split(',').Length;
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
        private string getRootPath()
        {
            return Regex.Match(_dte.ActiveDocument.Path, @"^.+\\ERP\\明源整体解决方案(?=\\Map)", RegexOptions.IgnoreCase).Value;
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
        private void toFunc(Window win)
        {
            var doc = (TextDocument)win.Document.Object("TextDocument");
            doc.Selection.StartOfDocument(false);

            // 匹配方法定义代码（vs的正则表达式有点特殊）
            string pattern = "";                
            switch (_serviceInfo.ParamNum)
            { 
                case 0:
                    pattern = string.Format(@"public JsonResult {0}\(\)", _serviceInfo.Func);
                    break;
                case 1:
                    pattern = string.Format(@"public JsonResult {0}\([:a \t\n]+\)", _serviceInfo.Func);
                    break;
                default:
                    pattern = string.Format(@"public JsonResult {0}\(([:a \t\n]+,)^{1}[:a \t\n]+\)", _serviceInfo.Func, _serviceInfo.ParamNum - 1);
                    break;
            }
            bool find = false;

            try
            {
                doc.Selection.SelectAll();
                find = doc.Selection.FindText(pattern, (int)(vsFindOptions.vsFindOptionsFromStart | 
                    vsFindOptions.vsFindOptionsRegularExpression | 
                    vsFindOptions.vsFindOptionsMatchInHiddenText));
            }
            catch
            {
                find = false;
                throw;
            }
            finally
            {
                if (!find)
                {
                    doc.Selection.StartOfDocument(false);
                }
            }
        }
    }
}

using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;

namespace ToAPI
{
    /// <summary>
    /// ServiceAPI相关操作
    /// </summary>
    public class ToDefineBiz
    {
        private ServiceFunc _serviceFunc;
        private DTE2 _dte;

        public ToDefineBiz(DTE2 dte)
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
            string selection = new PubBiz(_dte).GetSelection();
            if (Regex.IsMatch(selection, @"Mysoft\.\w+\.Services\."))
            {
                // 包含“命名空间”的识别为旧代码
                analyzeInfo4Old(selection);
            }
            else
            {
                analyzeInfo(selection);
            }

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
                if (_serviceFunc.ParamNum > -1)
                {
                    // 如果带参数匹配不到，那么仅匹配方法名
                    _serviceFunc.ParamNum = -1;

                    if (toFunc(win))
                    {
                        MessageBox.Show("未找到匹配方法，已定位到同名方法！");
                        return;
                    }
                }

                MessageBox.Show("未找到匹配方法！");
                return;
            }
        }

        #region 私有方法

        /// <summary>
        /// 解析代码设置目标对象
        /// </summary>
        /// <param name="code">例如：var oResult = SaleService.IsExistSaleRule(sPdtInfoJson);</param>
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
        /// 解析代码设置目标对象（适应旧写法）
        /// </summary>
        /// <param name="code">例如：serviceInfo: "Mysoft.Tzsy.Services.BusinessParamServices.CorePlanRuleConfigService.Save",</param>
        private void analyzeInfo4Old(string code)
        {
            // 截取引号内代码
            string service = Regex.Match(code, @"(?<=[""|'])Mysoft\.[\w|\.]+\.\w+(?=[""|'])").Value;
            string[] serviceArr = service.Split('.');
            if (serviceArr.Length < 3)
            {
                return;
            }

            _serviceFunc.Class = serviceArr[serviceArr.Length - 2];
            _serviceFunc.Func = serviceArr[serviceArr.Length - 1];
            _serviceFunc.Namespace = service.Replace("." + _serviceFunc.Class + "." + _serviceFunc.Func, "");
            _serviceFunc.ParamNum = -1;
        }

        /// <summary>
        /// 根据service名称获取js引用代码
        /// </summary>
        /// <param name="service"></param>
        /// <returns></returns>
        private string getJs(string service)
        {
            string path = _dte.ActiveDocument.Path + _dte.ActiveDocument.Name.Split('.')[0];
            string fullPath = path + ".aspx";

            if (!File.Exists(fullPath))
            {
                // 如果没有对应aspx文件，那么查找当前文件中是否存在相关信息
                fullPath = path + ".js";

                if (!File.Exists(fullPath))
                {
                    return "";
                }
            }

            using (var sw = new StreamReader(fullPath, Encoding.UTF8))
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
            string rootPath = Regex.Match(_dte.ActiveDocument.Path, @"^.+\\明源整体解决方案(?=\\Map)", RegexOptions.IgnoreCase).Value;
            string path = rootPath + _serviceFunc.FilePath;
            string extName = ".cs";

            if (!File.Exists(path + extName))
            {
                // 适应不规范的写法
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
            string pattern = @":a_ \.\<\>\t\n";                
            switch (_serviceFunc.ParamNum)
            {
                // -1不匹配参数个数
                case -1:
                    pattern = string.Format(@"public :c+ {0}\([^\)]*\)", _serviceFunc.Func);
                    break;
                case 0:
                    pattern = string.Format(@"public :c+ {0}\(\)", _serviceFunc.Func);
                    break;
                case 1:
                    pattern = string.Format(@"public :c+ {1}\([{0}]+\)", pattern, _serviceFunc.Func);
                    break;
                default:
                    pattern = string.Format(@"public :c+ {1}\(([{0}]+,)^{2}[{0}]+\)", pattern, _serviceFunc.Func, _serviceFunc.ParamNum - 1);
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

        #endregion
    }
}

using EnvDTE;
using EnvDTE80;
using System.Text.RegularExpressions;

namespace ToAPI
{
    /// <summary>
    /// ServiceAPI相关操作
    /// </summary>
    public class PubBiz
    {
        private DTE2 _dte;

        public PubBiz(DTE2 dte)
        {
            _dte = dte;
        }

        /// <summary>
        /// 获取触发行的代码
        /// </summary>
        /// <returns></returns>
        public string GetSelection()
        {
            // 触发内容
            var selection = (TextSelection)_dte.ActiveDocument.Selection;
            string code = selection.Text;

            // 没有选中内容则获取当前行代码
            if (string.IsNullOrEmpty(code))
            {
                // 触发起始点
                EditPoint editPoint = selection.AnchorPoint.CreateEditPoint();

                // 选中触发行
                code = editPoint.GetLines(editPoint.Line, editPoint.Line + 1);

            }
            else
            {
                // 多行转成一行
                code = Regex.Replace(code, @"\s+", s => "");
            }

            // 处理回调函数
            code = Regex.Replace(code, @",\s*?function\b.*$", s => ")");

            return code;
        }
    }
}

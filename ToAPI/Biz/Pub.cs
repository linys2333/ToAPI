using EnvDTE;
using EnvDTE80;

namespace ToAPI
{
    /// <summary>
    /// ServiceAPI相关操作
    /// </summary>
    public class Pub
    {
        private DTE2 _dte;

        public Pub(DTE2 dte)
        {
            _dte = dte;
        }

        /// <summary>
        /// 获取触发行的代码
        /// </summary>
        /// <returns></returns>
        public string GetSelection()
        {
            // 触发起始点
            var selection = (TextSelection)_dte.ActiveDocument.Selection;
            EditPoint editPoint = selection.AnchorPoint.CreateEditPoint();

            // 选中触发行
            // TODO：匹配多行
            string code = editPoint.GetLines(editPoint.Line, editPoint.Line + 1);

            return code;
        }
    }
}

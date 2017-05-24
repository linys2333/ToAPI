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
}

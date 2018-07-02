using System;
using System.IO;

namespace Bll.Tools
{
    public class GisHelper
    {
        private static readonly object lockObj = new object();

        /// <summary>
        /// 写入txt文件
        /// </summary>
        /// <param name="txtName">文件名称</param>
        /// <param name="txtContent">写入内容</param>
        /// <param name="txtContent">路径</param>
        public static void WriteTxt(string txtName, string txtContent, string txtPath)
        {
            try
            {
                lock (lockObj)
                {
                    //保存路径
                    if (!Directory.Exists(txtPath))
                    {
                        Directory.CreateDirectory(txtPath);
                    }                    

                    txtPath = txtPath + txtName + ".txt";
                    File.WriteAllText(txtPath, txtContent);
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            
        }

        /// <summary>
        /// 读取txt文件
        /// </summary>
        /// <param name="txtPath">路径</param>
        /// <param name="txtName">文件名称</param>
        /// <returns>文件内容</returns>
        public static string ReadTxt(string txtName, string txtPath)
        {
            lock (lockObj)
            {
                txtPath = txtPath + txtName + ".txt";
                string str = File.ReadAllText(txtPath, System.Text.Encoding.Default);

                return str;
            }
        }

        /// <summary>
        /// 校验txt
        /// </summary>
        /// <param name="name"></param>
        /// <param name="txtPath"></param>
        public static void CheckTxt(string name, string txtPath)
        {
            lock (lockObj)
            {
                if (Directory.Exists(txtPath))
                {
                    string pattern = name + "#*.txt";
                    string[] txtFiles = Directory.GetFiles(txtPath, pattern);

                    int len = txtFiles.Length;

                    while (len >= 5)
                    {
                        string _txtName = "";
                        var _date = DateTime.Now.AddDays(+1);
                        foreach (var item in txtFiles)
                        {
                            var tt = item.Split('#');
                            var tDate = DateTime.Parse(tt[1]);

                            if (DateTime.Compare(tDate, _date) < 0)
                            {
                                _txtName = item;
                                _date = tDate;
                            }
                        }

                        File.Delete(_txtName);

                        txtFiles = Directory.GetFiles(txtPath, pattern);
                        len = txtFiles.Length;
                    }
                }
            }
        }

        /// <summary>
        /// 删除txt文件
        /// </summary>
        /// <param name="txtName">名称</param>
        /// <param name="txtPath">路径</param>
        public static void DelTxt(string txtName, string txtPath)
        {
            if (Directory.Exists(txtPath))
            {
                string pattern = txtName + ".txt";
                string[] txtFiles = Directory.GetFiles(txtPath, pattern);

                foreach (var item in txtFiles)
                {
                    File.Delete(item);
                }
            }
        }

    }
}

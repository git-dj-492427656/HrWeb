using Bll.Dal;
using Bll.Tools;
using Newtonsoft.Json.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using Model;
using NPOI.SS.Util;

namespace HRweb.Controllers
{
    public class HrmController : HomeController
    {

        #region //异构数据源

        #region //pcMaker
        //PCMaker 规划预算
        [MyAuthAttribute]
        public ActionResult Hr_PcMaker_ghys()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //PcMaker规划预算数据
        public string GetHr_PcMaker_ghys()
        {
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                JArray array = new JArray();  //返回结果

                //比较开始日期，结束日期
                var cc = DateTime.Compare(start, end);
                if (cc > 0)
                {
                    return array.ToString();  //开始日期大于结束日期,返回空
                }
                
                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_PcMaker_ghys(start, end, mid);
                
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("pcComCode", dt.Rows[i]["pcComCode"].ToString());
                        _json.Add("comName", dt.Rows[i]["pcComName"].ToString());
                        _json.Add("cxNum", dt.Rows[i]["cxNum"].ToString());
                        _json.Add("yieEffic", dt.Rows[i]["yieEffic"].ToString());
                        _json.Add("gjEffic", dt.Rows[i]["gjEffic"].ToString());
                        _json.Add("workDays", dt.Rows[i]["workDays"].ToString());
                        _json.Add("yearly", dt.Rows[i]["yearly"].ToString());
                        _json.Add("monthly", dt.Rows[i]["monthly"].ToString());
                        _json.Add("addTime", dt.Rows[i]["addTime"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //通过id删除
        public string DelHr_PcMaker_ghys()
        {
            try
            {
                int id = int.Parse(Request["id"]);

                var dal = new dalPro();
                int i = dal.DelHr_PcMaker_ghysById(id);

                return i > 0 ? "删除成功！" : "删除失败！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //Excel数据导入
        public void ImportHr_PcMaker_ghys(ISheet sheet)
        {
            try
            {
                //组织单元集合(有权限)
                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string addPer = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var list = new List<Dictionary<string, string>>();

                int rowCount = sheet.LastRowNum + 1; //总行数
                for (int i = 1; i < rowCount; i++)
                {
                    //读到空行终止
                    var _row = sheet.GetRow(i);
                    if (_row == null)
                        break;
                    var _cell = sheet.GetRow(i).GetCell(0);
                    if (_cell == null)
                        break;
                    var _rowCell = sheet.GetRow(i).GetCell(0).StringCellValue;
                    if (string.IsNullOrWhiteSpace(_rowCell))
                        break;

                    string pcComName = sheet.GetRow(i).GetCell(0).ToString();
                    string pcComCode = sheet.GetRow(i).GetCell(1).ToString();
                    int yearly = int.Parse(sheet.GetRow(i).GetCell(2).ToString());
                    int monthly = int.Parse(sheet.GetRow(i).GetCell(3).ToString());
                    double cxNum = double.Parse(sheet.GetRow(i).GetCell(4).ToString());
                    double yieEffic = double.Parse(sheet.GetRow(i).GetCell(5).ToString());
                    double gjEffic = double.Parse(sheet.GetRow(i).GetCell(6).ToString());
                    double workDays = double.Parse(sheet.GetRow(i).GetCell(7).ToString());

                    //int dd = _row.LastCellNum;

                    var conditions = new Dictionary<string, string>();
                    conditions.Add("pcComCode", pcComCode);
                    conditions.Add("pcComName", pcComName);
                    conditions.Add("cxNum", cxNum.ToString());
                    conditions.Add("yieEffic", yieEffic.ToString());
                    conditions.Add("gjEffic", gjEffic.ToString());
                    conditions.Add("workDays", workDays.ToString());
                    conditions.Add("yearly", yearly.ToString());
                    conditions.Add("monthly", monthly.ToString());
                    conditions.Add("mid", midStr);
                    conditions.Add("addPer", addPer);

                    list.Add(conditions);

                }

                if (list != null)
                {
                    var dal = new dalPro();
                    foreach (var item in list)
                    {
                        dal.AddHr_PcMaker_ghys(item);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        
        //导出
        public void ExportHr_PcMaker_ghys()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("PcMaker规划预算", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("PCMaker规划预算"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("PcMaker公司名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("PcMaker公司编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("年 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("月 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产线数");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产效（立方/人/8H）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("构件产效（立方/人/8H）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("工作天数（天）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_PcMaker_ghys(start, end, mid); //数据

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["pcComName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["pcComCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yearly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["monthly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["cxNum"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yieEffic"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["gjEffic"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["workDays"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                    }
                }

                for (int k = 0; k < colCount; k++)
                {
                    //设置列宽
                    if (k < 2)
                    {
                        sheet.AutoSizeColumn(k);  //自适应宽度
                    }
                    else if (k > 4 && k <= 7)
                    {
                        sheet.SetColumnWidth(k, 30 * 256);
                    }
                    else
                    {
                        sheet.SetColumnWidth(k, 15 * 256);
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }


        //PCMaker 项目进展
        [MyAuthAttribute]
        public ActionResult Hr_PcMaker_xmjz()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //PcMaker项目进展
        public string GetHr_PcMaker_xmjz()
        {
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                JArray array = new JArray();  //返回结果

                //比较开始日期，结束日期
                var cc = DateTime.Compare(start, end);
                if (cc > 0)
                {
                    return array.ToString();  //开始日期大于结束日期,返回空
                }

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_PcMaker_xmjz(start, end, mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("pcComCode", dt.Rows[i]["pcComCode"].ToString());
                        _json.Add("comName", dt.Rows[i]["pcComName"].ToString());
                        _json.Add("cxNum", dt.Rows[i]["cxNum"].ToString());
                        _json.Add("proBudget", dt.Rows[i]["proBudget"].ToString());
                        _json.Add("progjBudget", dt.Rows[i]["progjBudget"].ToString());
                        _json.Add("yieEffic", dt.Rows[i]["yieEffic"].ToString());
                        _json.Add("gjEffic", dt.Rows[i]["gjEffic"].ToString());
                        _json.Add("workDays", dt.Rows[i]["workDays"].ToString());
                        _json.Add("yearly", dt.Rows[i]["yearly"].ToString());
                        _json.Add("monthly", dt.Rows[i]["monthly"].ToString());
                        _json.Add("addTime", dt.Rows[i]["addTime"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //通过id删除
        public string DelHr_PcMaker_xmjz()
        {
            try
            {
                int id = int.Parse(Request["id"]);

                var dal = new dalPro();
                int i = dal.DelHr_PcMaker_xmjzById(id);

                return i > 0 ? "删除成功！" : "删除失败！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //Excel数据导入
        public void ImportHr_PcMaker_xmjz(ISheet sheet)
        {
            try
            {
                //组织单元集合(有权限)
                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string addPer = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var list = new List<Dictionary<string, string>>();

                int rowCount = sheet.LastRowNum + 1; //总行数
                for (int i = 1; i < rowCount; i++)
                {
                    //读到空行终止
                    var _row = sheet.GetRow(i);
                    if (_row == null)
                        break;
                    var _cell = sheet.GetRow(i).GetCell(0);
                    if (_cell == null)
                        break;
                    var _rowCell = sheet.GetRow(i).GetCell(0).StringCellValue;
                    if (string.IsNullOrWhiteSpace(_rowCell))
                        break;

                    string pcComName = sheet.GetRow(i).GetCell(0).ToString();
                    string pcComCode = sheet.GetRow(i).GetCell(1).ToString();
                    int yearly = int.Parse(sheet.GetRow(i).GetCell(2).ToString());
                    int monthly = int.Parse(sheet.GetRow(i).GetCell(3).ToString());
                    double cxNum = double.Parse(sheet.GetRow(i).GetCell(4).ToString());
                    double proBudget = double.Parse(sheet.GetRow(i).GetCell(5).ToString());
                    double progjBudget = double.Parse(sheet.GetRow(i).GetCell(6).ToString());
                    double yieEffic = double.Parse(sheet.GetRow(i).GetCell(7).ToString());
                    double gjEffic = double.Parse(sheet.GetRow(i).GetCell(8).ToString());
                    double workDays = double.Parse(sheet.GetRow(i).GetCell(9).ToString());
                    

                    var conditions = new Dictionary<string, string>();
                    conditions.Add("pcComCode", pcComCode);
                    conditions.Add("pcComName", pcComName);
                    conditions.Add("cxNum", cxNum.ToString());
                    conditions.Add("proBudget", proBudget.ToString());
                    conditions.Add("progjBudget", progjBudget.ToString());
                    conditions.Add("yieEffic", yieEffic.ToString());
                    conditions.Add("gjEffic", gjEffic.ToString());
                    conditions.Add("workDays", workDays.ToString());
                    conditions.Add("yearly", yearly.ToString());
                    conditions.Add("monthly", monthly.ToString());
                    conditions.Add("mid", midStr);
                    conditions.Add("addPer", addPer);

                    list.Add(conditions);

                }

                if (list != null)
                {
                    var dal = new dalPro();
                    foreach (var item in list)
                    {
                        dal.AddHr_PcMaker_xmjz(item);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出
        public void ExportHr_PcMaker_xmjz()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("PcMaker项目进展", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("PCMaker项目进展"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("PcMaker公司名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("PcMaker公司编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("年 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("月 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产线数");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产量（万立方）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("构件产量（万立方）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产效（立方/人/8H）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("构件产效（立方/人/8H）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("工作天数（天）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_PcMaker_xmjz(start, end, mid); //数据

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["pcComName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["pcComCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yearly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["monthly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["cxNum"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["proBudget"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["progjBudget"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yieEffic"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["gjEffic"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["workDays"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                    }
                }

                for (int k = 0; k < colCount; k++)
                {
                    //设置列宽
                    if (k < 2)
                    {
                        sheet.AutoSizeColumn(k);  //自适应宽度
                    }
                    else if (k > 4 && k <= 9)
                    {
                        sheet.SetColumnWidth(k, 30 * 256);
                    }
                    else
                    {
                        sheet.SetColumnWidth(k, 15 * 256);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }


        //PCMaker 实际
        [MyAuthAttribute]
        public ActionResult Hr_PcMaker_fact()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //PcMaker实际数据
        public string GetHr_PcMaker_fact()
        {
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                JArray array = new JArray();  //返回结果

                //比较开始日期，结束日期
                var cc = DateTime.Compare(start, end);
                if (cc > 0)
                {
                    return array.ToString();  //开始日期大于结束日期,返回空
                }

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_PcMaker_fact(start, end, mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("pcComCode", dt.Rows[i]["pcComCode"].ToString());
                        _json.Add("comName", dt.Rows[i]["pcComName"].ToString());
                        _json.Add("cxNum", dt.Rows[i]["cxNum"].ToString());
                        _json.Add("yieEffic", dt.Rows[i]["yieEffic"].ToString());
                        _json.Add("gjEffic", dt.Rows[i]["gjEffic"].ToString());
                        _json.Add("yearly", dt.Rows[i]["yearly"].ToString());
                        _json.Add("monthly", dt.Rows[i]["monthly"].ToString());
                        _json.Add("addTime", dt.Rows[i]["addTime"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //通过id删除
        public string DelHr_PcMaker_fact()
        {
            try
            {
                int id = int.Parse(Request["id"]);

                var dal = new dalPro();
                int i = dal.DelHr_PcMaker_factById(id);

                return i > 0 ? "删除成功！" : "删除失败！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //Excel数据导入
        public void ImportHr_PcMaker_fact(ISheet sheet)
        {
            try
            {
                //组织单元集合(有权限)
                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string addPer = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var list = new List<Dictionary<string, string>>();

                int rowCount = sheet.LastRowNum + 1; //总行数
                for (int i = 1; i < rowCount; i++)
                {
                    //读到空行终止
                    var _row = sheet.GetRow(i);
                    if (_row == null)
                        break;
                    var _cell = sheet.GetRow(i).GetCell(0);
                    if (_cell == null)
                        break;
                    var _rowCell = sheet.GetRow(i).GetCell(0).StringCellValue;
                    if (string.IsNullOrWhiteSpace(_rowCell))
                        break;

                    string pcComName = sheet.GetRow(i).GetCell(0).ToString();
                    string pcComCode = sheet.GetRow(i).GetCell(1).ToString();
                    int yearly = int.Parse(sheet.GetRow(i).GetCell(2).ToString());
                    int monthly = int.Parse(sheet.GetRow(i).GetCell(3).ToString());
                    double cxNum = double.Parse(sheet.GetRow(i).GetCell(4).ToString());
                    double yieEffic = double.Parse(sheet.GetRow(i).GetCell(5).ToString());
                    double gjEffic = double.Parse(sheet.GetRow(i).GetCell(6).ToString());
                    
                    var conditions = new Dictionary<string, string>();
                    conditions.Add("pcComCode", pcComCode);
                    conditions.Add("pcComName", pcComName);
                    conditions.Add("cxNum", cxNum.ToString());
                    conditions.Add("yieEffic", yieEffic.ToString());
                    conditions.Add("gjEffic", gjEffic.ToString());
                    conditions.Add("yearly", yearly.ToString());
                    conditions.Add("monthly", monthly.ToString());
                    conditions.Add("mid", midStr);
                    conditions.Add("addPer", addPer);

                    list.Add(conditions);
                }

                if (list != null)
                {
                    var dal = new dalPro();
                    foreach (var item in list)
                    {
                        dal.AddHr_PcMaker_fact(item);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出
        public void ExportHr_PcMaker_fact()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("PcMaker实际", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("PCMaker实际"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("PcMaker公司名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("PcMaker公司编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("年 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("月 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产线数");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产效（立方/人/8H）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("构件产效（立方/人/8H）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_PcMaker_fact(start, end, mid); //数据

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["pcComName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["pcComCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yearly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["monthly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["cxNum"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yieEffic"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["gjEffic"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                    }
                }

                for (int k = 0; k < colCount; k++)
                {
                    //设置列宽
                    if (k < 2)
                    {
                        sheet.AutoSizeColumn(k);  //自适应宽度
                    }
                    else if (k > 4)
                    {
                        sheet.SetColumnWidth(k, 30 * 256);
                    }
                    else
                    {
                        sheet.SetColumnWidth(k, 15 * 256);
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }

        #endregion

        #region //bais

        //Bais 规划预算
        [MyAuthAttribute]
        public ActionResult Hr_Bais_ghys()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];
            
            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //Bais规划预算数据
        public string GetHr_Bais_ghys()
        {
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                JArray array = new JArray();  //返回结果

                //比较开始日期，结束日期
                var cc = DateTime.Compare(start, end);
                if (cc > 0)
                {
                    return array.ToString();  //开始日期大于结束日期,返回空
                }

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Bais_ghys(start, end, mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("baisComCode", dt.Rows[i]["baisComCode"].ToString());
                        _json.Add("baisComName", dt.Rows[i]["baisComName"].ToString());
                        _json.Add("htAmount", dt.Rows[i]["htAmount"].ToString());
                        _json.Add("ysAmount", dt.Rows[i]["ysAmount"].ToString());
                        _json.Add("yield", dt.Rows[i]["yield"].ToString());
                        _json.Add("lrAmount", dt.Rows[i]["lrAmount"].ToString());
                        _json.Add("yearly", dt.Rows[i]["yearly"].ToString());
                        _json.Add("monthly", dt.Rows[i]["monthly"].ToString());
                        _json.Add("addTime", dt.Rows[i]["addTime"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //通过id删除
        public string DelHr_Bais_ghys()
        {
            try
            {
                int id = int.Parse(Request["id"]);

                var dal = new dalPro();
                int i = dal.DelHr_Bais_ghysById(id);

                return i > 0 ? "删除成功！" : "删除失败！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //Excel数据导入
        public void ImportHr_Bais_ghys(ISheet sheet)
        {
            try
            {
                //组织单元集合(有权限)
                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string addPer = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var list = new List<Dictionary<string, string>>();

                int rowCount = sheet.LastRowNum + 1; //总行数
                for (int i = 1; i < rowCount; i++)
                {
                    //读到空行终止
                    var _row = sheet.GetRow(i);
                    if (_row == null)
                        break;
                    var _cell = sheet.GetRow(i).GetCell(0);
                    if (_cell == null)
                        break;
                    var _rowCell = sheet.GetRow(i).GetCell(0).StringCellValue;
                    if (string.IsNullOrWhiteSpace(_rowCell))
                        break;

                    string baisComName = sheet.GetRow(i).GetCell(0).ToString();
                    string baisComCode = sheet.GetRow(i).GetCell(1).ToString();
                    int yearly = int.Parse(sheet.GetRow(i).GetCell(2).ToString());
                    int monthly = int.Parse(sheet.GetRow(i).GetCell(3).ToString());
                    double htAmount = double.Parse(sheet.GetRow(i).GetCell(4).ToString());
                    double ysAmount = double.Parse(sheet.GetRow(i).GetCell(5).ToString());
                    double yield = double.Parse(sheet.GetRow(i).GetCell(6).ToString());
                    double lrAmount = double.Parse(sheet.GetRow(i).GetCell(7).ToString());

                    var conditions = new Dictionary<string, string>();
                    conditions.Add("baisComName", baisComName);
                    conditions.Add("baisComCode", baisComCode);
                    conditions.Add("htAmount", htAmount.ToString());
                    conditions.Add("ysAmount", ysAmount.ToString());
                    conditions.Add("yield", yield.ToString());
                    conditions.Add("lrAmount", lrAmount.ToString());
                    conditions.Add("yearly", yearly.ToString());
                    conditions.Add("monthly", monthly.ToString());
                    conditions.Add("mid", midStr);
                    conditions.Add("addPer", addPer);

                    list.Add(conditions);
                }

                if (list != null)
                {
                    var dal = new dalPro();
                    foreach (var item in list)
                    {
                        dal.AddHr_Bais_ghys(item);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出
        public void ExportHr_Bais_ghys()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("Bais规划预算", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("Bais规划预算"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("Bais公司名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("Bais公司编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("年 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("月 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("合同额（万元）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("营收（万元）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产量（万立方）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("利润（万元）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Bais_ghys(start, end, mid); //数据

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["baisComName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["baisComCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yearly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["monthly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["htAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["ysAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yield"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["lrAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                    }
                }

                
                for (int k = 0; k < colCount; k++)
                {
                    //设置列宽
                    if (k < 2)
                    {
                        sheet.AutoSizeColumn(k); //自适应宽度
                    }
                    else
                    {
                        sheet.SetColumnWidth(k, 20 * 256);
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }


        //Bais 项目进展
        [MyAuthAttribute]
        public ActionResult Hr_Bais_xmjz()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //Bais 项目进展数据
        public string GetHr_Bais_xmjz()
        {
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                JArray array = new JArray();  //返回结果

                //比较开始日期，结束日期
                var cc = DateTime.Compare(start, end);
                if (cc > 0)
                {
                    return array.ToString();  //开始日期大于结束日期,返回空
                }

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Bais_xmjz(start, end, mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("baisComCode", dt.Rows[i]["baisComCode"].ToString());
                        _json.Add("baisComName", dt.Rows[i]["baisComName"].ToString());
                        _json.Add("htAmount", dt.Rows[i]["htAmount"].ToString());
                        _json.Add("ysAmount", dt.Rows[i]["ysAmount"].ToString());
                        _json.Add("yield", dt.Rows[i]["yield"].ToString());
                        _json.Add("lrAmount", dt.Rows[i]["lrAmount"].ToString());
                        _json.Add("yearly", dt.Rows[i]["yearly"].ToString());
                        _json.Add("monthly", dt.Rows[i]["monthly"].ToString());
                        _json.Add("addTime", dt.Rows[i]["addTime"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //通过id删除
        public string DelHr_Bais_xmjz()
        {
            try
            {
                int id = int.Parse(Request["id"]);

                var dal = new dalPro();
                int i = dal.DelHr_Bais_xmjzById(id);

                return i > 0 ? "删除成功！" : "删除失败！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //Excel数据导入
        public void ImportHr_Bais_xmjz(ISheet sheet)
        {
            try
            {
                //组织单元集合(有权限)
                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string addPer = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var list = new List<Dictionary<string, string>>();

                int rowCount = sheet.LastRowNum + 1; //总行数
                for (int i = 1; i < rowCount; i++)
                {
                    //读到空行终止
                    var _row = sheet.GetRow(i);
                    if (_row == null)
                        break;
                    var _cell = sheet.GetRow(i).GetCell(0);
                    if (_cell == null)
                        break;
                    var _rowCell = sheet.GetRow(i).GetCell(0).StringCellValue;
                    if (string.IsNullOrWhiteSpace(_rowCell))
                        break;

                    string baisComName = sheet.GetRow(i).GetCell(0).ToString();
                    string baisComCode = sheet.GetRow(i).GetCell(1).ToString();
                    int yearly = int.Parse(sheet.GetRow(i).GetCell(2).ToString());
                    int monthly = int.Parse(sheet.GetRow(i).GetCell(3).ToString());
                    double htAmount = double.Parse(sheet.GetRow(i).GetCell(4).ToString());
                    double ysAmount = double.Parse(sheet.GetRow(i).GetCell(5).ToString());
                    double yield = double.Parse(sheet.GetRow(i).GetCell(6).ToString());
                    double lrAmount = double.Parse(sheet.GetRow(i).GetCell(7).ToString());

                    var conditions = new Dictionary<string, string>();
                    conditions.Add("baisComName", baisComName);
                    conditions.Add("baisComCode", baisComCode);
                    conditions.Add("htAmount", htAmount.ToString());
                    conditions.Add("ysAmount", ysAmount.ToString());
                    conditions.Add("yield", yield.ToString());
                    conditions.Add("lrAmount", lrAmount.ToString());
                    conditions.Add("yearly", yearly.ToString());
                    conditions.Add("monthly", monthly.ToString());
                    conditions.Add("mid", midStr);
                    conditions.Add("addPer", addPer);

                    list.Add(conditions);
                }

                if (list != null)
                {
                    var dal = new dalPro();
                    foreach (var item in list)
                    {
                        dal.AddHr_Bais_xmjz(item);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出
        public void ExportHr_Bais_xmjz()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("Bais项目进展", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("Bais项目进展"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("Bais公司名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("Bais公司编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("年 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("月 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("合同额（万元）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("营收（万元）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产量（万立方）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("利润（万元）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Bais_xmjz(start, end, mid); //数据

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["baisComName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["baisComCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yearly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["monthly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["htAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["ysAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yield"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["lrAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                    }
                }


                for (int k = 0; k < colCount; k++)
                {
                    //设置列宽
                    if (k < 2)
                    {
                        sheet.AutoSizeColumn(k); //自适应宽度
                    }
                    else
                    {
                        sheet.SetColumnWidth(k, 20 * 256);
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }


        //Bais 实际
        [MyAuthAttribute]
        public ActionResult Hr_Bais_fact()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //Bais 实际数据
        public string GetHr_Bais_fact()
        {
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                JArray array = new JArray();  //返回结果

                //比较开始日期，结束日期
                var cc = DateTime.Compare(start, end);
                if (cc > 0)
                {
                    return array.ToString();  //开始日期大于结束日期,返回空
                }

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Bais_fact(start, end, mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("baisComCode", dt.Rows[i]["baisComCode"].ToString());
                        _json.Add("baisComName", dt.Rows[i]["baisComName"].ToString());
                        _json.Add("htAmount", dt.Rows[i]["htAmount"].ToString());
                        _json.Add("ysAmount", dt.Rows[i]["ysAmount"].ToString());
                        _json.Add("yield", dt.Rows[i]["yield"].ToString());
                        _json.Add("lrAmount", dt.Rows[i]["lrAmount"].ToString());
                        _json.Add("yearly", dt.Rows[i]["yearly"].ToString());
                        _json.Add("monthly", dt.Rows[i]["monthly"].ToString());
                        _json.Add("addTime", dt.Rows[i]["addTime"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //通过id删除
        public string DelHr_Bais_fact()
        {
            try
            {
                int id = int.Parse(Request["id"]);

                var dal = new dalPro();
                int i = dal.DelHr_Bais_factById(id);

                return i > 0 ? "删除成功！" : "删除失败！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //Excel数据导入
        public void ImportHr_Bais_fact(ISheet sheet)
        {
            try
            {
                //组织单元集合(有权限)
                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string addPer = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var list = new List<Dictionary<string, string>>();

                int rowCount = sheet.LastRowNum + 1; //总行数
                for (int i = 1; i < rowCount; i++)
                {
                    //读到空行终止
                    var _row = sheet.GetRow(i);
                    if (_row == null)
                        break;
                    var _cell = sheet.GetRow(i).GetCell(0);
                    if (_cell == null)
                        break;
                    var _rowCell = sheet.GetRow(i).GetCell(0).StringCellValue;
                    if (string.IsNullOrWhiteSpace(_rowCell))
                        break;

                    string baisComName = sheet.GetRow(i).GetCell(0).ToString();
                    string baisComCode = sheet.GetRow(i).GetCell(1).ToString();
                    int yearly = int.Parse(sheet.GetRow(i).GetCell(2).ToString());
                    int monthly = int.Parse(sheet.GetRow(i).GetCell(3).ToString());
                    double htAmount = double.Parse(sheet.GetRow(i).GetCell(4).ToString());
                    double ysAmount = double.Parse(sheet.GetRow(i).GetCell(5).ToString());
                    double yield = double.Parse(sheet.GetRow(i).GetCell(6).ToString());
                    double lrAmount = double.Parse(sheet.GetRow(i).GetCell(7).ToString());

                    var conditions = new Dictionary<string, string>();
                    conditions.Add("baisComName", baisComName);
                    conditions.Add("baisComCode", baisComCode);
                    conditions.Add("htAmount", htAmount.ToString());
                    conditions.Add("ysAmount", ysAmount.ToString());
                    conditions.Add("yield", yield.ToString());
                    conditions.Add("lrAmount", lrAmount.ToString());
                    conditions.Add("yearly", yearly.ToString());
                    conditions.Add("monthly", monthly.ToString());
                    conditions.Add("mid", midStr);
                    conditions.Add("addPer", addPer);

                    list.Add(conditions);
                }

                if (list != null)
                {
                    var dal = new dalPro();
                    foreach (var item in list)
                    {
                        dal.AddHr_Bais_fact(item);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出
        public void ExportHr_Bais_fact()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("Bais实际", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("Bais实际"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("Bais公司名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("Bais公司编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("年 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("月 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("合同额（万元）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("营收（万元）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产量（万立方）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("利润（万元）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Bais_fact(start, end, mid); //数据

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["baisComName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["baisComCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yearly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["monthly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["htAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["ysAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yield"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["lrAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                    }
                }


                for (int k = 0; k < colCount; k++)
                {
                    //设置列宽
                    if (k < 2)
                    {
                        sheet.AutoSizeColumn(k); //自适应宽度
                    }
                    else
                    {
                        sheet.SetColumnWidth(k, 20 * 256);
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }


        //Bais 人工成本-收入
        [MyAuthAttribute]
        public ActionResult Hr_Bais_rgsr()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //Bais 人工收入数据
        public string GetHr_Bais_rgsr()
        {
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                JArray array = new JArray();  //返回结果

                //比较开始日期，结束日期
                var cc = DateTime.Compare(start, end);
                if (cc > 0)
                {
                    return array.ToString();  //开始日期大于结束日期,返回空
                }

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Bais_rgsr(start, end, mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("baisComCode", dt.Rows[i]["baisComCode"].ToString());
                        _json.Add("baisComName", dt.Rows[i]["baisComName"].ToString());
                        _json.Add("costType", dt.Rows[i]["costType"].ToString());
                        _json.Add("proportion", dt.Rows[i]["proportion"].ToString() + "%");
                        _json.Add("planBudget", dt.Rows[i]["planBudget"].ToString());
                        _json.Add("adjustBudget", dt.Rows[i]["adjustBudget"].ToString());
                        _json.Add("proBudget", dt.Rows[i]["proBudget"].ToString());
                        _json.Add("quotaLabor", dt.Rows[i]["quotaLabor"].ToString());
                        _json.Add("yearly", dt.Rows[i]["yearly"].ToString());
                        _json.Add("monthly", dt.Rows[i]["monthly"].ToString());
                        _json.Add("addTime", dt.Rows[i]["addTime"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //通过id删除
        public string DelHr_Bais_rgsr()
        {
            try
            {
                int id = int.Parse(Request["id"]);

                var dal = new dalPro();
                int i = dal.DelHr_Bais_rgsrById(id);

                return i > 0 ? "删除成功！" : "删除失败！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //Excel数据导入
        public void ImportHr_Bais_rgsr(ISheet sheet)
        {
            try
            {
                //组织单元集合(有权限)
                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string addPer = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var list = new List<Dictionary<string, string>>();

                int rowCount = sheet.LastRowNum + 1; //总行数
                for (int i = 1; i < rowCount; i++)
                {
                    //读到空行终止
                    var _row = sheet.GetRow(i);
                    if (_row == null)
                        break;
                    var _cell = sheet.GetRow(i).GetCell(0);
                    if (_cell == null)
                        break;
                    var _rowCell = sheet.GetRow(i).GetCell(0).StringCellValue;
                    if (string.IsNullOrWhiteSpace(_rowCell))
                        break;

                    string baisComName = sheet.GetRow(i).GetCell(0).ToString();
                    string baisComCode = sheet.GetRow(i).GetCell(1).ToString();
                    int yearly = int.Parse(sheet.GetRow(i).GetCell(2).ToString());
                    int monthly = int.Parse(sheet.GetRow(i).GetCell(3).ToString());
                    string costType = sheet.GetRow(i).GetCell(4).ToString();

                    //去掉百分比符号
                    string _pro = sheet.GetRow(i).GetCell(5).ToString();
                    double proportion = double.Parse(_pro.Substring(0, _pro.Length - 1));

                    double planBudget = double.Parse(sheet.GetRow(i).GetCell(6).ToString());
                    double adjustBudget = double.Parse(sheet.GetRow(i).GetCell(7).ToString());
                    double proBudget = double.Parse(sheet.GetRow(i).GetCell(8).ToString());
                    double quotaLabor = double.Parse(sheet.GetRow(i).GetCell(9).ToString());
                    

                    var conditions = new Dictionary<string, string>();
                    conditions.Add("baisComName", baisComName);
                    conditions.Add("baisComCode", baisComCode);
                    conditions.Add("costType", costType);
                    conditions.Add("planBudget", planBudget.ToString());
                    conditions.Add("adjustBudget", adjustBudget.ToString());
                    conditions.Add("proBudget", proBudget.ToString());
                    conditions.Add("quotaLabor", quotaLabor.ToString());
                    conditions.Add("proportion", proportion.ToString());
                    conditions.Add("yearly", yearly.ToString());
                    conditions.Add("monthly", monthly.ToString());
                    conditions.Add("mid", midStr);
                    conditions.Add("addPer", addPer);

                    list.Add(conditions);

                }

                if (list != null)
                {
                    var dal = new dalPro();
                    foreach (var item in list)
                    {
                        dal.AddHr_Bais_rgsr(item);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出
        public void ExportHr_Bais_rgsr()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("Bais人工成本-收入", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("Bais人工成本-收入"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("Bais公司名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("Bais公司编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("年 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("月 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("费用类别");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("比 例(%)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("规划预算(万元)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("调整预算(万元)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("项目进展(万元)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("实  际(万元)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Bais_rgsr(start, end, mid); //数据

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["baisComName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["baisComCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yearly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["monthly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["costType"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["proportion"].ToString() + "%");
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["planBudget"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["adjustBudget"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["proBudget"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["quotaLabor"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                    }
                }


                for (int k = 0; k < colCount; k++)
                {
                    //设置列宽
                    if (k < 2)
                    {
                        sheet.AutoSizeColumn(k);
                    }
                    else
                    {
                        sheet.SetColumnWidth(k, 20 * 256);
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }


        //Bais 人工成本-支出
        [MyAuthAttribute]
        public ActionResult Hr_Bais_rgzc()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //Bais 人工支出数据
        public string GetHr_Bais_rgzc()
        {
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                JArray array = new JArray();  //返回结果

                //比较开始日期，结束日期
                var cc = DateTime.Compare(start, end);
                if (cc > 0)
                {
                    return array.ToString();  //开始日期大于结束日期,返回空
                }

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Bais_rgzc(start, end, mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("baisComCode", dt.Rows[i]["baisComCode"].ToString());
                        _json.Add("baisComName", dt.Rows[i]["baisComName"].ToString());
                        _json.Add("costType", dt.Rows[i]["costType"].ToString());
                        _json.Add("quotaLabor", dt.Rows[i]["quotaLabor"].ToString());
                        _json.Add("yearly", dt.Rows[i]["yearly"].ToString());
                        _json.Add("monthly", dt.Rows[i]["monthly"].ToString());
                        _json.Add("addTime", dt.Rows[i]["addTime"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //通过id删除
        public string DelHr_Bais_rgzc()
        {
            try
            {
                int id = int.Parse(Request["id"]);

                var dal = new dalPro();
                int i = dal.DelHr_Bais_rgzcById(id);

                return i > 0 ? "删除成功！" : "删除失败！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //Excel数据导入
        public void ImportHr_Bais_rgzc(ISheet sheet)
        {
            try
            {
                //组织单元集合(有权限)
                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string addPer = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var list = new List<Dictionary<string, string>>();

                int rowCount = sheet.LastRowNum + 1; //总行数
                for (int i = 1; i < rowCount; i++)
                {
                    //读到空行终止
                    var _row = sheet.GetRow(i);
                    if (_row == null)
                        break;
                    var _cell = sheet.GetRow(i).GetCell(0);
                    if (_cell == null)
                        break;
                    var _rowCell = sheet.GetRow(i).GetCell(0).StringCellValue;
                    if (string.IsNullOrWhiteSpace(_rowCell))
                        break;

                    string baisComName = sheet.GetRow(i).GetCell(0).ToString();
                    string baisComCode = sheet.GetRow(i).GetCell(1).ToString();
                    int yearly = int.Parse(sheet.GetRow(i).GetCell(2).ToString());
                    int monthly = int.Parse(sheet.GetRow(i).GetCell(3).ToString());
                    string costType = sheet.GetRow(i).GetCell(4).ToString();                    
                    double quotaLabor = double.Parse(sheet.GetRow(i).GetCell(5).ToString());


                    var conditions = new Dictionary<string, string>();
                    conditions.Add("baisComName", baisComName);
                    conditions.Add("baisComCode", baisComCode);
                    conditions.Add("costType", costType);
                    conditions.Add("quotaLabor", quotaLabor.ToString());
                    conditions.Add("yearly", yearly.ToString());
                    conditions.Add("monthly", monthly.ToString());
                    conditions.Add("mid", midStr);
                    conditions.Add("addPer", addPer);

                    list.Add(conditions);

                }

                if (list != null)
                {
                    var dal = new dalPro();
                    foreach (var item in list)
                    {
                        dal.AddHr_Bais_rgzc(item);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出
        public void ExportHr_Bais_rgzc()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("Bais人工成本-支出", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("Bais人工成本-支出"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("Bais公司名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("Bais公司编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("年 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("月 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("费用类别");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("实际(万元)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Bais_rgzc(start, end, mid); //数据

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["baisComName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["baisComCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yearly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["monthly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["costType"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["quotaLabor"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                    }
                }


                for (int k = 0; k < colCount; k++)
                {
                    //设置列宽
                    if (k < 2 || k > 4)
                    {
                        sheet.AutoSizeColumn(k);
                    }
                    else
                    {
                        sheet.SetColumnWidth(k, 20 * 256);
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }

        #endregion

        #region //crm,bhr

        //Crm 规划预算
        [MyAuthAttribute]
        public ActionResult Hr_Crm_ghys()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //Crm 规划预算数据
        public string GetHr_Crm_ghys()
        {
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                JArray array = new JArray();  //返回结果

                //比较开始日期，结束日期
                var cc = DateTime.Compare(start, end);
                if (cc > 0)
                {
                    return array.ToString();  //开始日期大于结束日期,返回空
                }

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Crm_ghys(start, end, mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("crmComCode", dt.Rows[i]["crmComCode"].ToString());
                        _json.Add("crmComName", dt.Rows[i]["crmComName"].ToString());
                        _json.Add("crmDeptCode", dt.Rows[i]["crmDeptCode"].ToString());
                        _json.Add("crmDeptName", dt.Rows[i]["crmDeptName"].ToString());
                        _json.Add("htAmount", dt.Rows[i]["htAmount"].ToString());
                        _json.Add("yearly", dt.Rows[i]["yearly"].ToString());
                        _json.Add("monthly", dt.Rows[i]["monthly"].ToString());
                        _json.Add("addTime", dt.Rows[i]["addTime"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //通过id删除
        public string DelHr_Crm_ghys()
        {
            try
            {
                int id = int.Parse(Request["id"]);

                var dal = new dalPro();
                int i = dal.DelHr_Crm_ghysById(id);

                return i > 0 ? "删除成功！" : "删除失败！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //Excel数据导入
        public void ImportHr_Crm_ghys(ISheet sheet)
        {
            try
            {
                //组织单元集合(有权限)
                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string addPer = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var list = new List<Dictionary<string, string>>();

                int rowCount = sheet.LastRowNum + 1; //总行数
                for (int i = 1; i < rowCount; i++)
                {
                    //读到空行终止
                    var _row = sheet.GetRow(i);
                    if (_row == null)
                        break;
                    var _cell = sheet.GetRow(i).GetCell(0);
                    if (_cell == null)
                        break;
                    var _rowCell = sheet.GetRow(i).GetCell(0).StringCellValue;
                    if (string.IsNullOrWhiteSpace(_rowCell))
                        break;

                    string crmComName = sheet.GetRow(i).GetCell(0).ToString();
                    string crmComCode = sheet.GetRow(i).GetCell(1).ToString();
                    string crmDeptName = sheet.GetRow(i).GetCell(2).ToString();
                    string crmDeptCode = sheet.GetRow(i).GetCell(3).ToString();
                    int yearly = int.Parse(sheet.GetRow(i).GetCell(4).ToString());
                    int monthly = int.Parse(sheet.GetRow(i).GetCell(5).ToString());
                    double htAmount = double.Parse(sheet.GetRow(i).GetCell(6).ToString());


                    var conditions = new Dictionary<string, string>();
                    conditions.Add("crmComName", crmComName);
                    conditions.Add("crmComCode", crmComCode);
                    conditions.Add("crmDeptName", crmDeptName);
                    conditions.Add("crmDeptCode", crmDeptCode);
                    conditions.Add("htAmount", htAmount.ToString());
                    conditions.Add("yearly", yearly.ToString());
                    conditions.Add("monthly", monthly.ToString());
                    conditions.Add("mid", midStr);
                    conditions.Add("addPer", addPer);

                    list.Add(conditions);

                }

                if (list != null)
                {
                    var dal = new dalPro();
                    foreach (var item in list)
                    {
                        dal.AddHr_Crm_ghys(item);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出
        public void ExportHr_Crm_ghys()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("Crm规划预算", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("Crm规划预算"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("Crm公司名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("Crm公司编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("Crm部门名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("Crm部门编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("年 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("月 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("合同额(万元)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Crm_ghys(start, end, mid); //数据

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["crmComName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["crmComCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["crmDeptName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["crmDeptCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yearly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["monthly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["htAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                    }
                }


                for (int k = 0; k < colCount; k++)
                {
                    //设置列宽
                    if (k < 4 && k > 5)
                    {
                        sheet.AutoSizeColumn(k);
                    }
                    else
                    {
                        sheet.SetColumnWidth(k, 20 * 256);
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }

        
        //Crm 实际
        [MyAuthAttribute]
        public ActionResult Hr_Crm_fact()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //Crm 实际数据
        public string GetHr_Crm_fact()
        {
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                JArray array = new JArray();  //返回结果

                //比较开始日期，结束日期
                var cc = DateTime.Compare(start, end);
                if (cc > 0)
                {
                    return array.ToString();  //开始日期大于结束日期,返回空
                }

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Crm_fact(start, end, mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("crmComCode", dt.Rows[i]["crmComCode"].ToString());
                        _json.Add("crmComName", dt.Rows[i]["crmComName"].ToString());
                        _json.Add("crmDeptCode", dt.Rows[i]["crmDeptCode"].ToString());
                        _json.Add("crmDeptName", dt.Rows[i]["crmDeptName"].ToString());
                        _json.Add("syAmount", dt.Rows[i]["syAmount"].ToString());
                        _json.Add("yearly", dt.Rows[i]["yearly"].ToString());
                        _json.Add("monthly", dt.Rows[i]["monthly"].ToString());
                        _json.Add("addTime", dt.Rows[i]["addTime"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //通过id删除
        public string DelHr_Crm_fact()
        {
            try
            {
                int id = int.Parse(Request["id"]);

                var dal = new dalPro();
                int i = dal.DelHr_Crm_factById(id);

                return i > 0 ? "删除成功！" : "删除失败！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //Excel数据导入
        public void ImportHr_Crm_fact(ISheet sheet)
        {
            try
            {
                //组织单元集合(有权限)
                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string addPer = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var list = new List<Dictionary<string, string>>();

                int rowCount = sheet.LastRowNum + 1; //总行数
                for (int i = 1; i < rowCount; i++)
                {
                    //读到空行终止
                    var _row = sheet.GetRow(i);
                    if (_row == null)
                        break;
                    var _cell = sheet.GetRow(i).GetCell(0);
                    if (_cell == null)
                        break;
                    var _rowCell = sheet.GetRow(i).GetCell(0).StringCellValue;
                    if (string.IsNullOrWhiteSpace(_rowCell))
                        break;

                    string crmComName = sheet.GetRow(i).GetCell(0).ToString();
                    string crmComCode = sheet.GetRow(i).GetCell(1).ToString();
                    string crmDeptName = sheet.GetRow(i).GetCell(2).ToString();
                    string crmDeptCode = sheet.GetRow(i).GetCell(3).ToString();
                    int yearly = int.Parse(sheet.GetRow(i).GetCell(4).ToString());
                    int monthly = int.Parse(sheet.GetRow(i).GetCell(5).ToString());
                    double syAmount = double.Parse(sheet.GetRow(i).GetCell(6).ToString());


                    var conditions = new Dictionary<string, string>();
                    conditions.Add("crmComName", crmComName);
                    conditions.Add("crmComCode", crmComCode);
                    conditions.Add("crmDeptName", crmDeptName);
                    conditions.Add("crmDeptCode", crmDeptCode);
                    conditions.Add("syAmount", syAmount.ToString());
                    conditions.Add("yearly", yearly.ToString());
                    conditions.Add("monthly", monthly.ToString());
                    conditions.Add("mid", midStr);
                    conditions.Add("addPer", addPer);

                    list.Add(conditions);

                }

                if (list != null)
                {
                    var dal = new dalPro();
                    foreach (var item in list)
                    {
                        dal.AddHr_Crm_fact(item);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出
        public void ExportHr_Crm_fact()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("Crm实际", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("Crm实际"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("Crm公司名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("Crm公司编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("Crm部门名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("Crm部门编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("年 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("月 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("实际合同额(万元)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Crm_fact(start, end, mid); //数据

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["crmComName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["crmComCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["crmDeptName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["crmDeptCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yearly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["monthly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["syAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                    }
                }


                for (int k = 0; k < colCount; k++)
                {
                    //设置列宽
                    if (k < 4)
                    {
                        sheet.AutoSizeColumn(k);
                    }
                    else
                    {
                        sheet.SetColumnWidth(k, 20 * 256);
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }
        


        //Crm 调整预算
        [MyAuthAttribute]
        public ActionResult Hr_Crm_tzys()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //Crm 调整预算数据
        public string GetHr_Crm_tzys()
        {
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                JArray array = new JArray();  //返回结果

                //比较开始日期，结束日期
                var cc = DateTime.Compare(start, end);
                if (cc > 0)
                {
                    return array.ToString();  //开始日期大于结束日期,返回空
                }

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Crm_tzys(start, end, mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("crmComCode", dt.Rows[i]["crmComCode"].ToString());
                        _json.Add("crmComName", dt.Rows[i]["crmComName"].ToString());
                        _json.Add("crmDeptCode", dt.Rows[i]["crmDeptCode"].ToString());
                        _json.Add("crmDeptName", dt.Rows[i]["crmDeptName"].ToString());
                        _json.Add("htAmount", dt.Rows[i]["htAmount"].ToString());
                        _json.Add("syAmount", dt.Rows[i]["syAmount"].ToString());
                        _json.Add("ghsyAmount", dt.Rows[i]["ghsyAmount"].ToString());
                        _json.Add("tzsyAmount", dt.Rows[i]["tzsyAmount"].ToString());
                        _json.Add("yearly", dt.Rows[i]["yearly"].ToString());
                        _json.Add("monthly", dt.Rows[i]["monthly"].ToString());
                        _json.Add("addTime", dt.Rows[i]["addTime"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //通过id删除
        public string DelHr_Crm_tzys()
        {
            try
            {
                string ids = Request["ids"];

                var idList = ids.Split(',');

                var dal = new dalPro();
                foreach (var cc in idList)
                {
                    int _id = int.Parse(cc);
                    dal.DelHr_Crm_tzysById(_id);
                }

                return "删除成功！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //导出
        public void ExportHr_Crm_tzys()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("Crm调整预算", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("Crm调整预算"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("Crm公司名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("Crm公司编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("Crm部门名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("Crm部门编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("年 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("月 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("调整预算合同额(万元)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("实际合同额(万元)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("规划预算剩余合同额(万元)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("调整预算剩余合同额(万元)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Crm_tzys(start, end, mid); //数据

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["crmComName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["crmComCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["crmDeptName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["crmDeptCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yearly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["monthly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["htAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["syAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["ghsyAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["tzsyAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                    }
                }


                for (int k = 0; k < colCount; k++)
                {
                    //设置列宽
                    if (k < 4)
                    {
                        sheet.AutoSizeColumn(k);
                    }
                    else if (k > 5)
                    {
                        sheet.SetColumnWidth(k, 30 * 256);
                    }
                    else
                    {
                        sheet.SetColumnWidth(k, 15 * 256);
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }

        ////通过id修改 调整预算合同额
        public string SaveHr_Crm_tzys()
        {
            try
            {
                int id = int.Parse(Request["id"]);
                double htAmount = double.Parse(Request["htAmount"]);
                double syAmount = double.Parse(Request["syAmount"]);

                double tzsyAmount = htAmount - syAmount;
                
                var cookie = GetCookie();
                string name = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var dal = new dalPro();
                int i = dal.EditHr_Crm_tzysById(id, htAmount, tzsyAmount, name);

                return i > 0 ? "修改成功！" : "修改失败！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //// 同步数据
        public string SynHr_Crm_tzys()
        {
            try
            {
                int yearly = int.Parse(Request["yearly"]);
                int monthly = int.Parse(Request["monthly"]);

                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string name = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var dal = new dalPro();
                string result = "";

                //规划预算 数据源
                var dt = dal.GetHr_Crm_tzysFrom(yearly, monthly, midStr);
                if (dt.Rows.Count == 0)
                {
                    result = "该月度无Crm数据！";
                }
                else
                {
                    double htAmount = double.Parse(dt.Rows[0]["htAmount"].ToString()); //调整预算合同额 默认=规划预算合同额
                    double syAmount = double.Parse(dt.Rows[0]["syAmount"].ToString());
                    double _amount = htAmount - syAmount; //剩余合同额

                    var condition = new Dictionary<string, string>();
                    condition.Add("yearly", yearly.ToString());
                    condition.Add("monthly", monthly.ToString());
                    condition.Add("crmComCode", dt.Rows[0]["crmComCode"].ToString()); //公司 eas编码
                    condition.Add("crmComName", dt.Rows[0]["crmComName"].ToString());
                    condition.Add("crmDeptCode", dt.Rows[0]["crmDeptCode"].ToString());
                    condition.Add("crmDeptName", dt.Rows[0]["crmDeptName"].ToString());
                    condition.Add("htAmount", htAmount.ToString());
                    condition.Add("syAmount", syAmount.ToString());
                    condition.Add("ghsyAmount", _amount.ToString());
                    condition.Add("tzsyAmount", _amount.ToString());
                    condition.Add("mid", midStr);
                    condition.Add("addPer", name);

                    int i = dal.AddHr_Crm_tzys(condition);
                    if (i > 0)
                    {
                        result = "同步成功！";
                    }
                    else
                    {
                        result = "同步失败！";
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }



        //BHR 实际
        [MyAuthAttribute]
        public ActionResult Hr_Bhr_fact()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //Bhr 实际数据
        public string GetHr_Bhr_fact()
        {
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                JArray array = new JArray();  //返回结果

                //比较开始日期，结束日期
                var cc = DateTime.Compare(start, end);
                if (cc > 0)
                {
                    return array.ToString();  //开始日期大于结束日期,返回空
                }

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Bhr_fact(start, end, mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("easComCode", dt.Rows[i]["easComCode"].ToString());
                        _json.Add("easComName", dt.Rows[i]["easComName"].ToString());
                        _json.Add("easDeptCode", dt.Rows[i]["easDeptCode"].ToString());
                        _json.Add("easDeptName", dt.Rows[i]["easDeptName"].ToString());
                        _json.Add("easPostCode", dt.Rows[i]["easPostCode"].ToString());
                        _json.Add("easPostName", dt.Rows[i]["easPostName"].ToString());
                        _json.Add("staffId", dt.Rows[i]["staffId"].ToString());
                        _json.Add("postLevel", dt.Rows[i]["postLevel"].ToString());
                        _json.Add("postType", dt.Rows[i]["postType"].ToString());
                        _json.Add("wage", dt.Rows[i]["wage"].ToString());
                        _json.Add("workDays", dt.Rows[i]["workDays"].ToString());
                        _json.Add("yearly", dt.Rows[i]["yearly"].ToString());
                        _json.Add("monthly", dt.Rows[i]["monthly"].ToString());
                        _json.Add("addTime", dt.Rows[i]["addTime"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //通过id删除
        public string DelHr_Bhr_fact()
        {
            try
            {
                int id = int.Parse(Request["id"]);

                var dal = new dalPro();
                int i = dal.DelHr_Bhr_factById(id);

                return i > 0 ? "删除成功！" : "删除失败！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //Excel数据导入
        public void ImportHr_Bhr_fact(ISheet sheet)
        {
            try
            {
                //组织单元集合(有权限)
                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string addPer = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var list = new List<Dictionary<string, string>>();

                int rowCount = sheet.LastRowNum + 1; //总行数
                for (int i = 1; i < rowCount; i++)
                {
                    //读到空行终止
                    var _row = sheet.GetRow(i);
                    if (_row == null)
                        break;
                    var _cell = sheet.GetRow(i).GetCell(0);
                    if (_cell == null)
                        break;
                    var _rowCell = sheet.GetRow(i).GetCell(0).StringCellValue;
                    if (string.IsNullOrWhiteSpace(_rowCell))
                        break;

                    string easComName = sheet.GetRow(i).GetCell(0).ToString();
                    string easComCode = sheet.GetRow(i).GetCell(1).ToString();
                    string easDeptName = sheet.GetRow(i).GetCell(2).ToString();
                    string easDeptCode = sheet.GetRow(i).GetCell(3).ToString();
                    int yearly = int.Parse(sheet.GetRow(i).GetCell(4).ToString());
                    int monthly = int.Parse(sheet.GetRow(i).GetCell(5).ToString());
                    string easPostName = sheet.GetRow(i).GetCell(6).ToString();
                    string easPostCode = sheet.GetRow(i).GetCell(7).ToString();
                    string staffId = sheet.GetRow(i).GetCell(8).ToString();
                    string postLevel = sheet.GetRow(i).GetCell(9).ToString();
                    string postType = sheet.GetRow(i).GetCell(10).ToString();
                    double wage = double.Parse(sheet.GetRow(i).GetCell(11).ToString());
                    double workDays = double.Parse(sheet.GetRow(i).GetCell(12).ToString());


                    var conditions = new Dictionary<string, string>();
                    conditions.Add("easComCode", easComCode);
                    conditions.Add("easComName", easComName);
                    conditions.Add("easDeptCode", easDeptCode);
                    conditions.Add("easDeptName", easDeptName);
                    conditions.Add("easPostCode", easPostCode);
                    conditions.Add("easPostName", easPostName);
                    conditions.Add("staffId", staffId);
                    conditions.Add("postLevel", postLevel);
                    conditions.Add("postType", postType);
                    conditions.Add("wage", wage.ToString());
                    conditions.Add("workDays", workDays.ToString());
                    conditions.Add("yearly", yearly.ToString());
                    conditions.Add("monthly", monthly.ToString());
                    conditions.Add("mid", midStr);
                    conditions.Add("addPer", addPer);

                    list.Add(conditions);

                }

                if (list != null)
                {
                    var dal = new dalPro();
                    foreach (var item in list)
                    {
                        dal.AddHr_Bhr_fact(item);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出
        public void ExportHr_Bhr_fact()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("Bhr实际", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("Bhr实际"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("Bhr公司名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("Bhr公司编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("Bhr部门名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("Bhr部门编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("年 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("月 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("Bhr岗位名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("Bhr岗位编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("员工工号");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("职 级");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("岗位类别");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("定额工资");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("工作天数");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Bhr_fact(start, end, mid); //数据

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["easComName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["easComCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["easDeptName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["easDeptCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yearly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["monthly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["easPostName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["easPostCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["staffId"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["postLevel"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["postType"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["wage"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["workDays"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                    }
                }


                for (int k = 0; k < colCount; k++)
                {
                    //设置列宽
                    if (k < 4)
                    {
                        sheet.AutoSizeColumn(k);
                    }
                    else
                    {
                        sheet.SetColumnWidth(k, 20 * 256);
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }

        #endregion

        #endregion
        

        #region //岗位配备规则
        //岗位配备规则-营收
        [MyAuthAttribute]
        public ActionResult Hr_Rule_ys()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //岗位配备营收 数据
        public string GetHr_Rule_ys()
        {
            try
            {
                int yearly = int.Parse(Request["yearly"]);

                //返回结果
                var json = new JObject();
                json.Add("isExp", "0");
                json.Add("dataStr", "");

                //组织单元
                var cookie = GetCookie();
                string midStr = cookie["comId"];

                List<Hr_Rule_yshead> list1 = null; // 营收规则
                List<Hr_Rule_ysgw> list2 = null;   // 岗位配备人数
                List<Hr_Rule_ysfd> list3 = null;   // 岗位配备浮动人数
                
                #region //数据读取
                var dal = new dalPro();

                var dt1 = dal.GetHr_Rule_yshead(yearly, midStr);// 营收规则
                var dt2 = dal.GetHr_Rule_ysgw(yearly, midStr);  // 岗位配备人数
                var dt3 = dal.GetHr_Rule_ysfd(yearly, midStr);  // 岗位配备浮动人数
                if (dt1.Rows.Count > 0)
                {
                    list1 = new List<Hr_Rule_yshead>();
                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        var obj1 = new Hr_Rule_yshead();
                        obj1.id = dt1.Rows[i]["id"].ToString();
                        obj1.ysAmount = double.Parse(dt1.Rows[i]["ysAmount"].ToString());
                        obj1.yield = double.Parse(dt1.Rows[i]["yield"].ToString());
                        obj1.effic = double.Parse(dt1.Rows[i]["effic"].ToString());
                        obj1.dhbEffic = double.Parse(dt1.Rows[i]["dhbEffic"].ToString());
                        obj1.nqEffic = double.Parse(dt1.Rows[i]["nqEffic"].ToString());
                        obj1.wqEffic = double.Parse(dt1.Rows[i]["wqEffic"].ToString());
                        obj1.smzEffic = double.Parse(dt1.Rows[i]["smzEffic"].ToString());
                        obj1.workDays = double.Parse(dt1.Rows[i]["workDays"].ToString());
                        obj1.yearly = int.Parse(dt1.Rows[i]["yearly"].ToString());

                        list1.Add(obj1);
                    }
                }

                if (dt2.Rows.Count > 0)
                {
                    list2 = new List<Hr_Rule_ysgw>();
                    for (int i = 0; i < dt2.Rows.Count; i++)
                    {
                        var obj2 = new Hr_Rule_ysgw();
                        obj2.id = dt2.Rows[i]["id"].ToString();
                        obj2.deptCode = dt2.Rows[i]["deptCode"].ToString();
                        obj2.deptName = dt2.Rows[i]["deptName"].ToString();
                        obj2.postCode = dt2.Rows[i]["postCode"].ToString();
                        obj2.postName = dt2.Rows[i]["postName"].ToString();
                        obj2.postLevel = dt2.Rows[i]["postLevel"].ToString();
                        obj2.costType = dt2.Rows[i]["costType"].ToString();
                        obj2.quotaWage = double.Parse(dt2.Rows[i]["quotaWage"].ToString());
                        obj2.coreNum = int.Parse(dt2.Rows[i]["coreNum"].ToString());
                        obj2.boneNum = int.Parse(dt2.Rows[i]["boneNum"].ToString());
                        obj2.yearly = int.Parse(dt2.Rows[i]["yearly"].ToString());

                        list2.Add(obj2);
                    }
                }

                if (dt3.Rows.Count > 0)
                {
                    list3 = new List<Hr_Rule_ysfd>();
                    for (int i = 0; i < dt3.Rows.Count; i++)
                    {
                        var obj3 = new Hr_Rule_ysfd();
                        obj3.rid = dt3.Rows[i]["rid"].ToString();
                        obj3.gwid = dt3.Rows[i]["gwid"].ToString();
                        obj3.floatNum = int.Parse(dt3.Rows[i]["floatNum"].ToString());
                        obj3.yearly = int.Parse(dt3.Rows[i]["yearly"].ToString());

                        list3.Add(obj3);
                    }
                }
                #endregion

                if (null != list1 && null != list2 && null != list3)
                {
                    string result = "<table class='tblStyle2'>"; 

                    #region //营收规则
                    string tr1 = "<tr><td colspan='10'>月度营收（万元）</td>";
                    string tr2 = "<tr><td colspan='10'>月度产量（万立方）</td>";
                    string tr3 = "<tr><td colspan='10'>总产效（立方/人/天）-除销售人员外</td>";
                    string tr4 = "<tr><td colspan='10'>叠合板产效</td>";
                    string tr5 = "<tr><td colspan='10'>剪力墙、内墙产效</td>";
                    string tr6 = "<tr><td colspan='10'>外挂墙板产效</td>";
                    string tr7 = "<tr><td colspan='10'>三明治剪力墙产效</td>";
                    string tr8 = "<tr><td colspan='10'>月度可生产天数（天）</td>";

                    for (int j = 0; j < list1.Count(); j++)
                    {
                        tr1 += "<td>" + list1[j].ysAmount + "</td>";
                        tr2 += "<td>" + list1[j].yield + "</td>";
                        tr3 += "<td>" + list1[j].effic + "</td>";
                        tr4 += "<td>" + list1[j].dhbEffic + "</td>";
                        tr5 += "<td>" + list1[j].nqEffic + "</td>";
                        tr6 += "<td>" + list1[j].wqEffic + "</td>";
                        tr7 += "<td>" + list1[j].smzEffic + "</td>";
                        tr8 += "<td>" + list1[j].workDays + "</td>";
                    }
                    tr1 += "</tr>";
                    tr2 += "</tr>";
                    tr3 += "</tr>";
                    tr4 += "</tr>";
                    tr5 += "</tr>";
                    tr6 += "</tr>";
                    tr7 += "</tr>";
                    tr8 += "</tr>";

                    string tr9 = "<tr><td colspan='" + (10 + list1.Count()) + "'></td></tr>"; //空行

                    result += tr1 + tr2 + tr3 + tr4 + tr5 + tr6 + tr7 + tr8 + tr9;
                    #endregion

                    #region //配备岗位数据
                    //岗位配备人数-表头
                    string tr10 = "<tr>";
                    tr10 += "<td>序号</td>";
                    tr10 += "<td>部门</td>";
                    tr10 += "<td>部门编码</td>";
                    tr10 += "<td>职位名称</td>";
                    tr10 += "<td>职位编码</td>";
                    tr10 += "<td>职级</td>";
                    tr10 += "<td>费用类别</td>";
                    tr10 += "<td>定额工资</td>";
                    tr10 += "<td>核心</td>";
                    tr10 += "<td>骨干</td>";
                    tr10 += "<td colspan='" + list1.Count() + "'>浮动-按月度营收（万元）</td>";
                    tr10 += "</tr>";
                    result += tr10;

                    for (int m = 0; m < list2.Count(); m++)
                    {
                        string tr = "<tr>";
                        tr += "<td>" + (m + 1) + "</td>";
                        tr += "<td>" + list2[m].deptName + "</td>";
                        tr += "<td>" + list2[m].deptCode + "</td>";
                        tr += "<td>" + list2[m].postName + "</td>";
                        tr += "<td>" + list2[m].postCode + "</td>";
                        tr += "<td>" + list2[m].postLevel + "</td>";
                        tr += "<td>" + list2[m].costType + "</td>";
                        tr += "<td>" + list2[m].quotaWage + "</td>";
                        tr += "<td>" + list2[m].coreNum + "</td>";
                        tr += "<td>" + list2[m].boneNum + "</td>";

                        for (int n = 0; n < list1.Count(); n++)
                        {
                            tr += "<td>" + list3.Where(p => p.rid == list1[n].id && p.gwid == list2[m].id).First().floatNum + "</td>";
                        }
                        tr += "</tr>";

                        result += tr;
                    }
                    #endregion

                    result += "</table>";

                    json["isExp"] = "1";
                    json["dataStr"] = result;
                }
                
                return json.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //Excel数据导入
        public void ImportHr_Rule_ys(ISheet sheet)
        {
            try
            {
                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string comName = HttpUtility.UrlDecode(cookie["comName"], System.Text.Encoding.GetEncoding("GB2312"));
                string addPer = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));
                
                var list1 = new List<Hr_Rule_yshead>(); // 营收规则
                var list2 = new List<Hr_Rule_ysgw>();   // 岗位配备人数
                var list3 = new List<Hr_Rule_ysfd>();   // 岗位配备浮动人数

                int yearly = DateTime.Now.Year; // 当前年度

                #region //遍历excel数据
                int rowCount = sheet.LastRowNum + 1; //总行数

                var row1 = sheet.GetRow(1);
                var row2 = sheet.GetRow(2);
                var row3 = sheet.GetRow(3);
                var row4 = sheet.GetRow(4);
                var row5 = sheet.GetRow(5);
                var row6 = sheet.GetRow(6);
                var row7 = sheet.GetRow(7);
                var row8 = sheet.GetRow(8);

                int colCount = row1.LastCellNum;  //总列数
                for (int i = 10; i < colCount; i++)
                {
                    var obj1 = new Hr_Rule_yshead();

                    obj1.id = Guid.NewGuid().ToString();     //规则id
                    obj1.ysAmount = double.Parse(row1.GetCell(i).ToString());   //月度营收（万元）
                    obj1.yield = double.Parse(row2.GetCell(i).ToString());      //月度产量（万立方）
                    obj1.effic = double.Parse(row3.GetCell(i).ToString());      //总产效（立方/人/天）
                    obj1.dhbEffic = double.Parse(row4.GetCell(i).ToString());   //叠合板产效
                    obj1.nqEffic = double.Parse(row5.GetCell(i).ToString());    //剪力墙、内墙产效
                    obj1.wqEffic = double.Parse(row6.GetCell(i).ToString());    //外挂墙产效
                    obj1.smzEffic = double.Parse(row7.GetCell(i).ToString());   //三明治剪力墙产效
                    obj1.workDays = double.Parse(row8.GetCell(i).ToString());   //月度可生产天数（天）

                    obj1.yearly = yearly;
                    obj1.baisComCode = midStr;

                    list1.Add(obj1);
                }


                string deptName = "";
                string deptCode = "";
                for (int i = 11; i < rowCount; i++)
                {
                    var row = sheet.GetRow(i);

                    //空行终止
                    if (row == null)
                        break;
                    var cell = row.GetCell(3);
                    if (cell == null)
                        break;
                    string cellValue = cell.StringCellValue;
                    if (string.IsNullOrWhiteSpace(cellValue))
                        break;


                    // 岗位配备人数
                    var obj2 = new Hr_Rule_ysgw();

                    string gwid = Guid.NewGuid().ToString(); //id
                    obj2.id = gwid;

                    string _deptStr = row.GetCell(1).ToString();
                    if (_deptStr != "")
                    {
                        deptName = _deptStr;
                        deptCode = row.GetCell(2).ToString();
                    }

                    obj2.deptName = deptName;   //部门名称
                    obj2.deptCode = deptCode;   //部门编码

                    obj2.postName = row.GetCell(3).ToString();    //职位名称
                    obj2.postCode = row.GetCell(4).ToString();    //职位编码
                    obj2.postLevel = row.GetCell(5).ToString();   //职级
                    obj2.costType = row.GetCell(6).ToString();    //费用类别

                    obj2.quotaWage = double.Parse(row.GetCell(7).ToString());//定额工资
                    obj2.coreNum = int.Parse(row.GetCell(8).ToString());    //核心人数
                    obj2.boneNum = int.Parse(row.GetCell(9).ToString());    //骨干人数

                    obj2.yearly = yearly;
                    obj2.baisComCode = midStr;

                    list2.Add(obj2);

                    // 岗位配备 浮动人数
                    for (int j = 10; j < colCount; j++)
                    {
                        //读到空单元格 终止
                        var _cell = row.GetCell(j);
                        if (_cell == null)
                            break;

                        var obj3 = new Hr_Rule_ysfd();

                        obj3.rid = list1[j - 10].id;
                        obj3.gwid = gwid;
                        obj3.floatNum = int.Parse(row.GetCell(j).ToString());

                        obj3.yearly = yearly;
                        obj3.baisComCode = midStr;

                        list3.Add(obj3);
                    }
                }
                #endregion

                //保存数据
                var dal = new dalPro();
                dal.AddHr_Rule_ys(list1, list2, list3, addPer, int.Parse(midStr));
                
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出
        public void ExportHr_Rule_ys()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("岗位配备规则（营收）", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("岗位配备规则（营收）"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion                      

            #region //表体
            try
            {
                int yearly = int.Parse(Request["yearly"]);
                
                //组织单元
                var cookie = GetCookie();
                string midStr = cookie["comId"];

                List<Hr_Rule_yshead> list1 = null; // 营收规则
                List<Hr_Rule_ysgw> list2 = null;   // 岗位配备人数
                List<Hr_Rule_ysfd> list3 = null;   // 岗位配备浮动人数
                
                #region //数据读取
                var dal = new dalPro();

                var dt1 = dal.GetHr_Rule_yshead(yearly, midStr);// 营收规则
                var dt2 = dal.GetHr_Rule_ysgw(yearly, midStr);  // 岗位配备人数
                var dt3 = dal.GetHr_Rule_ysfd(yearly, midStr);  // 岗位配备浮动人数
                if (dt1.Rows.Count > 0)
                {
                    list1 = new List<Hr_Rule_yshead>();
                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        var obj1 = new Hr_Rule_yshead();
                        obj1.id = dt1.Rows[i]["id"].ToString();
                        obj1.ysAmount = double.Parse(dt1.Rows[i]["ysAmount"].ToString());
                        obj1.yield = double.Parse(dt1.Rows[i]["yield"].ToString());
                        obj1.effic = double.Parse(dt1.Rows[i]["effic"].ToString());
                        obj1.dhbEffic = double.Parse(dt1.Rows[i]["dhbEffic"].ToString());
                        obj1.nqEffic = double.Parse(dt1.Rows[i]["nqEffic"].ToString());
                        obj1.wqEffic = double.Parse(dt1.Rows[i]["wqEffic"].ToString());
                        obj1.smzEffic = double.Parse(dt1.Rows[i]["smzEffic"].ToString());
                        obj1.workDays = double.Parse(dt1.Rows[i]["workDays"].ToString());
                        obj1.yearly = int.Parse(dt1.Rows[i]["yearly"].ToString());

                        list1.Add(obj1);
                    }
                }

                if (dt2.Rows.Count > 0)
                {
                    list2 = new List<Hr_Rule_ysgw>();
                    for (int i = 0; i < dt2.Rows.Count; i++)
                    {
                        var obj2 = new Hr_Rule_ysgw();
                        obj2.id = dt2.Rows[i]["id"].ToString();
                        obj2.deptCode = dt2.Rows[i]["deptCode"].ToString();
                        obj2.deptName = dt2.Rows[i]["deptName"].ToString();
                        obj2.postCode = dt2.Rows[i]["postCode"].ToString();
                        obj2.postName = dt2.Rows[i]["postName"].ToString();
                        obj2.postLevel = dt2.Rows[i]["postLevel"].ToString();
                        obj2.costType = dt2.Rows[i]["costType"].ToString();
                        obj2.quotaWage = double.Parse(dt2.Rows[i]["quotaWage"].ToString());
                        obj2.coreNum = int.Parse(dt2.Rows[i]["coreNum"].ToString());
                        obj2.boneNum = int.Parse(dt2.Rows[i]["boneNum"].ToString());
                        obj2.yearly = int.Parse(dt2.Rows[i]["yearly"].ToString());

                        list2.Add(obj2);
                    }
                }

                if (dt3.Rows.Count > 0)
                {
                    list3 = new List<Hr_Rule_ysfd>();
                    for (int i = 0; i < dt3.Rows.Count; i++)
                    {
                        var obj3 = new Hr_Rule_ysfd();
                        obj3.rid = dt3.Rows[i]["rid"].ToString();
                        obj3.gwid = dt3.Rows[i]["gwid"].ToString();
                        obj3.floatNum = int.Parse(dt3.Rows[i]["floatNum"].ToString());
                        obj3.yearly = int.Parse(dt3.Rows[i]["yearly"].ToString());

                        list3.Add(obj3);
                    }
                }
                #endregion

                if (null != list1 && null != list2 && null != list3)
                {
                    int colCount = 10 + list1.Count();

                    #region //营收规则
                    var row0 = sheet.CreateRow(0);
                    row0.Height = 20 * 35;
                    row0.CreateCell(0).SetCellValue("岗位配备规则（营收）");
                    row0.GetCell(0).CellStyle = style;
                    
                    var row1 = sheet.CreateRow(1);
                    row1.Height = 20 * 25;
                    row1.CreateCell(0).SetCellValue("月度营收（万元）");
                    row1.GetCell(0).CellStyle = style1;
                    var row2 = sheet.CreateRow(2);
                    row2.Height = 20 * 25;
                    row2.CreateCell(0).SetCellValue("月度产量（万立方）");
                    row2.GetCell(0).CellStyle = style1;
                    var row3 = sheet.CreateRow(3);
                    row3.Height = 20 * 25;
                    row3.CreateCell(0).SetCellValue("总产效（立方/人/天）-除销售人员外");
                    row3.GetCell(0).CellStyle = style1;
                    var row4 = sheet.CreateRow(4);
                    row4.Height = 20 * 25;
                    row4.CreateCell(0).SetCellValue("叠合板产效");
                    row4.GetCell(0).CellStyle = style1;
                    var row5 = sheet.CreateRow(5);
                    row5.Height = 20 * 25;
                    row5.CreateCell(0).SetCellValue("剪力墙、内墙产效");
                    row5.GetCell(0).CellStyle = style1;
                    var row6 = sheet.CreateRow(6);
                    row6.Height = 20 * 25;
                    row6.CreateCell(0).SetCellValue("外挂墙板产效");
                    row6.GetCell(0).CellStyle = style1;
                    var row7 = sheet.CreateRow(7);
                    row7.Height = 20 * 25;
                    row7.CreateCell(0).SetCellValue("三明治剪力墙产效");
                    row7.GetCell(0).CellStyle = style1;
                    var row8 = sheet.CreateRow(8);
                    row8.Height = 20 * 25;
                    row8.CreateCell(0).SetCellValue("月度可生产天数（天）");
                    row8.GetCell(0).CellStyle = style1;
                    
                    //空行
                    var row9 = sheet.CreateRow(9);
                    row9.Height = 20 * 25;
                    row9.CreateCell(0).SetCellValue("");
                    row9.GetCell(0).CellStyle = style1;

                    for (int i = 1; i < 10; i++)
                    {
                        row0.CreateCell(i).SetCellValue("");
                        row0.GetCell(i).CellStyle = style;
                        row1.CreateCell(i).SetCellValue("");
                        row1.GetCell(i).CellStyle = style;
                        row2.CreateCell(i).SetCellValue("");
                        row2.GetCell(i).CellStyle = style;
                        row3.CreateCell(i).SetCellValue("");
                        row3.GetCell(i).CellStyle = style;
                        row4.CreateCell(i).SetCellValue("");
                        row4.GetCell(i).CellStyle = style;
                        row5.CreateCell(i).SetCellValue("");
                        row5.GetCell(i).CellStyle = style;
                        row6.CreateCell(i).SetCellValue("");
                        row6.GetCell(i).CellStyle = style;
                        row7.CreateCell(i).SetCellValue("");
                        row7.GetCell(i).CellStyle = style;
                        row8.CreateCell(i).SetCellValue("");
                        row8.GetCell(i).CellStyle = style;
                        row9.CreateCell(i).SetCellValue("");
                        row9.GetCell(i).CellStyle = style1;
                    }

                    //岗位配备-表头
                    var row10 = sheet.CreateRow(10);
                    row10.Height = 20 * 25;
                    row10.CreateCell(0).SetCellValue("序号");
                    row10.GetCell(0).CellStyle = style;
                    row10.CreateCell(1).SetCellValue("部门");
                    row10.GetCell(1).CellStyle = style;
                    row10.CreateCell(2).SetCellValue("部门编码");
                    row10.GetCell(2).CellStyle = style;
                    row10.CreateCell(3).SetCellValue("职位名称");
                    row10.GetCell(3).CellStyle = style;
                    row10.CreateCell(4).SetCellValue("职位编码");
                    row10.GetCell(4).CellStyle = style;
                    row10.CreateCell(5).SetCellValue("职级");
                    row10.GetCell(5).CellStyle = style;
                    row10.CreateCell(6).SetCellValue("费用类别");
                    row10.GetCell(6).CellStyle = style;
                    row10.CreateCell(7).SetCellValue("定额工资");
                    row10.GetCell(7).CellStyle = style;
                    row10.CreateCell(8).SetCellValue("核心");
                    row10.GetCell(8).CellStyle = style;
                    row10.CreateCell(9).SetCellValue("骨干");
                    row10.GetCell(9).CellStyle = style;
                    //row10.CreateCell(10).SetCellValue("");
                    //row10.GetCell(10).CellStyle = style;

                    for (int i = 0; i < list1.Count(); i++)
                    {
                        row0.CreateCell(i + 10).SetCellValue("");
                        row0.GetCell(i + 10).CellStyle = style;
                        row1.CreateCell(i + 10).SetCellValue(list1[i].ysAmount);
                        row1.GetCell(i + 10).CellStyle = style1;
                        row2.CreateCell(i + 10).SetCellValue(list1[i].yield);
                        row2.GetCell(i + 10).CellStyle = style1;
                        row3.CreateCell(i + 10).SetCellValue(list1[i].effic);
                        row3.GetCell(i + 10).CellStyle = style1;
                        row4.CreateCell(i + 10).SetCellValue(list1[i].dhbEffic);
                        row4.GetCell(i + 10).CellStyle = style1;
                        row5.CreateCell(i + 10).SetCellValue(list1[i].nqEffic);
                        row5.GetCell(i + 10).CellStyle = style1;
                        row6.CreateCell(i + 10).SetCellValue(list1[i].wqEffic);
                        row6.GetCell(i + 10).CellStyle = style1;
                        row7.CreateCell(i + 10).SetCellValue(list1[i].smzEffic);
                        row7.GetCell(i + 10).CellStyle = style1;
                        row8.CreateCell(i + 10).SetCellValue(list1[i].workDays);
                        row8.GetCell(i + 10).CellStyle = style1;
                        row9.CreateCell(i + 10).SetCellValue("");
                        row9.GetCell(i + 10).CellStyle = style1;
                        row10.CreateCell(i + 10).SetCellValue("");
                        row10.GetCell(i + 10).CellStyle = style;
                    }

                    //合并单元格
                    sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, colCount - 1));//起始行，结束行，起始列，结束列
                    sheet.AddMergedRegion(new CellRangeAddress(1, 1, 0, 9));
                    sheet.AddMergedRegion(new CellRangeAddress(2, 2, 0, 9));
                    sheet.AddMergedRegion(new CellRangeAddress(3, 3, 0, 9));
                    sheet.AddMergedRegion(new CellRangeAddress(4, 4, 0, 9));
                    sheet.AddMergedRegion(new CellRangeAddress(5, 5, 0, 9));
                    sheet.AddMergedRegion(new CellRangeAddress(6, 6, 0, 9));
                    sheet.AddMergedRegion(new CellRangeAddress(7, 7, 0, 9));
                    sheet.AddMergedRegion(new CellRangeAddress(8, 8, 0, 9));
                    sheet.AddMergedRegion(new CellRangeAddress(9, 9, 0, colCount - 1));

                    row10.GetCell(10).SetCellValue("浮动-按月度营收（万元）");
                    sheet.AddMergedRegion(new CellRangeAddress(10, 10, 10, colCount - 1));

                    #endregion

                    #region //配备岗位数据
                    for (int i = 0; i < list2.Count(); i++)
                    {
                        int m = i + 11;
                        var row = sheet.CreateRow(m);
                        row.Height = 20 * 25;

                        row.CreateCell(0).SetCellValue(i + 1);
                        row.GetCell(0).CellStyle = style1;
                        row.CreateCell(1).SetCellValue(list2[i].deptName);
                        row.GetCell(1).CellStyle = style1;
                        row.CreateCell(2).SetCellValue(list2[i].deptCode);
                        row.GetCell(2).CellStyle = style1;
                        row.CreateCell(3).SetCellValue(list2[i].postName);
                        row.GetCell(3).CellStyle = style1;
                        row.CreateCell(4).SetCellValue(list2[i].postCode);
                        row.GetCell(4).CellStyle = style1;
                        row.CreateCell(5).SetCellValue(list2[i].postLevel);
                        row.GetCell(5).CellStyle = style1;
                        row.CreateCell(6).SetCellValue(list2[i].costType);
                        row.GetCell(6).CellStyle = style1;
                        row.CreateCell(7).SetCellValue(list2[i].quotaWage);
                        row.GetCell(7).CellStyle = style1;
                        row.CreateCell(8).SetCellValue(list2[i].coreNum);
                        row.GetCell(8).CellStyle = style1;
                        row.CreateCell(9).SetCellValue(list2[i].boneNum);
                        row.GetCell(9).CellStyle = style1;

                        for (int j = 0; j < list1.Count(); j++)
                        {
                            row.CreateCell(10 + j).SetCellValue(list3.Where(p => p.rid == list1[j].id && p.gwid == list2[i].id).First().floatNum);
                            row.GetCell(10 + j).CellStyle = style1;
                        }
                    }
                    #endregion

                    
                    for (int k = 0; k < colCount; k++)
                    {
                        if (k > 0 && k < 5)
                        {
                            sheet.AutoSizeColumn(k);  //自适应宽度
                        }
                        else
                        {
                            sheet.SetColumnWidth(k, 10 * 256); //设置列宽
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }


        //岗位配备规则-合同
        [MyAuthAttribute]
        public ActionResult Hr_Rule_ht()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //岗位配备 合同规则 数据
        public string GetHr_Rule_ht()
        {
            try
            {
                int yearly = int.Parse(Request["yearly"]);
                string htType = Request["htType"];

                //返回结果
                var json = new JObject();
                json.Add("isExp", "0");
                json.Add("dataStr", "");

                //组织单元
                var cookie = GetCookie();
                string midStr = cookie["comId"];

                List<Hr_Rule_hthead> list1 = null; // 合同规则
                List<Hr_Rule_htgw> list2 = null;   // 岗位配备
                List<Hr_Rule_htfd> list3 = null;   // 岗位配备人数

                #region //数据读取
                var dal = new dalPro();

                var dt1 = dal.GetHr_Rule_hthead(yearly, htType, midStr);// 合同规则
                var dt2 = dal.GetHr_Rule_htgw(yearly, htType, midStr);  // 岗位配备
                var dt3 = dal.GetHr_Rule_htfd(yearly, htType, midStr);  // 岗位配备人数
                if (dt1.Rows.Count > 0)
                {
                    list1 = new List<Hr_Rule_hthead>();
                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        var obj1 = new Hr_Rule_hthead();
                        obj1.id = dt1.Rows[i]["id"].ToString();
                        obj1.htType = dt1.Rows[i]["htType"].ToString();
                        obj1.htTitle = dt1.Rows[i]["htTitle"].ToString();
                        obj1.htAmount = double.Parse(dt1.Rows[i]["htAmount"].ToString());
                        obj1.yearly = int.Parse(dt1.Rows[i]["yearly"].ToString());

                        list1.Add(obj1);
                    }
                }

                if (dt2.Rows.Count > 0)
                {
                    list2 = new List<Hr_Rule_htgw>();
                    for (int i = 0; i < dt2.Rows.Count; i++)
                    {
                        var obj2 = new Hr_Rule_htgw();
                        obj2.id = dt2.Rows[i]["id"].ToString();
                        obj2.deptCode = dt2.Rows[i]["deptCode"].ToString();
                        obj2.deptName = dt2.Rows[i]["deptName"].ToString();
                        obj2.postCode = dt2.Rows[i]["postCode"].ToString();
                        obj2.postName = dt2.Rows[i]["postName"].ToString();
                        obj2.postLevel = dt2.Rows[i]["postLevel"].ToString();
                        obj2.costType = dt2.Rows[i]["costType"].ToString();
                        obj2.quotaWage = double.Parse(dt2.Rows[i]["quotaWage"].ToString());
                        obj2.coreNum = int.Parse(dt2.Rows[i]["coreNum"].ToString());
                        obj2.boneNum = int.Parse(dt2.Rows[i]["boneNum"].ToString());
                        obj2.yearly = int.Parse(dt2.Rows[i]["yearly"].ToString());

                        list2.Add(obj2);
                    }
                }

                if (dt3.Rows.Count > 0)
                {
                    list3 = new List<Hr_Rule_htfd>();
                    for (int i = 0; i < dt3.Rows.Count; i++)
                    {
                        var obj3 = new Hr_Rule_htfd();
                        obj3.rid = dt3.Rows[i]["rid"].ToString();
                        obj3.gwid = dt3.Rows[i]["gwid"].ToString();
                        obj3.floatNum = int.Parse(dt3.Rows[i]["floatNum"].ToString());
                        obj3.yearly = int.Parse(dt3.Rows[i]["yearly"].ToString());

                        list3.Add(obj3);
                    }
                }
                #endregion

                if (null != list1 && null != list2 && null != list3)
                {
                    string result = "<table class='tblStyle2'>";

                    #region //合同规则
                    string tr1 = "<tr>";
                    tr1 += "<td rowspan='2'>序号</td>";
                    tr1 += "<td rowspan='2'>部门</td>";
                    tr1 += "<td rowspan='2'>部门编码</td>";
                    tr1 += "<td rowspan='2'>职位名称</td>";
                    tr1 += "<td rowspan='2'>职位编码</td>";
                    tr1 += "<td rowspan='2'>职级</td>";
                    tr1 += "<td rowspan='2'>费用类别</td>";
                    tr1 += "<td rowspan='2'>定额工资</td>";
                    tr1 += "<td rowspan='2'>核心</td>";
                    tr1 += "<td rowspan='2'>骨干</td>";

                    string tr2 = "<tr>";

                    for (int j = 0; j < list1.Count(); j++)
                    {
                        tr1 += "<td>" + list1[j].htTitle + "</td>";
                        tr2 += "<td>" + list1[j].htAmount + "</td>";
                    }
                    tr1 += "</tr>";
                    tr2 += "</tr>";

                    result += tr1 + tr2;
                    #endregion

                    #region //配备岗位数据

                    for (int m = 0; m < list2.Count(); m++)
                    {
                        string coreStr = "";
                        string boneStr = "是";
                        if (1 == list2[m].coreNum)
                        {
                            coreStr = "是";
                            boneStr = "";
                        }

                        string tr = "<tr>";
                        tr += "<td>" + (m + 1) + "</td>";
                        tr += "<td>" + list2[m].deptName + "</td>";
                        tr += "<td>" + list2[m].deptCode + "</td>";
                        tr += "<td>" + list2[m].postName + "</td>";
                        tr += "<td>" + list2[m].postCode + "</td>";
                        tr += "<td>" + list2[m].postLevel + "</td>";
                        tr += "<td>" + list2[m].costType + "</td>";
                        tr += "<td>" + list2[m].quotaWage + "</td>";
                        tr += "<td>" + coreStr + "</td>";
                        tr += "<td>" + boneStr + "</td>";

                        for (int n = 0; n < list1.Count(); n++)
                        {
                            tr += "<td>" + list3.Where(p => p.rid == list1[n].id && p.gwid == list2[m].id).First().floatNum + "</td>";
                        }
                        tr += "</tr>";

                        result += tr;
                    }
                    #endregion

                    result += "</table>";

                    json["isExp"] = "1";
                    json["dataStr"] = result;
                }

                return json.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //Excel数据导入
        public void ImportHr_Rule_ht(ISheet sheet, string htType)
        {
            try
            {
                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string comName = HttpUtility.UrlDecode(cookie["comName"], System.Text.Encoding.GetEncoding("GB2312"));
                string addPer = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var list1 = new List<Hr_Rule_hthead>(); // 合同规则
                var list2 = new List<Hr_Rule_htgw>();   // 岗位配备
                var list3 = new List<Hr_Rule_htfd>();   // 岗位配备数

                int yearly = DateTime.Now.Year; // 当前年度

                #region //遍历excel数据
                int rowCount = sheet.LastRowNum + 1; //总行数

                var row1 = sheet.GetRow(1);
                var row2 = sheet.GetRow(2);

                int colCount = row1.LastCellNum;  //总列数
                for (int i = 10; i < colCount; i++)
                {
                    var obj1 = new Hr_Rule_hthead();

                    obj1.id = Guid.NewGuid().ToString();     //规则id
                    obj1.htTitle = row1.GetCell(i).ToString();   //合同配置名称
                    obj1.htAmount = double.Parse(row2.GetCell(i).ToString());   //合同金额（万元）

                    obj1.htType = htType;
                    obj1.yearly = yearly;
                    obj1.baisComCode = midStr;

                    list1.Add(obj1);
                }


                string deptName = "";
                string deptCode = "";
                for (int i = 3; i < rowCount; i++)
                {
                    var row = sheet.GetRow(i);

                    //空行终止
                    if (row == null)
                        break;
                    var cell = row.GetCell(3);
                    if (cell == null)
                        break;
                    string cellValue = cell.StringCellValue;
                    if (string.IsNullOrWhiteSpace(cellValue))
                        break;


                    // 岗位配备
                    var obj2 = new Hr_Rule_htgw();

                    string gwid = Guid.NewGuid().ToString(); //id
                    obj2.id = gwid;

                    string _deptStr = row.GetCell(1).ToString();
                    if (_deptStr != "")
                    {
                        deptName = _deptStr;
                        deptCode = row.GetCell(2).ToString();
                    }

                    obj2.htType = htType;
                    obj2.deptName = deptName;   //部门名称
                    obj2.deptCode = deptCode;   //部门编码

                    obj2.postName = row.GetCell(3).ToString();    //职位名称
                    obj2.postCode = row.GetCell(4).ToString();    //职位编码
                    obj2.postLevel = row.GetCell(5).ToString();   //职级
                    obj2.costType = row.GetCell(6).ToString();    //费用类别

                    obj2.quotaWage = double.Parse(row.GetCell(7).ToString());//定额工资

                    string coreStr = row.GetCell(8).ToString();
                    if ("是" == coreStr)
                    {
                        obj2.coreNum = 1;    //核心
                        obj2.boneNum = 0;    //骨干
                    }
                    else
                    {
                        obj2.coreNum = 0;    //核心
                        obj2.boneNum = 1;    //骨干
                    }

                    obj2.yearly = yearly;
                    obj2.baisComCode = midStr;

                    list2.Add(obj2);

                    // 岗位配备 人数
                    for (int j = 10; j < colCount; j++)
                    {
                        //读到空单元格 终止
                        var _cell = row.GetCell(j);
                        if (_cell == null)
                            break;

                        var obj3 = new Hr_Rule_htfd();

                        obj3.rid = list1[j - 10].id;
                        obj3.gwid = gwid;
                        obj3.floatNum = int.Parse(row.GetCell(j).ToString());

                        obj3.htType = htType;
                        obj3.yearly = yearly;
                        obj3.baisComCode = midStr;

                        list3.Add(obj3);
                    }
                }
                #endregion

                //保存数据
                var dal = new dalPro();
                dal.AddHr_Rule_ht(list1, list2, list3, addPer, int.Parse(midStr));

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出
        public void ExportHr_Rule_ht()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("岗位配备规则（合同）", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("岗位配备规则（合同）"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion                      

            #region //表体
            try
            {
                int yearly = int.Parse(Request["yearly"]);
                string htType = Request["htType"];

                //组织单元
                var cookie = GetCookie();
                string midStr = cookie["comId"];

                List<Hr_Rule_hthead> list1 = null; // 营收规则
                List<Hr_Rule_htgw> list2 = null;   // 岗位配备人数
                List<Hr_Rule_htfd> list3 = null;   // 岗位配备浮动人数

                #region //数据读取
                var dal = new dalPro();

                var dt1 = dal.GetHr_Rule_hthead(yearly, htType, midStr);// 合同规则
                var dt2 = dal.GetHr_Rule_htgw(yearly, htType, midStr);  // 岗位配备
                var dt3 = dal.GetHr_Rule_htfd(yearly, htType, midStr);  // 岗位配备人数
                if (dt1.Rows.Count > 0)
                {
                    list1 = new List<Hr_Rule_hthead>();
                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        var obj1 = new Hr_Rule_hthead();
                        obj1.id = dt1.Rows[i]["id"].ToString();
                        obj1.htType = dt1.Rows[i]["htType"].ToString();
                        obj1.htTitle = dt1.Rows[i]["htTitle"].ToString();
                        obj1.htAmount = double.Parse(dt1.Rows[i]["htAmount"].ToString());
                        obj1.yearly = int.Parse(dt1.Rows[i]["yearly"].ToString());

                        list1.Add(obj1);
                    }
                }

                if (dt2.Rows.Count > 0)
                {
                    list2 = new List<Hr_Rule_htgw>();
                    for (int i = 0; i < dt2.Rows.Count; i++)
                    {
                        var obj2 = new Hr_Rule_htgw();
                        obj2.id = dt2.Rows[i]["id"].ToString();
                        obj2.deptCode = dt2.Rows[i]["deptCode"].ToString();
                        obj2.deptName = dt2.Rows[i]["deptName"].ToString();
                        obj2.postCode = dt2.Rows[i]["postCode"].ToString();
                        obj2.postName = dt2.Rows[i]["postName"].ToString();
                        obj2.postLevel = dt2.Rows[i]["postLevel"].ToString();
                        obj2.costType = dt2.Rows[i]["costType"].ToString();
                        obj2.quotaWage = double.Parse(dt2.Rows[i]["quotaWage"].ToString());
                        obj2.coreNum = int.Parse(dt2.Rows[i]["coreNum"].ToString());
                        obj2.boneNum = int.Parse(dt2.Rows[i]["boneNum"].ToString());
                        obj2.yearly = int.Parse(dt2.Rows[i]["yearly"].ToString());

                        list2.Add(obj2);
                    }
                }

                if (dt3.Rows.Count > 0)
                {
                    list3 = new List<Hr_Rule_htfd>();
                    for (int i = 0; i < dt3.Rows.Count; i++)
                    {
                        var obj3 = new Hr_Rule_htfd();
                        obj3.rid = dt3.Rows[i]["rid"].ToString();
                        obj3.gwid = dt3.Rows[i]["gwid"].ToString();
                        obj3.floatNum = int.Parse(dt3.Rows[i]["floatNum"].ToString());
                        obj3.yearly = int.Parse(dt3.Rows[i]["yearly"].ToString());

                        list3.Add(obj3);
                    }
                }
                #endregion

                if (null != list1 && null != list2 && null != list3)
                {
                    int colCount = 10 + list1.Count();

                    #region //营收规则
                    var row0 = sheet.CreateRow(0);
                    row0.Height = 20 * 35;
                    row0.CreateCell(0).SetCellValue("岗位配备规则（合同）—" + htType);
                    row0.GetCell(0).CellStyle = style;

                    var row1 = sheet.CreateRow(1);
                    row1.Height = 20 * 25;
                    row1.CreateCell(0).SetCellValue("序号");
                    row1.GetCell(0).CellStyle = style1;
                    row1.CreateCell(1).SetCellValue("部门");
                    row1.GetCell(1).CellStyle = style1;
                    row1.CreateCell(2).SetCellValue("部门编码");
                    row1.GetCell(2).CellStyle = style1;
                    row1.CreateCell(3).SetCellValue("职位名称");
                    row1.GetCell(3).CellStyle = style1;
                    row1.CreateCell(4).SetCellValue("职位编码");
                    row1.GetCell(4).CellStyle = style1;
                    row1.CreateCell(5).SetCellValue("职级");
                    row1.GetCell(5).CellStyle = style1;
                    row1.CreateCell(6).SetCellValue("费用类别");
                    row1.GetCell(6).CellStyle = style1;
                    row1.CreateCell(7).SetCellValue("定额工资");
                    row1.GetCell(7).CellStyle = style1;
                    row1.CreateCell(8).SetCellValue("核心");
                    row1.GetCell(8).CellStyle = style1;
                    row1.CreateCell(9).SetCellValue("骨干");
                    row1.GetCell(9).CellStyle = style1;

                    var row2 = sheet.CreateRow(2);

                    for (int i = 1; i < 10; i++)
                    {
                        row0.CreateCell(i).SetCellValue("");
                        row0.GetCell(i).CellStyle = style;
                        row2.CreateCell(i).SetCellValue("");
                        row2.GetCell(i).CellStyle = style1;
                    }                    

                    for (int i = 0; i < list1.Count(); i++)
                    {
                        row0.CreateCell(i + 10).SetCellValue("");
                        row0.GetCell(i + 10).CellStyle = style;
                        row1.CreateCell(i + 10).SetCellValue(list1[i].htTitle);
                        row1.GetCell(i + 10).CellStyle = style1;
                        row2.CreateCell(i + 10).SetCellValue(list1[i].htAmount);
                        row2.GetCell(i + 10).CellStyle = style1;
                    }

                    //合并单元格
                    sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, colCount - 1));//起始行，结束行，起始列，结束列
                    sheet.AddMergedRegion(new CellRangeAddress(1, 2, 0, 0));
                    sheet.AddMergedRegion(new CellRangeAddress(1, 2, 1, 1));
                    sheet.AddMergedRegion(new CellRangeAddress(1, 2, 2, 2));
                    sheet.AddMergedRegion(new CellRangeAddress(1, 2, 3, 3));
                    sheet.AddMergedRegion(new CellRangeAddress(1, 2, 4, 4));
                    sheet.AddMergedRegion(new CellRangeAddress(1, 2, 5, 5));
                    sheet.AddMergedRegion(new CellRangeAddress(1, 2, 6, 6));
                    sheet.AddMergedRegion(new CellRangeAddress(1, 2, 7, 7));
                    sheet.AddMergedRegion(new CellRangeAddress(1, 2, 8, 8));
                    sheet.AddMergedRegion(new CellRangeAddress(1, 2, 9, 9));

                    #endregion

                    #region //配备岗位数据
                    for (int i = 0; i < list2.Count(); i++)
                    {
                        string coreStr = "";
                        string boneStr = "是";
                        if (1 == list2[i].coreNum)
                        {
                            coreStr = "是";
                            boneStr = "";
                        }

                        int m = i + 3;
                        var row = sheet.CreateRow(m);
                        row.Height = 20 * 25;

                        row.CreateCell(0).SetCellValue(i + 1);
                        row.GetCell(0).CellStyle = style1;
                        row.CreateCell(1).SetCellValue(list2[i].deptName);
                        row.GetCell(1).CellStyle = style1;
                        row.CreateCell(2).SetCellValue(list2[i].deptCode);
                        row.GetCell(2).CellStyle = style1;
                        row.CreateCell(3).SetCellValue(list2[i].postName);
                        row.GetCell(3).CellStyle = style1;
                        row.CreateCell(4).SetCellValue(list2[i].postCode);
                        row.GetCell(4).CellStyle = style1;
                        row.CreateCell(5).SetCellValue(list2[i].postLevel);
                        row.GetCell(5).CellStyle = style1;
                        row.CreateCell(6).SetCellValue(list2[i].costType);
                        row.GetCell(6).CellStyle = style1;
                        row.CreateCell(7).SetCellValue(list2[i].quotaWage);
                        row.GetCell(7).CellStyle = style1;
                        row.CreateCell(8).SetCellValue(coreStr);
                        row.GetCell(8).CellStyle = style1;
                        row.CreateCell(9).SetCellValue(boneStr);
                        row.GetCell(9).CellStyle = style1;

                        for (int j = 0; j < list1.Count(); j++)
                        {
                            row.CreateCell(10 + j).SetCellValue(list3.Where(p => p.rid == list1[j].id && p.gwid == list2[i].id).First().floatNum);
                            row.GetCell(10 + j).CellStyle = style1;
                        }
                    }
                    #endregion


                    for (int k = 0; k < colCount; k++)
                    {
                        if (k > 0 && k < 5)
                        {
                            //sheet.AutoSizeColumn(k);  //自适应宽度
                            sheet.SetColumnWidth(k, 20 * 256); //设置列宽
                        }
                        else
                        {
                            sheet.SetColumnWidth(k, 10 * 256); //设置列宽
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }


        #endregion


        #region //编码对照

        //异构系统编码对照
        [MyAuthAttribute]
        public ActionResult Hr_Middle_sys()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //中间编码对照 数据
        public string GetHr_Middle_sys()
        {
            try
            {
                JArray array = new JArray();  //返回结果

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Middle_sys(mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("mType", dt.Rows[i]["mType"].ToString());
                        _json.Add("easCode", dt.Rows[i]["easCode"].ToString());
                        _json.Add("easName", dt.Rows[i]["easName"].ToString());
                        _json.Add("baisCode", dt.Rows[i]["baisCode"].ToString());
                        _json.Add("baisName", dt.Rows[i]["baisName"].ToString());
                        _json.Add("pcMakerCode", dt.Rows[i]["pcMakerCode"].ToString());
                        _json.Add("pcMakerName", dt.Rows[i]["pcMakerName"].ToString());
                        _json.Add("crmCode", dt.Rows[i]["crmCode"].ToString());
                        _json.Add("crmName", dt.Rows[i]["crmName"].ToString());
                        _json.Add("addTime", dt.Rows[i]["addTime"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //Excel数据导入
        public void ImportHr_Middle_sys(ISheet sheet)
        {
            try
            {
                //组织单元集合(有权限)
                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string addPer = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var list = new List<Dictionary<string, string>>();

                int rowCount = sheet.LastRowNum + 1; //总行数
                for (int i = 1; i < rowCount; i++)
                {
                    //读到空行终止
                    var _row = sheet.GetRow(i);
                    if (_row == null)
                        break;
                    var _cell = sheet.GetRow(i).GetCell(0);
                    if (_cell == null)
                        break;
                    var _rowCell = sheet.GetRow(i).GetCell(0).StringCellValue;
                    if (string.IsNullOrWhiteSpace(_rowCell))
                        break;

                    string mType = sheet.GetRow(i).GetCell(0).ToString();
                    string easCode = sheet.GetRow(i).GetCell(1).ToString();
                    string easName = sheet.GetRow(i).GetCell(2).ToString();
                    string baisCode = sheet.GetRow(i).GetCell(3).ToString();
                    string baisName = sheet.GetRow(i).GetCell(4).ToString();
                    string pcMakerCode = sheet.GetRow(i).GetCell(5).ToString();
                    string pcMakerName = sheet.GetRow(i).GetCell(6).ToString();
                    string crmCode = sheet.GetRow(i).GetCell(7).ToString();
                    string crmName = sheet.GetRow(i).GetCell(8).ToString();

                    var conditions = new Dictionary<string, string>();
                    conditions.Add("mType", mType);
                    conditions.Add("easCode", easCode);
                    conditions.Add("easName", easName);
                    conditions.Add("baisCode", baisCode);
                    conditions.Add("baisName", baisName);
                    conditions.Add("pcMakerCode", pcMakerCode);
                    conditions.Add("pcMakerName", pcMakerName);
                    conditions.Add("crmCode", crmCode);
                    conditions.Add("crmName", crmName);
                    conditions.Add("mid", midStr);
                    conditions.Add("addPer", addPer);

                    list.Add(conditions);
                }

                if (list != null)
                {
                    var dal = new dalPro();
                    foreach (var item in list)
                    {
                        dal.AddHr_Middle_sys(item);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出
        public void ExportHr_Middle_sys()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("异构系统编码对照", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("异构系统编码对照"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("编码类别");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("eas编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("eas名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("bais编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("bais名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("pcMaker编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("pcMaker名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("crm编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("crm名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Middle_sys(mid); //数据

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["mType"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["easCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["easName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["baisCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["baisName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["pcMakerCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["pcMakerName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["crmCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["crmName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                    }
                }

                sheet.SetColumnWidth(0, 15 * 256); //设置列宽
                for (int k = 1; k < colCount; k++)
                {
                    sheet.AutoSizeColumn(k);  //自适应宽度
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }

        //删除
        public string DelHr_Middle_sys()
        {
            try
            {
                string ids = Request["ids"];

                var idList = ids.Split(',');

                var dal = new dalPro();
                foreach (var cc in idList)
                {
                    int _id = int.Parse(cc);
                    dal.DelHr_Middle_sysById(_id);
                }

                return "删除成功！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }


        //岗位配备规则编码对照
        [MyAuthAttribute]
        public ActionResult Hr_Middle_rule()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //中间编码对照 数据
        public string GetHr_Middle_rule()
        {
            try
            {
                JArray array = new JArray();  //返回结果

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Middle_rule(mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("mType", dt.Rows[i]["mType"].ToString());
                        _json.Add("ruleCode", dt.Rows[i]["ruleCode"].ToString());
                        _json.Add("ruleName", dt.Rows[i]["ruleName"].ToString());
                        _json.Add("easCode", dt.Rows[i]["easCode"].ToString());
                        _json.Add("easName", dt.Rows[i]["easName"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //Excel数据导入
        public void ImportHr_Middle_rule(ISheet sheet)
        {
            try
            {
                //组织单元集合(有权限)
                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string addPer = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var list = new List<Dictionary<string, string>>();

                int rowCount = sheet.LastRowNum + 1; //总行数
                for (int i = 1; i < rowCount; i++)
                {
                    //读到空行终止
                    var _row = sheet.GetRow(i);
                    if (_row == null)
                        break;
                    var _cell = sheet.GetRow(i).GetCell(0);
                    if (_cell == null)
                        break;
                    var _rowCell = sheet.GetRow(i).GetCell(0).StringCellValue;
                    if (string.IsNullOrWhiteSpace(_rowCell))
                        break;

                    string mType = sheet.GetRow(i).GetCell(0).ToString();
                    string ruleCode = sheet.GetRow(i).GetCell(1).ToString();
                    string ruleName = sheet.GetRow(i).GetCell(2).ToString();
                    string easCode = sheet.GetRow(i).GetCell(3).ToString();
                    string easName = sheet.GetRow(i).GetCell(4).ToString();

                    var conditions = new Dictionary<string, string>();
                    conditions.Add("mType", mType);
                    conditions.Add("ruleCode", ruleCode);
                    conditions.Add("ruleName", ruleName);
                    conditions.Add("easCode", easCode);
                    conditions.Add("easName", easName);
                    conditions.Add("mid", midStr);
                    conditions.Add("addPer", addPer);

                    list.Add(conditions);
                }

                if (list != null)
                {
                    var dal = new dalPro();
                    foreach (var item in list)
                    {
                        dal.AddHr_Middle_rule(item);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        
        //导出
        public void ExportHr_Middle_rule()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("岗位配备规则编码对照", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("岗位配备规则编码对照"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("编码类别");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("岗位配备编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("岗位配备名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("eas编码");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("eas名称");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Middle_rule(mid); //数据

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["mType"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["ruleCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["ruleName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["easCode"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["easName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                    }
                }

                sheet.SetColumnWidth(0, 15 * 256); //设置列宽
                for (int k = 1; k < colCount; k++)
                {
                    sheet.AutoSizeColumn(k);  //自适应宽度
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }

        //删除
        public string DelHr_Middle_rule()
        {
            try
            {
                string ids = Request["ids"];

                var idList = ids.Split(',');

                var dal = new dalPro();
                foreach (var cc in idList)
                {
                    int _id = int.Parse(cc);
                    dal.DelHr_Middle_ruleById(_id);
                }

                return "删除成功！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }


        #endregion


        //年度列表
        [HttpPost]
        public string GetYearly()
        {
            try
            {
                //接口数据
                webService.Service1 wbs = new webService.Service1();
                DataTable dt = wbs.GetEASYeardt();
                if (dt.Rows.Count > 0)
                {
                    JArray arry = new JArray();

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        JObject _json = new JObject();
                        _json.Add("value", dt.Rows[i]["EASYear"].ToString());
                        _json.Add("text", dt.Rows[i]["EASYear"].ToString());

                        arry.Add(_json);
                    }
                    return arry.ToString();
                }

                return "";

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //上传Excel文件
        public JsonResult UploadExcel()
        {
            HttpFileCollectionBase files = Request.Files;
            HttpPostedFileBase fileData = files[0];//获取上传的文件

            string type = Request["uploadType"]; //类型

            try
            {
                if (null != fileData)
                {
                    string fileName = Path.GetFileName(fileData.FileName);// 原始文件名称
                    
                    IWorkbook workbook = WorkbookFactory.Create(fileData.InputStream); //自动判别Excel版本 2003以前或2007及以上
                    ISheet sheet = workbook.GetSheetAt(0); //表数据                    

                    switch (type)
                    {
                        case "PcMaker_ghys": ImportHr_PcMaker_ghys(sheet); break;
                        case "PcMaker_xmjz": ImportHr_PcMaker_xmjz(sheet); break;
                        case "PcMaker_fact": ImportHr_PcMaker_fact(sheet); break;
                        case "Bais_ghys": ImportHr_Bais_ghys(sheet); break;
                        case "Bais_xmjz": ImportHr_Bais_xmjz(sheet); break;
                        case "Bais_wlyj": ImportHr_Bais_fact(sheet); break;
                        case "Bais_rgsr": ImportHr_Bais_rgsr(sheet); break;
                        case "Bais_rgzc": ImportHr_Bais_rgzc(sheet); break;
                        case "Crm_ghys": ImportHr_Crm_ghys(sheet); break;
                        case "Crm_fact": ImportHr_Crm_fact(sheet); break;
                        case "Bhr_fact": ImportHr_Bhr_fact(sheet); break;
                        case "Rule_ys": ImportHr_Rule_ys(sheet); break;
                        case "Rule_ht":
                            string htType = Request["htType"];
                            ImportHr_Rule_ht(sheet, htType);
                            break;
                        case "Middle_sys": ImportHr_Middle_sys(sheet); break;
                        case "Middle_rule": ImportHr_Middle_rule(sheet); break;
                        default: break;
                    }


                    // 文件上传后的保存路径
                    string filePath = Server.MapPath("~/Uploads/UploadFiles/");
                    if (!Directory.Exists(filePath))
                    {
                        Directory.CreateDirectory(filePath);
                    }

                    string path = Path.GetExtension(fileData.FileName); //文件后缀
                    fileData.SaveAs(filePath + fileName + path);//保存文件
                    
                    return Json(new { Success = true }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return Json(new { Success = false, error = "请重新选择上传文件！" }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                return Json(new { Success = false, error = ex.Message }, JsonRequestBehavior.AllowGet);
            }
            finally
            {
                files = null;
                fileData = null;
            }
        }


        #region //规划、调整预算，项目进展，实际

        //规划预算
        [MyAuthAttribute]
        public ActionResult Hr_Midghys()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //规划预算 数据
        public string GetHr_Midghys()
        {
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                JArray array = new JArray();  //返回结果

                //比较开始日期，结束日期
                var cc = DateTime.Compare(start, end);
                if (cc > 0)
                {
                    return array.ToString();  //开始日期大于结束日期,返回空
                }

                //组织单元
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);
                
                var dal = new dalPro();
                var dt = dal.GetHr_Midghys(start, end, mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("comName", dt.Rows[i]["comName"].ToString());
                        _json.Add("cxNum", dt.Rows[i]["cxNum"].ToString());
                        _json.Add("htAmount", dt.Rows[i]["htAmount"].ToString());
                        _json.Add("ysAmount", dt.Rows[i]["ysAmount"].ToString());
                        _json.Add("lrAmount", dt.Rows[i]["lrAmount"].ToString());
                        _json.Add("yield", dt.Rows[i]["yield"].ToString());
                        _json.Add("yieEffic", dt.Rows[i]["yieEffic"].ToString());
                        _json.Add("gjEffic", dt.Rows[i]["gjEffic"].ToString());
                        _json.Add("proTeams", dt.Rows[i]["proTeams"].ToString());
                        _json.Add("workDays", dt.Rows[i]["workDays"].ToString());
                        _json.Add("yearly", dt.Rows[i]["yearly"].ToString());
                        _json.Add("monthly", dt.Rows[i]["monthly"].ToString());
                        _json.Add("mid", dt.Rows[i]["mid"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出规划预算列表
        public void ExportHr_Midghys()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("规划预算", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("规划预算"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("公 司");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("年 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("月 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产线数");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("合同额（万元）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("营收（万元）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("利润（万元）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产量（万立方）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产效（立方/人/8H）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("构件产效（立方/人/8H）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("项目组数");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("工作天数");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            
            #endregion

            #region //表体
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Midghys(start, end, mid);
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["comName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yearly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["monthly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["cxNum"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["htAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["ysAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["lrAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yield"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yieEffic"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["gjEffic"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["proTeams"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["workDays"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        
                    }
                }

                sheet.AutoSizeColumn(0);  //自适应宽度
                for (int k = 1; k < colCount; k++)
                {
                    //设置列宽
                    if (k > 3 && k < 10)
                    {
                        sheet.SetColumnWidth(k, 30 * 256);
                    }
                    else
                    {
                        sheet.SetColumnWidth(k, 15 * 256);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }

        //删除
        public string DelHr_Midghys()
        {
            try
            {
                string ids = Request["ids"];
                
                var idList = ids.Split(',');

                var dal = new dalPro();
                foreach (var cc in idList)
                {
                    int _id = int.Parse(cc);
                    dal.DelHr_MidghysById(_id);
                }

                return "删除成功！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //// 同步数据
        public string SynHr_Midghys()
        {
            try
            {
                int yearly = int.Parse(Request["yearly"]);
                int monthly = int.Parse(Request["monthly"]);

                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string name = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var dal = new dalPro();
                string result = "";

                //PcMaker 数据源
                var dt_pc = dal.GetHr_PcMaker_ghysByMonth(yearly, monthly, midStr);
                if (dt_pc.Rows.Count == 0)
                {
                    result = "该月度无PcMaker数据！";
                    return result;
                }

                //Bais 数据源
                var dt_bs = dal.GetHr_Bais_ghysByMonth(yearly, monthly, midStr);
                if (dt_bs.Rows.Count == 0)
                {
                    result = "该月度无Bais数据！";
                    return result;
                }

                //Crm 数据源
                var dt_crm = dal.GetHr_Crm_ghysByMonth(yearly, monthly, midStr);
                if (dt_crm.Rows.Count == 0)
                {
                    result = "该月度无CRM数据！";
                    return result;
                }

                var condition = new Dictionary<string, string>();
                condition.Add("yearly", yearly.ToString());
                condition.Add("monthly", monthly.ToString());
                condition.Add("easComCode", dt_pc.Rows[0]["easCode"].ToString()); //公司 eas编码
                condition.Add("cxNum", dt_pc.Rows[0]["cxNum"].ToString());
                condition.Add("proTeams", dt_crm.Rows.Count.ToString());   //项目组数=crm该月度部门个数
                condition.Add("htAmount", dt_bs.Rows[0]["htAmount"].ToString());
                condition.Add("ysAmount", dt_bs.Rows[0]["ysAmount"].ToString());
                condition.Add("lrAmount", dt_bs.Rows[0]["lrAmount"].ToString());
                condition.Add("yield", dt_bs.Rows[0]["yield"].ToString());
                condition.Add("yieEffic", dt_pc.Rows[0]["yieEffic"].ToString());
                condition.Add("gjEffic", dt_pc.Rows[0]["gjEffic"].ToString());
                condition.Add("workDays", dt_pc.Rows[0]["workDays"].ToString());
                condition.Add("mid", midStr);
                condition.Add("addPer", name);

                int i = dal.AddHr_Midghys(condition);
                if (i > 0)
                {
                    result = "同步成功！";
                }
                else
                {
                    result = "同步失败！";
                }

                return result;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        

        //调整预算
        [MyAuthAttribute]
        public ActionResult Hr_Midtzys()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //调整预算 数据
        public string GetHr_Midtzys()
        {
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                JArray array = new JArray();  //返回结果

                //比较开始日期，结束日期
                var cc = DateTime.Compare(start, end);
                if (cc > 0)
                {
                    return array.ToString();  //开始日期大于结束日期,返回空
                }

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Midtzys(start, end, mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("comName", dt.Rows[i]["comName"].ToString());
                        _json.Add("cxNum", dt.Rows[i]["cxNum"].ToString());
                        _json.Add("htAmount", dt.Rows[i]["htAmount"].ToString());
                        _json.Add("ysAmount", dt.Rows[i]["ysAmount"].ToString());
                        _json.Add("lrAmount", dt.Rows[i]["lrAmount"].ToString());
                        _json.Add("yield", dt.Rows[i]["yield"].ToString());
                        _json.Add("yieEffic", dt.Rows[i]["yieEffic"].ToString());
                        _json.Add("gjEffic", dt.Rows[i]["gjEffic"].ToString());
                        _json.Add("proTeams", dt.Rows[i]["proTeams"].ToString());
                        _json.Add("workDays", dt.Rows[i]["workDays"].ToString());
                        _json.Add("yearly", dt.Rows[i]["yearly"].ToString());
                        _json.Add("monthly", dt.Rows[i]["monthly"].ToString());
                        _json.Add("mid", dt.Rows[i]["mid"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出调整预算列表
        public void ExportHr_Midtzys()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("调整预算", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("调整预算"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("公 司");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("年 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("月 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产线数");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("合同额（万元）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("营收（万元）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("利润（万元）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产量（万立方）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产效（立方/人/8H）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("构件产效（立方/人/8H）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("项目组数");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("工作天数");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Midtzys(start, end, mid);
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["comName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yearly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["monthly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["cxNum"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["htAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["ysAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["lrAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yield"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yieEffic"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["gjEffic"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["proTeams"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["workDays"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;

                    }
                }

                sheet.AutoSizeColumn(0);  //自适应宽度
                for (int k = 1; k < colCount; k++)
                {
                    //设置列宽
                    if (k > 3 && k < 10)
                    {
                        sheet.SetColumnWidth(k, 30 * 256);
                    }
                    else
                    {
                        sheet.SetColumnWidth(k, 15 * 256);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }

        //删除
        public string DelHr_Midtzys()
        {
            try
            {
                string ids = Request["ids"];

                var idList = ids.Split(',');

                var dal = new dalPro();
                foreach (var cc in idList)
                {
                    int _id = int.Parse(cc);
                    dal.DelHr_MidtzysById(_id);
                }

                return "删除成功！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //// 同步数据
        public string SynHr_Midtzys()
        {
            try
            {
                int yearly = int.Parse(Request["yearly"]);
                int monthly = int.Parse(Request["monthly"]);

                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string name = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var dal = new dalPro();
                string result = "";

                //规划预算 数据源
                var dt = dal.GetHr_MidghysByMonth(yearly, monthly, midStr);
                if (dt.Rows.Count == 0)
                {
                    result = "该月度无规划数据！";
                }
                else
                {
                    var condition = new Dictionary<string, string>();
                    condition.Add("yearly", yearly.ToString());
                    condition.Add("monthly", monthly.ToString());
                    condition.Add("easComCode", dt.Rows[0]["easComCode"].ToString()); //公司 eas编码
                    condition.Add("cxNum", dt.Rows[0]["cxNum"].ToString());
                    condition.Add("proTeams", dt.Rows[0]["proTeams"].ToString());   
                    condition.Add("htAmount", dt.Rows[0]["htAmount"].ToString());
                    condition.Add("ysAmount", dt.Rows[0]["ysAmount"].ToString());
                    condition.Add("lrAmount", dt.Rows[0]["lrAmount"].ToString());
                    condition.Add("yield", dt.Rows[0]["yield"].ToString());
                    condition.Add("yieEffic", dt.Rows[0]["yieEffic"].ToString());
                    condition.Add("gjEffic", dt.Rows[0]["gjEffic"].ToString());
                    condition.Add("workDays", dt.Rows[0]["workDays"].ToString());
                    condition.Add("mid", midStr);
                    condition.Add("addPer", name);

                    int i = dal.AddHr_Midtzys(condition);
                    if (i > 0)
                    {
                        result = "同步成功！";
                    }
                    else
                    {
                        result = "同步失败！";
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        ////通过id修改 调整预算合同额
        public string SaveHr_Midtzys()
        {
            try
            {
                int id = int.Parse(Request["id"]);

                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string name = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var dal = new dalPro();

                var condition = new Dictionary<string, string>();
                condition.Add("cxNum", Request["cxNum"]);
                condition.Add("htAmount", Request["htAmount"]);
                condition.Add("ysAmount", Request["ysAmount"]);
                condition.Add("lrAmount", Request["lrAmount"]);
                condition.Add("yield", Request["yieldVal"]);
                condition.Add("yieEffic", Request["yieEffic"]);
                condition.Add("gjEffic", Request["gjEffic"]);
                condition.Add("proTeams", Request["proTeams"]);
                condition.Add("workDays", Request["workDays"]);
                condition.Add("mid", midStr);
                condition.Add("addPer", name);

                int i = dal.EditHr_MidtzysById(id, condition);

                return i > 0 ? "修改成功！" : "修改失败！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }


        //项目进展
        [MyAuthAttribute]
        public ActionResult Hr_Midxmjz()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //项目进展 数据
        public string GetHr_Midxmjz()
        {
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                JArray array = new JArray();  //返回结果

                //比较开始日期，结束日期
                var cc = DateTime.Compare(start, end);
                if (cc > 0)
                {
                    return array.ToString();  //开始日期大于结束日期,返回空
                }

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Midxmjz(start, end, mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("comName", dt.Rows[i]["comName"].ToString());
                        _json.Add("cxNum", dt.Rows[i]["cxNum"].ToString());
                        _json.Add("htAmount", dt.Rows[i]["htAmount"].ToString());
                        _json.Add("ysAmount", dt.Rows[i]["ysAmount"].ToString());
                        _json.Add("lrAmount", dt.Rows[i]["lrAmount"].ToString());
                        _json.Add("proBudget", dt.Rows[i]["proBudget"].ToString());
                        _json.Add("progjBudget", dt.Rows[i]["progjBudget"].ToString());
                        _json.Add("yieEffic", dt.Rows[i]["yieEffic"].ToString());
                        _json.Add("gjEffic", dt.Rows[i]["gjEffic"].ToString());
                        _json.Add("proTeams", dt.Rows[i]["proTeams"].ToString());
                        _json.Add("workDays", dt.Rows[i]["workDays"].ToString());
                        _json.Add("yearly", dt.Rows[i]["yearly"].ToString());
                        _json.Add("monthly", dt.Rows[i]["monthly"].ToString());
                        _json.Add("mid", dt.Rows[i]["mid"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出项目进展列表
        public void ExportHr_Midxmjz()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("项目进展", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("项目进展"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("公 司");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("年 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("月 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产线数");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("合同额（万元）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("营收（万元）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("利润（万元）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产量（万立方）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("构件产量（万立方）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产效（立方/人/8H）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("构件产效（立方/人/8H）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("项目组数");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("工作天数");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Midxmjz(start, end, mid);
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["comName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yearly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["monthly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["cxNum"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["htAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["ysAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["lrAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["proBudget"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["progjBudget"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yieEffic"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["gjEffic"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["proTeams"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["workDays"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;

                    }
                }

                sheet.AutoSizeColumn(0);  //自适应宽度
                for (int k = 1; k < colCount; k++)
                {
                    //设置列宽
                    if (k > 3 && k < 10)
                    {
                        sheet.SetColumnWidth(k, 30 * 256);
                    }
                    else
                    {
                        sheet.SetColumnWidth(k, 15 * 256);
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }

        //删除
        public string DelHr_Midxmjz()
        {
            try
            {
                string ids = Request["ids"];

                var idList = ids.Split(',');

                var dal = new dalPro();
                foreach (var cc in idList)
                {
                    int _id = int.Parse(cc);
                    dal.DelHr_MidxmjzById(_id);
                }

                return "删除成功！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //// 同步数据
        public string SynHr_Midxmjz()
        {
            try
            {
                int yearly = int.Parse(Request["yearly"]);
                int monthly = int.Parse(Request["monthly"]);

                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string name = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var dal = new dalPro();
                string result = "";

                //PcMaker 数据源
                var dt_pc = dal.GetHr_PcMaker_xmjzByMonth(yearly, monthly, midStr);
                if (dt_pc.Rows.Count == 0)
                {
                    result = "该月度无PcMaker数据！";
                    return result;
                }

                //Bais 数据源
                var dt_bs = dal.GetHr_Bais_xmjzByMonth(yearly, monthly, midStr);
                if (dt_bs.Rows.Count == 0)
                {
                    result = "该月度无Bais数据！";
                    return result;
                }

                var condition = new Dictionary<string, string>();
                condition.Add("yearly", yearly.ToString());
                condition.Add("monthly", monthly.ToString());
                condition.Add("easComCode", dt_pc.Rows[0]["easCode"].ToString()); //公司 eas编码
                condition.Add("cxNum", dt_pc.Rows[0]["cxNum"].ToString());
                condition.Add("htAmount", dt_bs.Rows[0]["htAmount"].ToString());
                condition.Add("ysAmount", dt_bs.Rows[0]["ysAmount"].ToString());
                condition.Add("lrAmount", dt_bs.Rows[0]["lrAmount"].ToString());
                condition.Add("proBudget", dt_pc.Rows[0]["proBudget"].ToString());
                condition.Add("progjBudget", dt_pc.Rows[0]["progjBudget"].ToString());
                condition.Add("yieEffic", dt_pc.Rows[0]["yieEffic"].ToString());
                condition.Add("gjEffic", dt_pc.Rows[0]["gjEffic"].ToString());
                condition.Add("proTeams", "1");
                condition.Add("workDays", dt_pc.Rows[0]["workDays"].ToString());
                condition.Add("mid", midStr);
                condition.Add("addPer", name);

                int i = dal.AddHr_Midxmjz(condition);
                if (i > 0)
                {
                    result = "同步成功！";
                }
                else
                {
                    result = "同步失败！";
                }

                return result;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }


        //实际
        [MyAuthAttribute]
        public ActionResult Hr_Midfact()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //实际 数据
        public string GetHr_Midfact()
        {
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                JArray array = new JArray();  //返回结果

                //比较开始日期，结束日期
                var cc = DateTime.Compare(start, end);
                if (cc > 0)
                {
                    return array.ToString();  //开始日期大于结束日期,返回空
                }

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Midfact(start, end, mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("comName", dt.Rows[i]["comName"].ToString());
                        _json.Add("cxNum", dt.Rows[i]["cxNum"].ToString());
                        _json.Add("htAmount", dt.Rows[i]["htAmount"].ToString());
                        _json.Add("ysAmount", dt.Rows[i]["ysAmount"].ToString());
                        _json.Add("lrAmount", dt.Rows[i]["lrAmount"].ToString());
                        _json.Add("yield", dt.Rows[i]["yield"].ToString());
                        _json.Add("yieEffic", dt.Rows[i]["yieEffic"].ToString());
                        _json.Add("gjEffic", dt.Rows[i]["gjEffic"].ToString());
                        _json.Add("proTeams", dt.Rows[i]["proTeams"].ToString());
                        _json.Add("workDays", dt.Rows[i]["workDays"].ToString());
                        _json.Add("yearly", dt.Rows[i]["yearly"].ToString());
                        _json.Add("monthly", dt.Rows[i]["monthly"].ToString());
                        _json.Add("mid", dt.Rows[i]["mid"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出实际列表
        public void ExportHr_Midfact()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("实际", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("实际"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("公 司");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("年 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("月 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产线数");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("合同额（万元）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("营收（万元）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("利润（万元）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产量（万立方）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("产效（立方/人/8H）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("构件产效（立方/人/8H）");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("项目组数");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("工作天数");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Midfact(start, end, mid);
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["comName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yearly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["monthly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["cxNum"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["htAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["ysAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["lrAmount"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yield"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yieEffic"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["gjEffic"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["proTeams"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["workDays"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;

                    }
                }

                sheet.AutoSizeColumn(0);  //自适应宽度
                for (int k = 1; k < colCount; k++)
                {
                    //设置列宽
                    if (k > 3 && k < 10) 
                    {
                        sheet.SetColumnWidth(k, 30 * 256); 
                    }
                    else
                    {
                        sheet.SetColumnWidth(k, 15 * 256);
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }

        //删除
        public string DelHr_Midfact()
        {
            try
            {
                string ids = Request["ids"];

                var idList = ids.Split(',');

                var dal = new dalPro();
                foreach (var cc in idList)
                {
                    int _id = int.Parse(cc);
                    dal.DelHr_MidfactById(_id);
                }

                return "删除成功！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //// 同步数据
        public string SynHr_Midfact()
        {
            try
            {
                int yearly = int.Parse(Request["yearly"]);
                int monthly = int.Parse(Request["monthly"]);

                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string name = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var dal = new dalPro();
                string result = "";

                //PcMaker 数据源
                var dt_pc = dal.GetHr_PcMaker_factByMonth(yearly, monthly, midStr);
                if (dt_pc.Rows.Count == 0)
                {
                    result = "该月度无PcMaker数据！";
                    return result;
                }

                //Bais 数据源
                var dt_bs = dal.GetHr_Bais_factByMonth(yearly, monthly, midStr);
                if (dt_bs.Rows.Count == 0)
                {
                    result = "该月度无Bais数据！";
                    return result;
                }

                //Bhr 数据源
                var dt_hr = dal.GetHr_Bhr_factByMonth(yearly, monthly, midStr);
                if (dt_hr.Rows.Count == 0)
                {
                    result = "该月度无Bhr数据！";
                    return result;
                }

                List<Hr_Bhr_fact> list = new List<Hr_Bhr_fact>();
                for (int i = 0; i < dt_hr.Rows.Count; i++)
                {
                    var obj = new Hr_Bhr_fact();

                    obj.easComCode = dt_hr.Rows[i]["easComCode"].ToString();
                    obj.easDeptCode = dt_hr.Rows[i]["easDeptCode"].ToString();
                    obj.easPostCode = dt_hr.Rows[i]["easPostCode"].ToString();
                    obj.easPostName = dt_hr.Rows[i]["easPostName"].ToString();
                    obj.easDeptName = dt_hr.Rows[i]["easDeptName"].ToString();
                    obj.postLevel = dt_hr.Rows[i]["postLevel"].ToString();
                    obj.postType = dt_hr.Rows[i]["postType"].ToString();
                    obj.workDays = double.Parse(dt_hr.Rows[i]["workDays"].ToString());

                    list.Add(obj);
                }
                
                //项目组数：取客户一部，客户二部 一直到 客户N部  的部门统计数
                double _proTeams = list.Where(p => p.easDeptName.StartsWith("客户") && p.easDeptName.EndsWith("部")).Count();

                //工作天数：BHR（岗位职级为'OO'人员的实际工作天数之和）
                double _workDays = list.Where(p => p.postLevel == "OO").Sum(p => p.workDays);                

                var condition = new Dictionary<string, string>();
                condition.Add("yearly", yearly.ToString());
                condition.Add("monthly", monthly.ToString());
                condition.Add("easComCode", dt_pc.Rows[0]["easCode"].ToString()); //公司 eas编码
                condition.Add("cxNum", dt_pc.Rows[0]["cxNum"].ToString());
                condition.Add("htAmount", dt_bs.Rows[0]["htAmount"].ToString());
                condition.Add("ysAmount", dt_bs.Rows[0]["ysAmount"].ToString());
                condition.Add("lrAmount", dt_bs.Rows[0]["lrAmount"].ToString());
                condition.Add("yield", dt_bs.Rows[0]["yield"].ToString());
                condition.Add("yieEffic", dt_pc.Rows[0]["yieEffic"].ToString());
                condition.Add("gjEffic", dt_pc.Rows[0]["gjEffic"].ToString());
                condition.Add("proTeams", _proTeams.ToString());
                condition.Add("workDays", _workDays.ToString());
                condition.Add("mid", midStr);
                condition.Add("addPer", name);

                int cc = dal.AddHr_Midfact(condition);
                if (cc > 0)
                {
                    result = "同步成功！";
                }
                else
                {
                    result = "同步失败！";
                }

                return result;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        #endregion

        #region //(非)市场人数
        
        //非市场人数
        [MyAuthAttribute]
        public ActionResult Hr_Midysrs()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //非市场人数 数据
        public string GetHr_Midysrs()
        {
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                JArray array = new JArray();  //返回结果

                //比较开始日期，结束日期
                var cc = DateTime.Compare(start, end);
                if (cc > 0)
                {
                    return array.ToString();  //开始日期大于结束日期,返回空
                }

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Midysrs(start, end, mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("comName", dt.Rows[i]["comName"].ToString());
                        _json.Add("deptName", dt.Rows[i]["deptName"].ToString());
                        _json.Add("postName", dt.Rows[i]["postName"].ToString());
                        _json.Add("coreQuota", dt.Rows[i]["coreQuota"].ToString());
                        _json.Add("coreActual", dt.Rows[i]["coreActual"].ToString());
                        _json.Add("boneQuota", dt.Rows[i]["boneQuota"].ToString());
                        _json.Add("boneActual", dt.Rows[i]["boneActual"].ToString());
                        _json.Add("floatQuota", dt.Rows[i]["floatQuota"].ToString());
                        _json.Add("floatActual", dt.Rows[i]["floatActual"].ToString());
                        _json.Add("floatFore", dt.Rows[i]["floatFore"].ToString());
                        _json.Add("floatghys", dt.Rows[i]["floatghys"].ToString());
                        _json.Add("floattzys", dt.Rows[i]["floattzys"].ToString());
                        _json.Add("yearly", dt.Rows[i]["yearly"].ToString());
                        _json.Add("monthly", dt.Rows[i]["monthly"].ToString());
                        _json.Add("mid", dt.Rows[i]["mid"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出
        public void ExportHr_Midysrs()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("非市场人数", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("非市场人数"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("公 司");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("部 门");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("岗 位");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("年 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("月 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("核心人数(定额)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("核心人数(实际)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("骨干人数(定额)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("骨干人数(实际)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("浮动人数(规划预算)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("浮动人数(调整预算)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("浮动人数(定额)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("浮动人数(实际)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("浮动人数(预测)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Midysrs(start, end, mid);
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["comName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["deptName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["postName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yearly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["monthly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["coreQuota"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["coreActual"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["boneQuota"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["boneActual"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["floatghys"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["floattzys"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["floatQuota"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["floatActual"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["floatFore"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;

                    }
                }

                for (int k = 1; k < colCount; k++)
                {
                    //设置列宽
                    if (k < 3)
                    {
                        sheet.AutoSizeColumn(0);  //自适应宽度
                    }
                    else if (k > 4)
                    {
                        sheet.SetColumnWidth(k, 20 * 256);
                    }
                    else
                    {
                        sheet.SetColumnWidth(k, 15 * 256);
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }
        
        //删除
        public string DelHr_Midysrs()
        {
            try
            {
                string ids = Request["ids"];

                var idList = ids.Split(',');

                var dal = new dalPro();
                foreach (var cc in idList)
                {
                    int _id = int.Parse(cc);
                    dal.DelHr_MidysrsById(_id);
                }

                return "删除成功！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //// 同步数据
        public string SynHr_Midysrs()
        {
            try
            {
                int yearly = int.Parse(Request["yearly"]);
                int monthly = int.Parse(Request["monthly"]);

                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string name = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var dal = new dalPro();
                string result = "";

                //规划预算-产量
                var dt_gh = dal.GetHr_MidghysByMonth(yearly, monthly, midStr);
                if (dt_gh.Rows.Count == 0)
                {
                    result = "该月度无规划预算数据！";
                    return result;
                }

                //调整预算-产量
                var dt_tz = dal.GetHr_MidghysByMonth(yearly, monthly, midStr);
                if (dt_tz.Rows.Count == 0)
                {
                    result = "该月度无调整预算数据！";
                    return result;
                }

                //pcMaker 项目进展-产量，构件产量
                var dt_pc = dal.GetHr_PcMaker_xmjzByMonth(yearly, monthly, midStr);
                if (dt_pc.Rows.Count == 0)
                {
                    result = "该月度无PcMaker项目进展数据！";
                    return result;
                }

                //Bhr 数据源
                List<Hr_Bhr_fact> list_hr = null;
                var dt_hr = dal.GetHr_Bhr_factByMonth2(yearly, monthly, midStr);
                if (dt_hr.Rows.Count > 0)
                {
                    list_hr = new List<Hr_Bhr_fact>();
                    for (int i = 0; i < dt_hr.Rows.Count; i++)
                    {
                        var _obj = new Hr_Bhr_fact();

                        _obj.easComCode = dt_hr.Rows[i]["easComCode"].ToString();
                        _obj.ruleDeptCode = dt_hr.Rows[i]["ruleDeptCode"].ToString();
                        _obj.rulePostCode = dt_hr.Rows[i]["rulePostCode"].ToString();
                        _obj.postType = dt_hr.Rows[i]["postType"].ToString();

                        list_hr.Add(_obj);
                    }
                }

                //岗位配备规则
                List<Hr_Rule_yshead> list1 = null; // 营收规则
                List<Hr_Rule_ysgw> list2 = null;   // 岗位配备人数
                List<Hr_Rule_ysfd> list3 = null;   // 岗位配备浮动人数

                #region //岗位配备规则数据
                var dt1 = dal.GetHr_Rule_yshead(yearly, midStr);
                var dt2 = dal.GetHr_Rule_ysgw(yearly, midStr);
                var dt3 = dal.GetHr_Rule_ysfd(yearly, midStr);
                if (dt1.Rows.Count > 0)
                {
                    list1 = new List<Hr_Rule_yshead>();
                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        var obj1 = new Hr_Rule_yshead();
                        obj1.id = dt1.Rows[i]["id"].ToString();
                        obj1.yield = double.Parse(dt1.Rows[i]["yield"].ToString());

                        list1.Add(obj1);
                    }
                }

                if (dt2.Rows.Count > 0)
                {
                    list2 = new List<Hr_Rule_ysgw>();
                    for (int i = 0; i < dt2.Rows.Count; i++)
                    {
                        var obj2 = new Hr_Rule_ysgw();
                        obj2.id = dt2.Rows[i]["id"].ToString();
                        obj2.deptCode = dt2.Rows[i]["deptCode"].ToString();
                        obj2.postCode = dt2.Rows[i]["postCode"].ToString();
                        obj2.coreNum = int.Parse(dt2.Rows[i]["coreNum"].ToString());
                        obj2.boneNum = int.Parse(dt2.Rows[i]["boneNum"].ToString());

                        list2.Add(obj2);
                    }
                }

                if (dt3.Rows.Count > 0)
                {
                    list3 = new List<Hr_Rule_ysfd>();
                    for (int i = 0; i < dt3.Rows.Count; i++)
                    {
                        var obj3 = new Hr_Rule_ysfd();
                        obj3.rid = dt3.Rows[i]["rid"].ToString();
                        obj3.gwid = dt3.Rows[i]["gwid"].ToString();
                        obj3.floatNum = int.Parse(dt3.Rows[i]["floatNum"].ToString());

                        list3.Add(obj3);
                    }
                }
                #endregion

                string easComCode = dt_gh.Rows[0]["easComCode"].ToString();  //eas公司编码
                double ghysYield = double.Parse(dt_gh.Rows[0]["yield"].ToString()); //规划预算-产量
                double tzysYield = double.Parse(dt_tz.Rows[0]["yield"].ToString()); //调整预算-产量
                double proBudget = double.Parse(dt_pc.Rows[0]["proBudget"].ToString());    //PcMaker 项目进展产量
                double progjBudget = double.Parse(dt_pc.Rows[0]["progjBudget"].ToString()); //PcMaker 项目进展构件产量

                List<Hr_Midysrs> list = new List<Hr_Midysrs>();
                foreach (var item in list2)
                {
                    var obj = new Hr_Midysrs();

                    obj.easComCode = easComCode;
                    obj.ruleDeptCode = item.deptCode;
                    obj.rulePostCode = item.postCode;
                    
                    obj.coreQuota = item.coreNum; //定额核心
                    obj.boneQuota = item.boneNum; //定额骨干

                    //部门，岗位编码相同，岗位类别不同的 实际人数
                    if (null != list_hr)
                    {
                        obj.coreActual = list_hr.Where(p => p.postType == "核心" && p.ruleDeptCode == item.deptCode && p.rulePostCode == item.postCode).Count();
                        obj.boneActual = list_hr.Where(p => p.postType == "骨干" && p.ruleDeptCode == item.deptCode && p.rulePostCode == item.postCode).Count();
                        obj.floatActual = list_hr.Where(p => p.postType == "浮动" && p.ruleDeptCode == item.deptCode && p.rulePostCode == item.postCode).Count();
                    }
                    else
                    {
                        obj.coreActual = 0;
                        obj.boneActual = 0;
                        obj.floatActual = 0;
                    }

                    //浮动人数
                    /*
                       * 根据产量 匹配 营收规则配备表里的月度产量区间 月度量在a与b之间 取b的人数
                       * 根据构件月度量匹配 营收规则配备表里的月度产量区间 月度量在a与b之间 取b的人数
                       * 超过区间最大值，报错异常
                       */

                    //筛选出 浮动人数 对应的规则id
                    string ghys_rid = "";
                    string tzys_rid = "";
                    string pro_rid = "";
                    string progj_rid = "";
                    foreach (var item2 in list1)
                    {
                        if (ghysYield <= item2.yield && ghys_rid == "")
                        {
                            ghys_rid = item2.id;
                        }
                        if (tzysYield <= item2.yield && tzys_rid == "")
                        {
                            tzys_rid = item2.id;
                        }
                        if (proBudget <= item2.yield && pro_rid == "")
                        {
                            pro_rid = item2.id;
                        }
                        if (progjBudget <= item2.yield && progj_rid == "")
                        {
                            progj_rid = item2.id;
                        }
                    }

                    //根据规则id和岗位配备人数id 取浮动人数
                    obj.floatQuota = list3.Where(p => p.rid == progj_rid && p.gwid == item.id).First().floatNum; //定额浮动
                    obj.floatQuota = list3.Where(p => p.rid == pro_rid && p.gwid == item.id).First().floatNum;  //预测浮动
                    obj.floatQuota = list3.Where(p => p.rid == ghys_rid && p.gwid == item.id).First().floatNum; //规划预算浮动
                    obj.floatQuota = list3.Where(p => p.rid == tzys_rid && p.gwid == item.id).First().floatNum; //调整预算浮动
                    
                    list.Add(obj);
                }

                //保存数据
                foreach (var item in list)
                {
                    var condition = new Dictionary<string, string>();
                    condition.Add("easComCode", item.easComCode); //公司 eas编码
                    condition.Add("ruleDeptCode", item.ruleDeptCode);
                    condition.Add("rulePostCode", item.rulePostCode);
                    condition.Add("coreQuota", item.coreQuota.ToString());
                    condition.Add("coreActual", item.coreActual.ToString());
                    condition.Add("boneQuota", item.boneQuota.ToString());
                    condition.Add("boneActual", item.boneActual.ToString());
                    condition.Add("floatQuota", item.floatQuota.ToString());
                    condition.Add("floatActual", item.floatActual.ToString());
                    condition.Add("floatFore", item.floatFore.ToString());
                    condition.Add("floatghys", item.floatghys.ToString());
                    condition.Add("floattzys", item.floattzys.ToString());
                    condition.Add("yearly", yearly.ToString());
                    condition.Add("monthly", monthly.ToString());
                    condition.Add("mid", midStr);
                    condition.Add("addPer", name);

                    dal.AddHr_Midysrs(condition);
                }

                result = "同步成功！";
                return result;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }


        //市场人数
        [MyAuthAttribute]
        public ActionResult Hr_Midhtrs()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //市场人数数据
        public string GetHr_Midhtrs()
        {
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                JArray array = new JArray();  //返回结果

                //比较开始日期，结束日期
                var cc = DateTime.Compare(start, end);
                if (cc > 0)
                {
                    return array.ToString();  //开始日期大于结束日期,返回空
                }

                //组织单元
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Midhtrs(start, end, mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("comName", dt.Rows[i]["comName"].ToString());
                        _json.Add("deptName", dt.Rows[i]["deptName"].ToString());
                        _json.Add("postName", dt.Rows[i]["postName"].ToString());
                        _json.Add("core_ghys", dt.Rows[i]["core_ghys"].ToString());
                        _json.Add("core_tzys", dt.Rows[i]["core_tzys"].ToString());
                        _json.Add("bone_ghys", dt.Rows[i]["bone_ghys"].ToString());
                        _json.Add("bone_tzys", dt.Rows[i]["bone_tzys"].ToString());
                        _json.Add("yearly", dt.Rows[i]["yearly"].ToString());
                        _json.Add("monthly", dt.Rows[i]["monthly"].ToString());
                        _json.Add("mid", dt.Rows[i]["mid"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出
        public void ExportHr_Midhtrs()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("市场人数", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("市场人数"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("公 司");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("部 门");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("岗 位");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("年 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("月 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("核心人数(规划预算)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("核心人数(调整预算)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("骨干人数(规划预算)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("骨干人数(调整预算)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Midhtrs(start, end, mid);
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["comName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["deptName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["postName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yearly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["monthly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["core_ghys"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["core_tzys"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["bone_ghys"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["bone_tzys"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;

                    }
                }

                for (int k = 1; k < colCount; k++)
                {
                    //设置列宽
                    if (k < 3)
                    {
                        sheet.AutoSizeColumn(0);  //自适应宽度
                    }
                    else if (k > 4)
                    {
                        sheet.SetColumnWidth(k, 20 * 256);
                    }
                    else
                    {
                        sheet.SetColumnWidth(k, 15 * 256);
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }

        //删除
        public string DelHr_Midhtrs()
        {
            try
            {
                string ids = Request["ids"];

                var idList = ids.Split(',');

                var dal = new dalPro();
                foreach (var cc in idList)
                {
                    int _id = int.Parse(cc);
                    dal.DelHr_MidhtrsById(_id);
                }

                return "删除成功！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //// 同步数据
        public string SynHr_Midhtrs()
        {
            try
            {
                int yearly = int.Parse(Request["yearly"]);
                int monthly = int.Parse(Request["monthly"]);

                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string name = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var dal = new dalPro();
                string result = "";

                #region//Crm 调整预算
                List<Hr_Crm_tzys> list_crm = null;
                var dt = dal.GetHr_Crm_tzysByMonth(yearly, monthly, midStr);
                if (dt.Rows.Count == 0)
                {
                    result = "该月度无Crm调整预算数据！";
                    return result;
                }
                else
                {
                    list_crm = new List<Hr_Crm_tzys>();
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _obj = new Hr_Crm_tzys();

                        _obj.easComCode = dt.Rows[i]["easComCode"].ToString();
                        _obj.ruleDeptCode = dt.Rows[i]["ruleDeptCode"].ToString();
                        _obj.ghsyAmount = double.Parse(dt.Rows[i]["ghsyAmount"].ToString());
                        _obj.tzsyAmount = double.Parse(dt.Rows[i]["tzsyAmount"].ToString());

                        list_crm.Add(_obj);
                    }
                }
                #endregion

                string htType = "单个项目组";

                #region//岗位配备规则-单个项目组
                List<Hr_Rule_hthead> list1 = null; // 合同规则
                List<Hr_Rule_htgw> list2 = null;   // 岗位配备
                List<Hr_Rule_htfd> list3 = null;   // 岗位配备人数


                var dt1 = dal.GetHr_Rule_hthead(yearly, htType, midStr);// 合同规则
                var dt2 = dal.GetHr_Rule_htgw(yearly, htType, midStr);  // 岗位配备
                var dt3 = dal.GetHr_Rule_htfd(yearly, htType, midStr);  // 岗位配备人数
                if (dt1.Rows.Count > 0)
                {
                    list1 = new List<Hr_Rule_hthead>();
                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        var obj1 = new Hr_Rule_hthead();
                        obj1.id = dt1.Rows[i]["id"].ToString();
                        obj1.htType = dt1.Rows[i]["htType"].ToString();
                        obj1.htAmount = double.Parse(dt1.Rows[i]["htAmount"].ToString());

                        list1.Add(obj1);
                    }
                }

                if (dt2.Rows.Count > 0)
                {
                    list2 = new List<Hr_Rule_htgw>();
                    for (int i = 0; i < dt2.Rows.Count; i++)
                    {
                        var obj2 = new Hr_Rule_htgw();
                        obj2.id = dt2.Rows[i]["id"].ToString();
                        obj2.deptCode = dt2.Rows[i]["deptCode"].ToString();
                        obj2.postCode = dt2.Rows[i]["postCode"].ToString();
                        obj2.coreNum = int.Parse(dt2.Rows[i]["coreNum"].ToString());
                        obj2.boneNum = int.Parse(dt2.Rows[i]["boneNum"].ToString());

                        list2.Add(obj2);
                    }
                }

                if (dt3.Rows.Count > 0)
                {
                    list3 = new List<Hr_Rule_htfd>();
                    for (int i = 0; i < dt3.Rows.Count; i++)
                    {
                        var obj3 = new Hr_Rule_htfd();
                        obj3.rid = dt3.Rows[i]["rid"].ToString();
                        obj3.gwid = dt3.Rows[i]["gwid"].ToString();
                        obj3.floatNum = int.Parse(dt3.Rows[i]["floatNum"].ToString());

                        list3.Add(obj3);
                    }
                }
                #endregion

                htType = "综合部";

                #region//岗位配备规则-综合部
                List<Hr_Rule_hthead> list4 = null; // 合同规则
                List<Hr_Rule_htgw> list5 = null;   // 岗位配备
                List<Hr_Rule_htfd> list6 = null;   // 岗位配备人数


                var dt4 = dal.GetHr_Rule_hthead(yearly, htType, midStr);// 合同规则
                var dt5 = dal.GetHr_Rule_htgw(yearly, htType, midStr);  // 岗位配备
                var dt6 = dal.GetHr_Rule_htfd(yearly, htType, midStr);  // 岗位配备人数
                if (dt4.Rows.Count > 0)
                {
                    list4 = new List<Hr_Rule_hthead>();
                    for (int i = 0; i < dt4.Rows.Count; i++)
                    {
                        var obj4 = new Hr_Rule_hthead();
                        obj4.id = dt4.Rows[i]["id"].ToString();
                        obj4.htType = dt4.Rows[i]["htType"].ToString();
                        obj4.htAmount = double.Parse(dt4.Rows[i]["htAmount"].ToString());

                        list4.Add(obj4);
                    }
                }

                if (dt5.Rows.Count > 0)
                {
                    list5 = new List<Hr_Rule_htgw>();
                    for (int i = 0; i < dt5.Rows.Count; i++)
                    {
                        var obj5 = new Hr_Rule_htgw();
                        obj5.id = dt5.Rows[i]["id"].ToString();
                        obj5.deptCode = dt5.Rows[i]["deptCode"].ToString();
                        obj5.postCode = dt5.Rows[i]["postCode"].ToString();
                        obj5.coreNum = int.Parse(dt5.Rows[i]["coreNum"].ToString());
                        obj5.boneNum = int.Parse(dt5.Rows[i]["boneNum"].ToString());

                        list5.Add(obj5);
                    }
                }

                if (dt6.Rows.Count > 0)
                {
                    list6 = new List<Hr_Rule_htfd>();
                    for (int i = 0; i < dt6.Rows.Count; i++)
                    {
                        var obj6 = new Hr_Rule_htfd();
                        obj6.rid = dt6.Rows[i]["rid"].ToString();
                        obj6.gwid = dt6.Rows[i]["gwid"].ToString();
                        obj6.floatNum = int.Parse(dt6.Rows[i]["floatNum"].ToString());

                        list6.Add(obj6);
                    }
                }
                #endregion

                //同步数据
                var list = new List<Hr_Midhtrs>();

                #region//单个项目组人数
                foreach (var item in list2)
                {
                    var _obj1 = new Hr_Midhtrs();
                    _obj1.easComCode = list_crm[0].easComCode;
                    _obj1.ruleDeptCode = item.deptCode;
                    _obj1.rulePostCode = item.postCode;

                    /*
                       * 根据 规划预算剩余合同额 匹配 合同规则配备表里的 合同额区间 合同额在a与b之间 取b的人数
                       * 根据 调整预算剩余合同额 匹配 合同规则配备表里的 合同额区间 合同额在a与b之间 取b的人数
                       * 超过区间最大值，报错异常
                       */

                    double ghsyAmount = list_crm.Where(p => p.ruleDeptCode == item.deptCode).First().ghsyAmount;
                    double tzsyAmount = list_crm.Where(p => p.ruleDeptCode == item.deptCode).First().tzsyAmount;

                    //筛选出 人数 对应的规则id
                    string ghsy_rid = "";
                    string tzsy_rid = "";
                    foreach (var item2 in list1)
                    {
                        if (ghsyAmount <= item2.htAmount && ghsy_rid == "")
                        {
                            ghsy_rid = item2.id;
                        }
                        if (tzsyAmount <= item2.htAmount && tzsy_rid == "")
                        {
                            tzsy_rid = item2.id;
                        }
                    }

                    //根据 规则id和岗位配备id 取人数
                    if (item.coreNum == 1)
                    {
                        _obj1.core_ghys = list3.Where(p => p.rid == ghsy_rid && p.gwid == item.id).First().floatNum; //规划预算核心人数
                        _obj1.core_tzys = list3.Where(p => p.rid == tzsy_rid && p.gwid == item.id).First().floatNum; //调整预算核心人数

                        _obj1.bone_ghys = 0; //规划预算骨干人数
                        _obj1.bone_tzys = 0; //调整预算骨干人数
                    }
                    else
                    {
                        _obj1.core_ghys = 0; //规划预算核心人数
                        _obj1.core_tzys = 0; //调整预算核心人数
                        _obj1.bone_ghys = list3.Where(p => p.rid == ghsy_rid && p.gwid == item.id).First().floatNum; //规划预算骨干人数
                        _obj1.bone_tzys = list3.Where(p => p.rid == tzsy_rid && p.gwid == item.id).First().floatNum; //调整预算骨干人数
                    }

                    list.Add(_obj1);
                }
                #endregion

                #region//综合部人数
                double ghsyAll = list_crm.Sum(p => p.ghsyAmount); //规划预算剩余合同额之和
                double tzsyAll = list_crm.Sum(p => p.tzsyAmount); //调整预算剩余合同额之和

                foreach (var item5 in list5)
                {
                    var _obj2 = new Hr_Midhtrs();

                    _obj2.easComCode = list_crm[0].easComCode;
                    _obj2.ruleDeptCode = item5.deptCode;
                    _obj2.rulePostCode = item5.postCode;

                    /*
                       * 根据 规划预算剩余合同额之和 匹配 合同规则配备表里的 合同额区间 合同额在a与b之间 取b的人数
                       * 根据 调整预算剩余合同额之和 匹配 合同规则配备表里的 合同额区间 合同额在a与b之间 取b的人数
                       * 超过区间最大值，报错异常
                       */

                    //筛选出 人数 对应的规则id
                    string ghsy_rid = "";
                    string tzsy_rid = "";
                    foreach (var item4 in list4)
                    {
                        if (ghsyAll <= item4.htAmount && ghsy_rid == "")
                        {
                            ghsy_rid = item4.id;
                        }
                        if (tzsyAll <= item4.htAmount && tzsy_rid == "")
                        {
                            tzsy_rid = item4.id;
                        }
                    }

                    //根据 规则id和岗位配备id 取人数
                    if (item5.coreNum == 1)
                    {
                        _obj2.core_ghys = list6.Where(p => p.rid == ghsy_rid && p.gwid == item5.id).First().floatNum; //规划预算核心人数
                        _obj2.core_tzys = list6.Where(p => p.rid == tzsy_rid && p.gwid == item5.id).First().floatNum; //调整预算核心人数

                        _obj2.bone_ghys = 0; //规划预算骨干人数
                        _obj2.bone_tzys = 0; //调整预算骨干人数
                    }
                    else
                    {
                        _obj2.core_ghys = 0; //规划预算核心人数
                        _obj2.core_tzys = 0; //调整预算核心人数
                        _obj2.bone_ghys = list6.Where(p => p.rid == ghsy_rid && p.gwid == item5.id).First().floatNum; //规划预算骨干人数
                        _obj2.bone_tzys = list6.Where(p => p.rid == tzsy_rid && p.gwid == item5.id).First().floatNum; //调整预算骨干人数
                    }

                    list.Add(_obj2);
                }
                #endregion

                #region//保存数据
                foreach (var item in list)
                {
                    var condition = new Dictionary<string, string>();
                    condition.Add("easComCode", item.easComCode); //公司 eas编码
                    condition.Add("ruleDeptCode", item.ruleDeptCode);
                    condition.Add("rulePostCode", item.rulePostCode);
                    condition.Add("core_ghys", item.core_ghys.ToString());
                    condition.Add("core_tzys", item.core_tzys.ToString());
                    condition.Add("bone_ghys", item.bone_ghys.ToString());
                    condition.Add("bone_tzys", item.bone_tzys.ToString());
                    condition.Add("yearly", yearly.ToString());
                    condition.Add("monthly", monthly.ToString());
                    condition.Add("mid", midStr);
                    condition.Add("addPer", name);

                    dal.AddHr_Midhtrs(condition);
                }
                #endregion

                result = "同步成功！";
                return result;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        #endregion

        #region //人工费

        //人工费-收入
        [MyAuthAttribute]
        public ActionResult Hr_Midrgsr()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //人工费-收入 数据
        public string GetHr_Midrgsr()
        {
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                JArray array = new JArray();  //返回结果

                //比较开始日期，结束日期
                var cc = DateTime.Compare(start, end);
                if (cc > 0)
                {
                    return array.ToString();  //开始日期大于结束日期,返回空
                }

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Midrgsr(start, end, mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("comName", dt.Rows[i]["comName"].ToString());
                        _json.Add("costType", dt.Rows[i]["costType"].ToString());
                        _json.Add("planBudget", dt.Rows[i]["planBudget"].ToString());
                        _json.Add("adjustBudget", dt.Rows[i]["adjustBudget"].ToString());
                        _json.Add("proBudget", dt.Rows[i]["proBudget"].ToString());
                        _json.Add("quotaLabor", dt.Rows[i]["quotaLabor"].ToString());
                        _json.Add("proportion", dt.Rows[i]["proportion"].ToString() + "%");
                        _json.Add("yearly", dt.Rows[i]["yearly"].ToString());
                        _json.Add("monthly", dt.Rows[i]["monthly"].ToString());
                        _json.Add("mid", dt.Rows[i]["mid"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出列表
        public void ExportHr_Midrgsr()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("人工费-收入", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("人工费-收入"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("公 司");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("年 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("月 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("费用类别");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("规划预算");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("调整预算");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("项目进展");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("实际");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("比例(%)");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Midrgsr(start, end, mid);
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["comName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yearly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["monthly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["costType"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["planBudget"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["adjustBudget"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["proBudget"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["quotaLabor"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["proportion"].ToString() + "%");
                        dataRow.GetCell(j).CellStyle = style1;

                    }
                }

                sheet.AutoSizeColumn(0);  //自适应宽度
                for (int k = 1; k < colCount; k++)
                {
                    sheet.SetColumnWidth(k, 15 * 256); //设置列宽
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }

        //删除
        public string DelHr_Midrgsr()
        {
            try
            {
                string ids = Request["ids"];

                var idList = ids.Split(',');

                var dal = new dalPro();
                foreach (var cc in idList)
                {
                    int _id = int.Parse(cc);
                    dal.DelHr_MidrgsrById(_id);
                }

                return "删除成功！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //// 同步数据
        public string SynHr_Midrgsr()
        {
            try
            {
                int yearly = int.Parse(Request["yearly"]);
                int monthly = int.Parse(Request["monthly"]);

                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string name = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var dal = new dalPro();
                string result = "";

                var dt = dal.GetHr_Bais_rgsrByMonth(yearly, monthly, midStr);
                if (dt.Rows.Count > 0)
                {
                    var condition = new Dictionary<string, string>();
                    condition.Add("yearly", yearly.ToString());
                    condition.Add("monthly", monthly.ToString());
                    condition.Add("easComCode", dt.Rows[0]["easCode"].ToString()); //公司 eas编码
                    condition.Add("costType", dt.Rows[0]["costType"].ToString());
                    condition.Add("planBudget", dt.Rows[0]["planBudget"].ToString());
                    condition.Add("adjustBudget", dt.Rows[0]["adjustBudget"].ToString());
                    condition.Add("proBudget", dt.Rows[0]["proBudget"].ToString());
                    condition.Add("quotaLabor", dt.Rows[0]["quotaLabor"].ToString());
                    condition.Add("proportion", dt.Rows[0]["proportion"].ToString());
                    condition.Add("mid", midStr);
                    condition.Add("addPer", name);

                    int i = dal.AddHr_Midrgsr(condition);
                    if (i > 0)
                    {
                        result = "同步成功！";
                    }
                    else
                    {
                        result = "同步失败！";
                    }
                }
                else
                {
                    result = "无Bais人工费-收入数据源！";
                }

                return result;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }



        //人工费-支出
        [MyAuthAttribute]
        public ActionResult Hr_Midrgzc()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //人工费-支出 数据
        public string GetHr_Midrgzc()
        {
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                JArray array = new JArray();  //返回结果

                //比较开始日期，结束日期
                var cc = DateTime.Compare(start, end);
                if (cc > 0)
                {
                    return array.ToString();  //开始日期大于结束日期,返回空
                }

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Midrgzc(start, end, mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("comName", dt.Rows[i]["comName"].ToString());
                        _json.Add("deptName", dt.Rows[i]["deptName"].ToString());
                        _json.Add("postName", dt.Rows[i]["postName"].ToString());
                        _json.Add("costType", dt.Rows[i]["costType"].ToString());
                        _json.Add("planBudget", dt.Rows[i]["planBudget"].ToString());
                        _json.Add("adjustBudget", dt.Rows[i]["adjustBudget"].ToString());
                        _json.Add("proBudget", dt.Rows[i]["proBudget"].ToString());
                        _json.Add("yearly", dt.Rows[i]["yearly"].ToString());
                        _json.Add("monthly", dt.Rows[i]["monthly"].ToString());
                        _json.Add("mid", dt.Rows[i]["mid"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出列表
        public void ExportHr_Midrgzc()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("人工费-支出", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("人工费-支出"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("公 司");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("部 门");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("岗 位");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("年 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("月 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("费用类别");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("规划预算");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("调整预算");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("项目进展");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Midrgzc(start, end, mid);
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["comName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["deptName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["postName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yearly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["monthly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["costType"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["planBudget"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["adjustBudget"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["proBudget"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;

                    }
                }

                sheet.AutoSizeColumn(0);  //自适应宽度
                for (int k = 1; k < colCount; k++)
                {
                    sheet.SetColumnWidth(k, 15 * 256); //设置列宽
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }

        //删除
        public string DelHr_Midrgzc()
        {
            try
            {
                string ids = Request["ids"];

                var idList = ids.Split(',');

                var dal = new dalPro();
                foreach (var cc in idList)
                {
                    int _id = int.Parse(cc);
                    dal.DelHr_MidrgzcById(_id);
                }

                return "删除成功！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        ////同步数据
        public string SynHr_Midrgzc()
        {
            try
            {
                int yearly = int.Parse(Request["yearly"]);
                int monthly = int.Parse(Request["monthly"]);

                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string name = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var dal = new dalPro();
                string result = "";

                #region//非市场人数
                List<Hr_Midysrs> list1 = null;
                var dt1 = dal.GetHr_MidysrsByMonth(yearly, monthly, midStr);
                if (dt1.Rows.Count == 0)
                {
                    result = "该月度无非市场人数数据！";
                    return result;
                }
                else
                {
                    list1 = new List<Hr_Midysrs>();
                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        var obj1 = new Hr_Midysrs();

                        obj1.easComCode = dt1.Rows[i]["easComCode"].ToString();
                        obj1.ruleDeptCode = dt1.Rows[i]["ruleDeptCode"].ToString();
                        obj1.rulePostCode = dt1.Rows[i]["rulePostCode"].ToString();
                        obj1.coreQuota = int.Parse(dt1.Rows[i]["coreQuota"].ToString());
                        obj1.boneQuota = int.Parse(dt1.Rows[i]["boneQuota"].ToString());
                        obj1.floatQuota = int.Parse(dt1.Rows[i]["floatQuota"].ToString());
                        obj1.floatghys = int.Parse(dt1.Rows[i]["floatghys"].ToString());
                        obj1.floattzys = int.Parse(dt1.Rows[i]["floattzys"].ToString());
                        obj1.coreActual = int.Parse(dt1.Rows[i]["coreActual"].ToString());
                        obj1.boneActual = int.Parse(dt1.Rows[i]["boneActual"].ToString());
                        obj1.floatActual = int.Parse(dt1.Rows[i]["floatActual"].ToString());

                        list1.Add(obj1);
                    }
                }
                #endregion

                #region//市场人数
                List<Hr_Midhtrs> list2 = null;
                var dt2 = dal.GetHr_MidhtrsByMonth(yearly, monthly, midStr);
                if (dt2.Rows.Count == 0)
                {
                    result = "该月度无市场人数数据！";
                    return result;
                }
                else
                {
                    list2 = new List<Hr_Midhtrs>();
                    for (int i = 0; i < dt2.Rows.Count; i++)
                    {
                        var obj2 = new Hr_Midhtrs();

                        obj2.easComCode = dt2.Rows[i]["easComCode"].ToString();
                        obj2.ruleDeptCode = dt2.Rows[i]["ruleDeptCode"].ToString();
                        obj2.rulePostCode = dt2.Rows[i]["rulePostCode"].ToString();
                        obj2.core_ghys = int.Parse(dt2.Rows[i]["core_ghys"].ToString());
                        obj2.bone_ghys = int.Parse(dt2.Rows[i]["bone_ghys"].ToString());
                        obj2.core_tzys = int.Parse(dt2.Rows[i]["core_tzys"].ToString());
                        obj2.bone_tzys = int.Parse(dt2.Rows[i]["bone_tzys"].ToString());

                        list2.Add(obj2);
                    }
                }
                #endregion

                #region//实际人数及工资
                var list3 = new List<Hr_Bhr_fact>();
                var dt3 = dal.GetHr_Bhr_factByMonth2(yearly, monthly, midStr);
                if (dt3.Rows.Count > 0)
                {
                    for (int i = 0; i < dt3.Rows.Count; i++)
                    {
                        var obj3 = new Hr_Bhr_fact();

                        obj3.easComCode = dt3.Rows[i]["easComCode"].ToString();
                        obj3.ruleDeptCode = dt3.Rows[i]["ruleDeptCode"].ToString();
                        obj3.rulePostCode = dt3.Rows[i]["rulePostCode"].ToString();
                        obj3.postType = dt3.Rows[i]["postType"].ToString();
                        obj3.wage = double.Parse(dt3.Rows[i]["wage"].ToString());

                        list3.Add(obj3);
                    }
                }
                #endregion

                #region//营收规则 岗位配备
                List<Hr_Rule_ysgw> list4 = null;
                var dt4 = dal.GetHr_Rule_ysgw(yearly, midStr);
                if (dt4.Rows.Count == 0)
                {
                    result = "无营收岗位配备规则数据！";
                    return result;
                }
                else
                {
                    list4 = new List<Hr_Rule_ysgw>();
                    for (int i = 0; i < dt4.Rows.Count; i++)
                    {
                        var obj4 = new Hr_Rule_ysgw();

                        obj4.deptCode = dt4.Rows[i]["deptCode"].ToString();
                        obj4.postCode = dt4.Rows[i]["postCode"].ToString();
                        obj4.costType = dt4.Rows[i]["costType"].ToString();
                        obj4.quotaWage = double.Parse(dt4.Rows[i]["quotaWage"].ToString());

                        list4.Add(obj4);
                    }
                }
                #endregion
                
                #region//合同规则 岗位配备
                List<Hr_Rule_htgw> list5 = null;
                var dt5 = dal.GetHr_Rule_htgw(yearly, "", midStr);
                if (dt5.Rows.Count == 0)
                {
                    result = "无合同岗位配备规则数据！";
                    return result;
                }
                else
                {
                    list5 = new List<Hr_Rule_htgw>();
                    for (int i = 0; i < dt5.Rows.Count; i++)
                    {
                        var obj5 = new Hr_Rule_htgw();

                        obj5.deptCode = dt5.Rows[i]["deptCode"].ToString();
                        obj5.postCode = dt5.Rows[i]["postCode"].ToString();
                        obj5.costType = dt5.Rows[i]["costType"].ToString();
                        obj5.quotaWage = double.Parse(dt5.Rows[i]["quotaWage"].ToString());

                        list5.Add(obj5);
                    }
                }
                #endregion


                //同步数据
                var list = new List<Hr_Midrgzc>();

                #region//非市场
                foreach (var item4 in list4)
                {
                    var _obj1 = new Hr_Midrgzc();

                    _obj1.easComCode = list1[0].easComCode;
                    _obj1.ruleDeptCode = item4.deptCode;
                    _obj1.rulePostCode = item4.postCode;
                    _obj1.costType = item4.costType;

                    _obj1.yearly = yearly;
                    _obj1.monthly = monthly;

                    /*定额人数(A) 与 实际人数(B) 比对取工资(W)
                        * 
                        * A = 0, W = 0
                        * B = 0, 工资取岗位配备表定额工资(w1) W = A * w1；
                        * B = 0, 岗位配备表没值，W = 0；（不存在）
                        * B > 0, B > A 工资取岗位实际工资平均工资(w2) W = A * w2；
                        * B > 0, B = A 工资取岗位实际工资值(w3)和  W = w3 和值；
                        * B > 0, B < A A其中的B个人取实际工资值，剩余 A - B 个人从实际工资值由高到低依次匹配取值，W = 最后求和；
                        * 
                        */

                    //岗位人数
                    int coreQuota = 0; //核心定额
                    int boneQuota = 0; //骨干定额
                    int floatQuota = 0;//浮动定额
                    int floatghys = 0; //浮动规划预算
                    int floattzys = 0; //浮动调整预算
                    int coreActual = 0; //核心实际
                    int boneActual = 0; //骨干实际
                    int floatActual = 0;//浮动实际

                    var _list1 = list1.Where(p => p.ruleDeptCode == item4.deptCode && p.rulePostCode == item4.postCode).ToList();
                    if (_list1.Count() > 0)
                    {
                        coreQuota = _list1[0].coreQuota;
                        boneQuota = _list1[0].boneQuota;
                        floatQuota = _list1[0].floatQuota;
                        floatghys = _list1[0].floatghys;
                        floattzys = _list1[0].floattzys;
                        coreActual = _list1[0].coreActual;
                        boneActual = _list1[0].boneActual;
                        floatActual = _list1[0].floatActual;
                    }

                    //定额工资
                    double _quotaWage = item4.quotaWage;

                    //实际工资 列表
                    var _list3_core = list3.Where(p => p.ruleDeptCode == item4.deptCode && p.rulePostCode == item4.postCode && p.postType == "核心").ToList();
                    var _list3_bone = list3.Where(p => p.ruleDeptCode == item4.deptCode && p.rulePostCode == item4.postCode && p.postType == "骨干").ToList();
                    var _list3_float = list3.Where(p => p.ruleDeptCode == item4.deptCode && p.rulePostCode == item4.postCode && p.postType == "浮动").ToList();

                    
                    double coreWage = SetWage(coreQuota, coreActual, _list3_core, _quotaWage);  //核心人数总工资                    
                    double boneWage = SetWage(boneQuota, boneActual, _list3_bone, _quotaWage);  //骨干人数总工资                    
                    double floatWage = SetWage(floatQuota, floatActual, _list3_float, _quotaWage);//浮动人数总工资                    
                    double ghysWage = SetWage(floatghys, floatActual, _list3_float, _quotaWage);//浮动 规划预算人数总工资
                    double tzysWage = SetWage(floattzys, floatActual, _list3_float, _quotaWage);//浮动 调整预算人数总工资

                    _obj1.planBudget = coreWage + boneWage + ghysWage; //规划预算 = 核心 + 骨干 + 规划预算（浮动）
                    _obj1.adjustBudget = coreWage + boneWage + tzysWage; //调整预算 = 核心 + 骨干 + 调整预算（浮动）
                    _obj1.proBudget = coreWage + boneWage + floatWage; //项目进展 = 核心 + 骨干 + 浮动

                    list.Add(_obj1);
                }
                #endregion

                #region//市场
                foreach (var item5 in list5)
                {
                    var _obj2 = new Hr_Midrgzc();

                    _obj2.easComCode = list2[0].easComCode;
                    _obj2.ruleDeptCode = item5.deptCode;
                    _obj2.rulePostCode = item5.postCode;
                    _obj2.costType = item5.costType;

                    _obj2.yearly = yearly;
                    _obj2.monthly = monthly;

                    /*定额人数(A) 与 实际人数(B) 比对取工资(W)
                        * 
                        * A = 0, W = 0
                        * B = 0, 工资取岗位配备表定额工资(w1) W = A * w1；
                        * B = 0, 岗位配备表没值，W = 0；（不存在）
                        * B > 0, B > A 工资取岗位实际工资平均工资(w2) W = A * w2；
                        * B > 0, B = A 工资取岗位实际工资值(w3)和  W = w3 和值；
                        * B > 0, B < A A其中的B个人取实际工资值，剩余 A - B 个人从实际工资值由高到低依次匹配取值，W = 最后求和；
                        * 
                        */

                    //岗位定额人数
                    int core_ghys = 0; //核心 规划预算
                    int bone_ghys = 0; //骨干 规划预算
                    int core_tzys = 0; //核心 调整预算
                    int bone_tzys = 0; //骨干 调整预算

                    var _list2 = list2.Where(p => p.ruleDeptCode == item5.deptCode && p.rulePostCode == item5.postCode).ToList();
                    if (_list2.Count() > 0)
                    {
                        core_ghys = _list2[0].core_ghys;
                        bone_ghys = _list2[0].bone_ghys;
                        core_tzys = _list2[0].core_tzys;
                        bone_tzys = _list2[0].bone_tzys;
                    }

                    //定额工资
                    double _quotaWage = item5.quotaWage;

                    //实际工资 列表
                    var _list3_core = list3.Where(p => p.ruleDeptCode == item5.deptCode && p.rulePostCode == item5.postCode && p.postType == "核心").ToList();
                    var _list3_bone = list3.Where(p => p.ruleDeptCode == item5.deptCode && p.rulePostCode == item5.postCode && p.postType == "骨干").ToList();

                    //实际人数
                    int coreActual = _list3_core.Count(); //核心实际
                    int boneActual = _list3_bone.Count(); //骨干实际

                    double coreWage = SetWage(core_ghys, coreActual, _list3_core, _quotaWage);  //规划预算 核心人数总工资
                    double boneWage = SetWage(bone_ghys, boneActual, _list3_bone, _quotaWage);  //规划预算 骨干人数总工资
                    double coreWage2 = SetWage(core_tzys, coreActual, _list3_core, _quotaWage); //调整预算 核心人数总工资
                    double boneWage2 = SetWage(bone_tzys, boneActual, _list3_bone, _quotaWage); //调整预算 骨干人数总工资

                    _obj2.planBudget = coreWage + boneWage; //规划预算 核心 + 骨干
                    _obj2.adjustBudget = coreWage2 + boneWage2; //调整预算 核心 + 骨干
                    _obj2.proBudget = coreWage2 + boneWage2; //调整预算 核心 + 骨干

                    list.Add(_obj2);
                }
                #endregion

                #region//保存数据
                foreach (var item in list)
                {
                    var condition = new Dictionary<string, string>();
                    condition.Add("easComCode", item.easComCode); //公司 eas编码
                    condition.Add("ruleDeptCode", item.ruleDeptCode);
                    condition.Add("rulePostCode", item.rulePostCode);
                    condition.Add("costType", item.costType);
                    condition.Add("planBudget", item.planBudget.ToString());
                    condition.Add("adjustBudget", item.adjustBudget.ToString());
                    condition.Add("proBudget", item.proBudget.ToString());
                    condition.Add("yearly", yearly.ToString());
                    condition.Add("monthly", monthly.ToString());
                    condition.Add("mid", midStr);
                    condition.Add("addPer", name);

                    dal.AddHr_Midrgzc(condition);
                }
                #endregion

                result = "同步成功！";
                return result;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //// 计算预算工资
        public double SetWage(int quotaNum, int actualNum, List<Hr_Bhr_fact> list, double quotaWage)
        {
            double result = 0;
            if (quotaNum > 0)  //定额人数大于0
            {
                if (actualNum > 0)  //实际人数大于0
                {
                    if (actualNum > quotaNum) //实际人数大于定额人数 
                    {
                        double avg = list.Average(p => p.wage); // 实际工资平均值
                        result = avg * quotaNum; //实际平均工资*定额人数
                    }
                    else if (actualNum == quotaNum) //实际人数等于定额人数
                    {
                        result = list.Sum(p => p.wage); //实际工资总和
                    }
                    else  //实际人数小于定额人数
                    {
                        list = list.OrderByDescending(p => p.wage).ToList(); //从高到低排序

                        while (quotaNum > 0)
                        {
                            for (int i = 0; i < actualNum; i++)
                            {
                                result += list[i].wage; //按定额人数，实际工资从高到低循环取值，最后求和

                                quotaNum--;
                                if (quotaNum == 0)
                                {
                                    break;
                                }
                            }
                        }
                    }
                }
                else
                {
                    result = quotaWage * quotaNum; //定额工资*定额人数
                }
            }

            return result;
        }



        //人工费-支出实际
        [MyAuthAttribute]
        public ActionResult Hr_Midrgsj()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //人工费-支出实际 数据
        public string GetHr_Midrgsj()
        {
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                JArray array = new JArray();  //返回结果

                //比较开始日期，结束日期
                var cc = DateTime.Compare(start, end);
                if (cc > 0)
                {
                    return array.ToString();  //开始日期大于结束日期,返回空
                }

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Midrgsj(start, end, mid);

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var _json = new JObject();

                        _json.Add("id", dt.Rows[i]["id"].ToString());
                        _json.Add("comName", dt.Rows[i]["comName"].ToString());
                        _json.Add("costType", dt.Rows[i]["costType"].ToString());
                        _json.Add("quotaLabor", dt.Rows[i]["quotaLabor"].ToString());
                        _json.Add("yearly", dt.Rows[i]["yearly"].ToString());
                        _json.Add("monthly", dt.Rows[i]["monthly"].ToString());
                        _json.Add("mid", dt.Rows[i]["mid"].ToString());

                        array.Add(_json);
                    }
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //导出列表
        public void ExportHr_Midrgsj()
        {
            var attachment = "attachment; filename=" + HttpUtility.UrlEncode("人工费-支出(实际)", System.Text.Encoding.UTF8) + DateTime.Now.ToString("yyyyMMdd") + ".xls";
            Response.Clear();
            Response.ContentEncoding = Encoding.UTF8;
            Response.AddHeader("content-disposition", attachment);
            Response.ContentType = "application/ms-excel";

            IWorkbook workbook = new HSSFWorkbook();//创建工作簿
            var sheet = workbook.CreateSheet("人工费-收入(实际)"); //创建工作表

            #region//样式
            ICellStyle style = workbook.CreateCellStyle();//样式
            style.Alignment = HorizontalAlignment.Center;//水平居中
            style.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            style.BorderBottom = BorderStyle.Thin; //边框
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            //背景颜色
            style.FillForegroundColor = 31;
            style.FillPattern = FillPattern.SolidForeground;

            //样式1
            var style1 = workbook.CreateCellStyle();
            style1.Alignment = HorizontalAlignment.Center;
            style1.VerticalAlignment = VerticalAlignment.Center;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            #endregion

            #region//表头
            int colCount = 0;

            var headRow0 = sheet.CreateRow(0);
            headRow0.Height = 20 * 30;//高度
            headRow0.CreateCell(colCount).SetCellValue("公 司");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("年 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("月 度");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("费用类别");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;
            headRow0.CreateCell(colCount).SetCellValue("实际");
            headRow0.GetCell(colCount).CellStyle = style;
            colCount++;

            #endregion

            #region //表体
            try
            {
                string startDate = Request["startDate"];
                string endDate = Request["endDate"];

                DateTime start = DateTime.Parse(startDate);  //开始日期
                DateTime end = DateTime.Parse(endDate);  //结束日期

                //组织单元集合(有权限)
                var cookie = GetCookie();
                int mid = int.Parse(cookie["comId"]);

                var dal = new dalPro();
                var dt = dal.GetHr_Midrgsj(start, end, mid);
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var dataRow = sheet.CreateRow(i + 1);
                        dataRow.Height = 20 * 25;

                        int j = 0; //列数
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["comName"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["yearly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["monthly"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["costType"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;
                        j++;
                        dataRow.CreateCell(j).SetCellValue(dt.Rows[i]["quotaLabor"].ToString());
                        dataRow.GetCell(j).CellStyle = style1;

                    }
                }

                sheet.AutoSizeColumn(0);  //自适应宽度
                for (int k = 1; k < colCount; k++)
                {
                    sheet.SetColumnWidth(k, 15 * 256); //设置列宽
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            #endregion

            workbook.Write(Response.OutputStream);
            Response.End();
        }

        //删除
        public string DelHr_Midrgsj()
        {
            try
            {
                string ids = Request["ids"];

                var idList = ids.Split(',');

                var dal = new dalPro();
                foreach (var cc in idList)
                {
                    int _id = int.Parse(cc);
                    dal.DelHr_MidrgsjById(_id);
                }

                return "删除成功！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //// 同步数据
        public string SynHr_Midrgsj()
        {
            try
            {
                int yearly = int.Parse(Request["yearly"]);
                int monthly = int.Parse(Request["monthly"]);

                var cookie = GetCookie();
                string midStr = cookie["comId"];
                string name = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));

                var dal = new dalPro();
                string result = "";

                var dt = dal.GetHr_Bais_rgzcByMonth(yearly, monthly, midStr);
                if (dt.Rows.Count > 0)
                {
                    var condition = new Dictionary<string, string>();
                    condition.Add("yearly", yearly.ToString());
                    condition.Add("monthly", monthly.ToString());
                    condition.Add("easComCode", dt.Rows[0]["easCode"].ToString()); //公司 eas编码
                    condition.Add("costType", dt.Rows[0]["costType"].ToString());
                    condition.Add("quotaLabor", dt.Rows[0]["quotaLabor"].ToString());
                    condition.Add("mid", midStr);
                    condition.Add("addPer", name);

                    int i = dal.AddHr_Midrgsj(condition);
                    if (i > 0)
                    {
                        result = "同步成功！";
                    }
                    else
                    {
                        result = "同步失败！";
                    }
                }
                else
                {
                    result = "无Bais人工费-支出数据源！";
                }

                return result;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        #endregion




        //// HRM调整预算
        [MyAuthAttribute]
        public ActionResult Hr_Model()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }


        //获取HRM调整预算数据
        public string GetHr_Model()
        {
            try
            {
                int yearly = int.Parse(Request["yearly"]);

                List<Model> list = SetModelData(yearly);  // 获取数据

                JArray array = new JArray();
                int count = 0;
                foreach (var item in list)
                {
                    count++;

                    var json = new JObject();

                    json.Add("id", count);
                    json.Add("classify", item.classify);
                    json.Add("goal", item.goal);
                    json.Add("month1", item.month1);
                    json.Add("month2", item.month2);
                    json.Add("month3", item.month3);
                    json.Add("month4", item.month4);
                    json.Add("month5", item.month5);
                    json.Add("month6", item.month6);
                    json.Add("month7", item.month7);
                    json.Add("month8", item.month8);
                    json.Add("month9", item.month9);
                    json.Add("month10", item.month10);
                    json.Add("month11", item.month11);
                    json.Add("month12", item.month12);

                    array.Add(json);
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<Model> SetModelData(int yearly)
        {
            try
            {
                //组织单元
                var cookie = GetCookie();
                string midStr = cookie["comId"];
                int mid = int.Parse(midStr);

                DateTime start = DateTime.Parse(yearly + "-01-01");
                DateTime end = DateTime.Parse(yearly + "-12-31");

                var list = new List<Model>();

                #region//初始返回结果对象 共39条
                for (int i = 0; i < 39; i++)
                {
                    var obj = new Model();
                    obj.classify = "";  //项目/分类
                    obj.goal = "1";     //目标
                    obj.month1 = "1";
                    obj.month2 = "1";
                    obj.month3 = "1";
                    obj.month4 = "1";
                    obj.month5 = "1";
                    obj.month6 = "1";
                    obj.month7 = "1";
                    obj.month8 = "1";
                    obj.month9 = "1";
                    obj.month10 = "1";
                    obj.month11 = "1";
                    obj.month12 = "1";

                    list.Add(obj);
                }
                #endregion

                var dal = new dalPro();

                #region// 调整预算数据
                var list_tzys = new List<Hr_Midtzys>();

                var dt1 = dal.GetHr_Midtzys(start, end, mid);
                if (dt1.Rows.Count > 0)
                {
                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        var obj1 = new Hr_Midtzys();

                        obj1.htAmount = double.Parse(dt1.Rows[i]["htAmount"].ToString());
                        obj1.ysAmount = double.Parse(dt1.Rows[i]["ysAmount"].ToString());
                        obj1.lrAmount = double.Parse(dt1.Rows[i]["lrAmount"].ToString());
                        obj1.yield = double.Parse(dt1.Rows[i]["yield"].ToString());
                        obj1.yieEffic = double.Parse(dt1.Rows[i]["yieEffic"].ToString());
                        obj1.proTeams = int.Parse(dt1.Rows[i]["proTeams"].ToString());
                        obj1.monthly = int.Parse(dt1.Rows[i]["monthly"].ToString());

                        list_tzys.Add(obj1);
                    }
                }

                //分开取 1到12月 调整预算
                var list_tzys1 = list_tzys.Where(p => p.monthly == 1).ToList();
                var list_tzys2 = list_tzys.Where(p => p.monthly == 2).ToList();
                var list_tzys3 = list_tzys.Where(p => p.monthly == 3).ToList();
                var list_tzys4 = list_tzys.Where(p => p.monthly == 4).ToList();
                var list_tzys5 = list_tzys.Where(p => p.monthly == 5).ToList();
                var list_tzys6 = list_tzys.Where(p => p.monthly == 6).ToList();
                var list_tzys7 = list_tzys.Where(p => p.monthly == 7).ToList();
                var list_tzys8 = list_tzys.Where(p => p.monthly == 8).ToList();
                var list_tzys9 = list_tzys.Where(p => p.monthly == 9).ToList();
                var list_tzys10 = list_tzys.Where(p => p.monthly == 10).ToList();
                var list_tzys11 = list_tzys.Where(p => p.monthly == 11).ToList();
                var list_tzys12 = list_tzys.Where(p => p.monthly == 12).ToList();

                #endregion

                #region //序号2-5，6，13
                //序号2
                list[0].classify = "合同额(万)";
                list[0].goal = list_tzys.Sum(p => p.htAmount).ToString();
                list[0].month1 = list_tzys1[0].htAmount.ToString();
                list[0].month2 = list_tzys2[0].htAmount.ToString();
                list[0].month3 = list_tzys3[0].htAmount.ToString();
                list[0].month4 = list_tzys4[0].htAmount.ToString();
                list[0].month5 = list_tzys5[0].htAmount.ToString();
                list[0].month6 = list_tzys6[0].htAmount.ToString();
                list[0].month7 = list_tzys7[0].htAmount.ToString();
                list[0].month8 = list_tzys8[0].htAmount.ToString();
                list[0].month9 = list_tzys9[0].htAmount.ToString();
                list[0].month10 = list_tzys10[0].htAmount.ToString();
                list[0].month11 = list_tzys11[0].htAmount.ToString();
                list[0].month12 = list_tzys12[0].htAmount.ToString();

                //序号3
                list[1].classify = "营收(万)";
                list[1].goal = list_tzys.Sum(p => p.ysAmount).ToString();
                list[1].month1 = list_tzys1[0].ysAmount.ToString();
                list[1].month2 = list_tzys2[0].ysAmount.ToString();
                list[1].month3 = list_tzys3[0].ysAmount.ToString();
                list[1].month4 = list_tzys4[0].ysAmount.ToString();
                list[1].month5 = list_tzys5[0].ysAmount.ToString();
                list[1].month6 = list_tzys6[0].ysAmount.ToString();
                list[1].month7 = list_tzys7[0].ysAmount.ToString();
                list[1].month8 = list_tzys8[0].ysAmount.ToString();
                list[1].month9 = list_tzys9[0].ysAmount.ToString();
                list[1].month10 = list_tzys10[0].ysAmount.ToString();
                list[1].month11 = list_tzys11[0].ysAmount.ToString();
                list[1].month12 = list_tzys12[0].ysAmount.ToString();
                //序号4
                list[2].classify = "利润(万)";
                list[2].goal = list_tzys.Sum(p => p.lrAmount).ToString();
                list[2].month1 = list_tzys1[0].lrAmount.ToString();
                list[2].month2 = list_tzys2[0].lrAmount.ToString();
                list[2].month3 = list_tzys3[0].lrAmount.ToString();
                list[2].month4 = list_tzys4[0].lrAmount.ToString();
                list[2].month5 = list_tzys5[0].lrAmount.ToString();
                list[2].month6 = list_tzys6[0].lrAmount.ToString();
                list[2].month7 = list_tzys7[0].lrAmount.ToString();
                list[2].month8 = list_tzys8[0].lrAmount.ToString();
                list[2].month9 = list_tzys9[0].lrAmount.ToString();
                list[2].month10 = list_tzys10[0].lrAmount.ToString();
                list[2].month11 = list_tzys11[0].lrAmount.ToString();
                list[2].month12 = list_tzys12[0].lrAmount.ToString();
                //序号5
                list[3].classify = "产量(万立方)";
                list[3].goal = list_tzys.Sum(p => p.yield).ToString();
                list[3].month1 = list_tzys1[0].yield.ToString();
                list[3].month2 = list_tzys2[0].yield.ToString();
                list[3].month3 = list_tzys3[0].yield.ToString();
                list[3].month4 = list_tzys4[0].yield.ToString();
                list[3].month5 = list_tzys5[0].yield.ToString();
                list[3].month6 = list_tzys6[0].yield.ToString();
                list[3].month7 = list_tzys7[0].yield.ToString();
                list[3].month8 = list_tzys8[0].yield.ToString();
                list[3].month9 = list_tzys9[0].yield.ToString();
                list[3].month10 = list_tzys10[0].yield.ToString();
                list[3].month11 = list_tzys11[0].yield.ToString();
                list[3].month12 = list_tzys12[0].yield.ToString();


                //序号6
                list[4].classify = "综合产效（m³/8H/人）";
                list[4].goal = (list_tzys.Sum(p => p.yieEffic) / 12).ToString(); //取平均值
                list[4].month1 = list_tzys1[0].yield.ToString();
                list[4].month2 = list_tzys2[0].yield.ToString();
                list[4].month3 = list_tzys3[0].yield.ToString();
                list[4].month4 = list_tzys4[0].yield.ToString();
                list[4].month5 = list_tzys5[0].yield.ToString();
                list[4].month6 = list_tzys6[0].yield.ToString();
                list[4].month7 = list_tzys7[0].yield.ToString();
                list[4].month8 = list_tzys8[0].yield.ToString();
                list[4].month9 = list_tzys9[0].yield.ToString();
                list[4].month10 = list_tzys10[0].yield.ToString();
                list[4].month11 = list_tzys11[0].yield.ToString();
                list[4].month12 = list_tzys12[0].yield.ToString();


                //序号13
                list[11].classify = "销售团队数";
                list[11].goal = (list_tzys.Sum(p => p.proTeams) / 12).ToString();
                list[11].month1 = list_tzys1[0].proTeams.ToString();
                list[11].month2 = list_tzys2[0].proTeams.ToString();
                list[11].month3 = list_tzys3[0].proTeams.ToString();
                list[11].month4 = list_tzys4[0].proTeams.ToString();
                list[11].month5 = list_tzys5[0].proTeams.ToString();
                list[11].month6 = list_tzys6[0].proTeams.ToString();
                list[11].month7 = list_tzys7[0].proTeams.ToString();
                list[11].month8 = list_tzys8[0].proTeams.ToString();
                list[11].month9 = list_tzys9[0].proTeams.ToString();
                list[11].month10 = list_tzys10[0].proTeams.ToString();
                list[11].month11 = list_tzys11[0].proTeams.ToString();
                list[11].month12 = list_tzys12[0].proTeams.ToString();

                #endregion


                #region// 人数数据
                //市场人数
                var list_htrs = new List<Hr_Midhtrs>();

                var dt2 = dal.GetHr_Midhtrs(start, end, mid);
                if (dt2.Rows.Count > 0)
                {
                    for (int i = 0; i < dt2.Rows.Count; i++)
                    {
                        var obj2 = new Hr_Midhtrs();
                        
                        obj2.core_tzys = int.Parse(dt2.Rows[i]["core_tzys"].ToString());
                        obj2.bone_tzys = int.Parse(dt2.Rows[i]["bone_tzys"].ToString());
                        obj2.monthly = int.Parse(dt2.Rows[i]["monthly"].ToString());

                        list_htrs.Add(obj2);
                    }
                }
                
                //非市场人数
                var list_ysrs = new List<Hr_Midysrs>();

                var dt3 = dal.GetHr_Midysrs(start, end, mid);
                if (dt3.Rows.Count > 0)
                {
                    for (int i = 0; i < dt3.Rows.Count; i++)
                    {
                        var obj3 = new Hr_Midysrs();
                        
                        obj3.coreQuota = int.Parse(dt3.Rows[i]["coreQuota"].ToString());
                        obj3.boneQuota = int.Parse(dt3.Rows[i]["boneQuota"].ToString());
                        obj3.floattzys = int.Parse(dt3.Rows[i]["floattzys"].ToString());
                        obj3.monthly = int.Parse(dt3.Rows[i]["monthly"].ToString());
                        obj3.postName = dt3.Rows[i]["postName"].ToString();
                        obj3.postLevel= dt3.Rows[i]["postLevel"].ToString();

                        list_ysrs.Add(obj3);
                    }
                }
                #endregion
                
                #region//序号11，12，14—19
                //序号11
                /* a + b
                  * 市场人数表：取公司的 调整预算核心人数之和 a
                  * 非市场人数表：取公司 核心定额人数之和 b
                  * 年度目标：取平均值（四舍五入取整数）
                  */
                list[9].classify = "其中：核心";
                list[9].goal = (list_htrs.Sum(p => p.core_tzys) / 12 + list_ysrs.Sum(p => p.coreQuota) / 12).ToString();
                list[9].month1 = (list_htrs.Where(p => p.monthly == 1).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 1).Sum(p => p.coreQuota)).ToString();
                list[9].month2 = (list_htrs.Where(p => p.monthly == 2).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 2).Sum(p => p.coreQuota)).ToString();
                list[9].month3 = (list_htrs.Where(p => p.monthly == 3).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 3).Sum(p => p.coreQuota)).ToString();
                list[9].month4 = (list_htrs.Where(p => p.monthly == 4).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 4).Sum(p => p.coreQuota)).ToString();
                list[9].month5 = (list_htrs.Where(p => p.monthly == 5).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 5).Sum(p => p.coreQuota)).ToString();
                list[9].month6 = (list_htrs.Where(p => p.monthly == 6).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 6).Sum(p => p.coreQuota)).ToString();
                list[9].month7 = (list_htrs.Where(p => p.monthly == 7).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 7).Sum(p => p.coreQuota)).ToString();
                list[9].month8 = (list_htrs.Where(p => p.monthly == 8).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 8).Sum(p => p.coreQuota)).ToString();
                list[9].month9 = (list_htrs.Where(p => p.monthly == 9).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 9).Sum(p => p.coreQuota)).ToString();
                list[9].month10 = (list_htrs.Where(p => p.monthly == 10).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 10).Sum(p => p.coreQuota)).ToString();
                list[9].month11 = (list_htrs.Where(p => p.monthly == 11).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 11).Sum(p => p.coreQuota)).ToString();
                list[9].month12 = (list_htrs.Where(p => p.monthly == 12).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 12).Sum(p => p.coreQuota)).ToString();

                //序号12
                /* a + b
                  * 市场人数表：取公司的 调整预算骨干人数之和 a
                  * 非市场人数表：取公司 骨干定额人数之和 b
                  * 年度目标：取平均值（四舍五入取整数）
                  */
                list[10].classify = "骨干";
                list[10].goal = (list_htrs.Sum(p => p.bone_tzys) / 12 + list_ysrs.Sum(p => p.boneQuota) / 12).ToString();
                list[10].month1 = (list_htrs.Where(p => p.monthly == 1).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 1).Sum(p => p.boneQuota)).ToString();
                list[10].month2 = (list_htrs.Where(p => p.monthly == 2).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 2).Sum(p => p.boneQuota)).ToString();
                list[10].month3 = (list_htrs.Where(p => p.monthly == 3).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 3).Sum(p => p.boneQuota)).ToString();
                list[10].month4 = (list_htrs.Where(p => p.monthly == 4).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 4).Sum(p => p.boneQuota)).ToString();
                list[10].month5 = (list_htrs.Where(p => p.monthly == 5).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 5).Sum(p => p.boneQuota)).ToString();
                list[10].month6 = (list_htrs.Where(p => p.monthly == 6).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 6).Sum(p => p.boneQuota)).ToString();
                list[10].month7 = (list_htrs.Where(p => p.monthly == 7).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 7).Sum(p => p.boneQuota)).ToString();
                list[10].month8 = (list_htrs.Where(p => p.monthly == 8).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 8).Sum(p => p.boneQuota)).ToString();
                list[10].month9 = (list_htrs.Where(p => p.monthly == 9).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 9).Sum(p => p.boneQuota)).ToString();
                list[10].month10 = (list_htrs.Where(p => p.monthly == 10).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 10).Sum(p => p.boneQuota)).ToString();
                list[10].month11 = (list_htrs.Where(p => p.monthly == 11).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 11).Sum(p => p.boneQuota)).ToString();
                list[10].month12 = (list_htrs.Where(p => p.monthly == 12).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 12).Sum(p => p.boneQuota)).ToString();

                //序号14
                /* 
                  * 市场人数表：取公司的 调整预算（核心 + 骨干）人数之和
                  * 年度目标：取平均值（四舍五入取整数）
                  */
                list[12].classify = "其中：营销中心";
                list[12].goal = (list_htrs.Sum(p => p.core_tzys + p.bone_tzys) / 12).ToString();
                list[12].month1 = (list_htrs.Where(p => p.monthly == 1).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[12].month2 = (list_htrs.Where(p => p.monthly == 2).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[12].month3 = (list_htrs.Where(p => p.monthly == 3).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[12].month4 = (list_htrs.Where(p => p.monthly == 4).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[12].month5 = (list_htrs.Where(p => p.monthly == 5).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[12].month6 = (list_htrs.Where(p => p.monthly == 6).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[12].month7 = (list_htrs.Where(p => p.monthly == 7).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[12].month8 = (list_htrs.Where(p => p.monthly == 8).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[12].month9 = (list_htrs.Where(p => p.monthly == 9).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[12].month10 = (list_htrs.Where(p => p.monthly == 10).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[12].month11 = (list_htrs.Where(p => p.monthly == 11).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[12].month12 = (list_htrs.Where(p => p.monthly == 12).Sum(p => p.core_tzys + p.bone_tzys)).ToString();

                //序号15
                /* 
                  * 非市场人数表：取岗位职级为A—F 的 核心定额人数 + 骨干定额人数 + 浮动调整预算人数
                  * 年度目标：取平均值（四舍五入取整数）
                  */
                var _list_ysrs = list_ysrs.Where(p => p.postLevel == "A" || p.postLevel == "B" || p.postLevel == "C"
                    || p.postLevel == "D" || p.postLevel == "E" || p.postLevel == "F").ToList();

                list[13].classify = "干部";
                list[13].goal = (_list_ysrs.Sum(p => p.coreQuota + p.boneQuota + p.floattzys) / 12).ToString();
                list[13].month1 = (_list_ysrs.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[13].month2 = (_list_ysrs.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[13].month3 = (_list_ysrs.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[13].month4 = (_list_ysrs.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[13].month5 = (_list_ysrs.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[13].month6 = (_list_ysrs.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[13].month7 = (_list_ysrs.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[13].month8 = (_list_ysrs.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[13].month9 = (_list_ysrs.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[13].month10 = (_list_ysrs.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[13].month11 = (_list_ysrs.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[13].month12 = (_list_ysrs.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                //序号16
                /* 
                  * 非市场人数表：取岗位职级为OP 的 核心定额人数 + 骨干定额人数 + 浮动调整预算人数
                  * 年度目标：取平均值（四舍五入取整数）
                  */
                _list_ysrs = list_ysrs.Where(p => p.postLevel == "OP").ToList();
                list[14].classify = "OP";
                list[14].goal = (_list_ysrs.Sum(p => p.coreQuota + p.boneQuota + p.floattzys) / 12).ToString();
                list[14].month1 = (_list_ysrs.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[14].month2 = (_list_ysrs.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[14].month3 = (_list_ysrs.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[14].month4 = (_list_ysrs.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[14].month5 = (_list_ysrs.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[14].month6 = (_list_ysrs.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[14].month7 = (_list_ysrs.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[14].month8 = (_list_ysrs.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[14].month9 = (_list_ysrs.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[14].month10 = (_list_ysrs.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[14].month11 = (_list_ysrs.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[14].month12 = (_list_ysrs.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                //序号17
                /* 
                  * 非市场人数表：取岗位职级为OO 的 核心定额人数 + 骨干定额人数 + 浮动调整预算人数
                  * 年度目标：取平均值（四舍五入取整数）
                  */
                _list_ysrs = list_ysrs.Where(p => p.postLevel == "OO").ToList();
                list[15].classify = "OO";
                list[15].goal = (_list_ysrs.Sum(p => p.coreQuota + p.boneQuota + p.floattzys) / 12).ToString();
                list[15].month1 = (_list_ysrs.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[15].month2 = (_list_ysrs.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[15].month3 = (_list_ysrs.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[15].month4 = (_list_ysrs.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[15].month5 = (_list_ysrs.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[15].month6 = (_list_ysrs.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[15].month7 = (_list_ysrs.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[15].month8 = (_list_ysrs.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[15].month9 = (_list_ysrs.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[15].month10 = (_list_ysrs.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[15].month11 = (_list_ysrs.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[15].month12 = (_list_ysrs.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                //序号18
                /* 
                  * 非市场人数表：取BMI部门（部门名称为“BMI供应链”）中岗位职级为OO 的 核心定额人数 + 骨干定额人数 + 浮动调整预算人数
                  * 年度目标：取平均值（四舍五入取整数）
                  */
                _list_ysrs = list_ysrs.Where(p => p.postName == "BMI供应链").ToList();
                list[16].classify = "BMI";
                list[16].goal = (_list_ysrs.Sum(p => p.coreQuota + p.boneQuota + p.floattzys) / 12).ToString();
                list[16].month1 = (_list_ysrs.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[16].month2 = (_list_ysrs.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[16].month3 = (_list_ysrs.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[16].month4 = (_list_ysrs.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[16].month5 = (_list_ysrs.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[16].month6 = (_list_ysrs.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[16].month7 = (_list_ysrs.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[16].month8 = (_list_ysrs.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[16].month9 = (_list_ysrs.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[16].month10 = (_list_ysrs.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[16].month11 = (_list_ysrs.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[16].month12 = (_list_ysrs.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                //序号19
                /* 
                  * 非市场人数表：取BMI部门（部门名称为“BPL生产”）中岗位职级为OO 的 核心定额人数 + 骨干定额人数 + 浮动调整预算人数
                  * 年度目标：取平均值（四舍五入取整数）
                  */
                _list_ysrs = list_ysrs.Where(p => p.postName == "BPL生产").ToList();
                list[17].classify = "BPL";
                list[17].goal = (_list_ysrs.Sum(p => p.coreQuota + p.boneQuota + p.floattzys) / 12).ToString();
                list[17].month1 = (_list_ysrs.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[17].month2 = (_list_ysrs.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[17].month3 = (_list_ysrs.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[17].month4 = (_list_ysrs.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[17].month5 = (_list_ysrs.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[17].month6 = (_list_ysrs.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[17].month7 = (_list_ysrs.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[17].month8 = (_list_ysrs.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[17].month9 = (_list_ysrs.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[17].month10 = (_list_ysrs.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[17].month11 = (_list_ysrs.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                list[17].month12 = (_list_ysrs.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                #endregion
                
                #region//序号10
                /*
                  * 总人数 = 营销中心 + 干部 + OP + OO
                  *  = 14 + 15 + 16 + 17
                  */
                list[8].classify = "总人数";
                list[8].goal = (int.Parse(list[12].goal) + int.Parse(list[13].goal) + int.Parse(list[14].goal) + int.Parse(list[15].goal)).ToString();
                list[8].month1 = (int.Parse(list[12].month1) + int.Parse(list[13].month1) + int.Parse(list[14].month1) + int.Parse(list[15].month1)).ToString();
                list[8].month2 = (int.Parse(list[12].month2) + int.Parse(list[13].month2) + int.Parse(list[14].month2) + int.Parse(list[15].month2)).ToString();
                list[8].month3 = (int.Parse(list[12].month3) + int.Parse(list[13].month3) + int.Parse(list[14].month3) + int.Parse(list[15].month3)).ToString();
                list[8].month4 = (int.Parse(list[12].month4) + int.Parse(list[13].month4) + int.Parse(list[14].month4) + int.Parse(list[15].month4)).ToString();
                list[8].month5 = (int.Parse(list[12].month5) + int.Parse(list[13].month5) + int.Parse(list[14].month5) + int.Parse(list[15].month5)).ToString();
                list[8].month6 = (int.Parse(list[12].month6) + int.Parse(list[13].month6) + int.Parse(list[14].month6) + int.Parse(list[15].month6)).ToString();
                list[8].month7 = (int.Parse(list[12].month7) + int.Parse(list[13].month7) + int.Parse(list[14].month7) + int.Parse(list[15].month7)).ToString();
                list[8].month8 = (int.Parse(list[12].month8) + int.Parse(list[13].month8) + int.Parse(list[14].month8) + int.Parse(list[15].month8)).ToString();
                list[8].month9 = (int.Parse(list[12].month9) + int.Parse(list[13].month9) + int.Parse(list[14].month9) + int.Parse(list[15].month9)).ToString();
                list[8].month10 = (int.Parse(list[12].month10) + int.Parse(list[13].month10) + int.Parse(list[14].month10) + int.Parse(list[15].month10)).ToString();
                list[8].month11 = (int.Parse(list[12].month11) + int.Parse(list[13].month11) + int.Parse(list[14].month11) + int.Parse(list[15].month11)).ToString();
                list[8].month12 = (int.Parse(list[12].month12) + int.Parse(list[13].month12) + int.Parse(list[14].month12) + int.Parse(list[15].month12)).ToString();

                #endregion
                                

                #region//人工成本数据
                //收入-人工成本
                var list_rgsr = new List<Hr_Midrgsr>();

                var dt4 = dal.GetHr_Midrgsr(start, end, mid);
                if (dt4.Rows.Count > 0)
                {
                    for (int i = 0; i < dt4.Rows.Count; i++)
                    {
                        var obj4 = new Hr_Midrgsr();
                        
                        obj4.costType = dt4.Rows[i]["costType"].ToString();
                        obj4.adjustBudget = double.Parse(dt4.Rows[i]["adjustBudget"].ToString());
                        obj4.monthly = int.Parse(dt4.Rows[i]["monthly"].ToString());

                        list_rgsr.Add(obj4);
                    }
                }

                //支出-人工成本
                var list_rgzc = new List<Hr_Midrgzc>();

                var dt5 = dal.GetHr_Midrgzc(start, end, mid);
                if (dt5.Rows.Count > 0)
                {
                    for (int i = 0; i < dt5.Rows.Count; i++)
                    {
                        var obj5 = new Hr_Midrgzc();

                        obj5.costType = dt5.Rows[i]["costType"].ToString();
                        obj5.adjustBudget = double.Parse(dt5.Rows[i]["adjustBudget"].ToString());
                        obj5.monthly = int.Parse(dt5.Rows[i]["monthly"].ToString());

                        list_rgzc.Add(obj5);
                    }
                }
                #endregion

                #region//序号20—26
                //序号20
                /*
                  * 收入-人工成本表：取对应月份的调整预算之和
                  * 年度目标：1-12月合计值
                  */
                list[18].classify = "收入-人工成本 (万元)";
                list[18].goal = list_rgsr.Sum(p => p.adjustBudget).ToString();
                list[18].month1 = list_rgsr.Where(p => p.monthly == 1).Sum(p => p.adjustBudget).ToString();
                list[18].month2 = list_rgsr.Where(p => p.monthly == 2).Sum(p => p.adjustBudget).ToString();
                list[18].month3 = list_rgsr.Where(p => p.monthly == 3).Sum(p => p.adjustBudget).ToString();
                list[18].month4 = list_rgsr.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                list[18].month5 = list_rgsr.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                list[18].month6 = list_rgsr.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                list[18].month7 = list_rgsr.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                list[18].month8 = list_rgsr.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                list[18].month9 = list_rgsr.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                list[18].month10 = list_rgsr.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                list[18].month11 = list_rgsr.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                list[18].month12 = list_rgsr.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                //序号21
                /*
                  * 收入-人工成本表：取对应月份的费用类别为‘市场人工’调整预算之和
                  * 年度目标：1-12月合计值
                  */
                var _list_rgsr = list_rgsr.Where(p => p.costType == "市场人工").ToList();
                list[19].classify = "其中：市场人工";
                list[19].goal = _list_rgsr.Sum(p => p.adjustBudget).ToString();
                list[19].month1 = _list_rgsr.Where(p => p.monthly == 1).Sum(p => p.adjustBudget).ToString();
                list[19].month2 = _list_rgsr.Where(p => p.monthly == 2).Sum(p => p.adjustBudget).ToString();
                list[19].month3 = _list_rgsr.Where(p => p.monthly == 3).Sum(p => p.adjustBudget).ToString();
                list[19].month4 = _list_rgsr.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                list[19].month5 = _list_rgsr.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                list[19].month6 = _list_rgsr.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                list[19].month7 = _list_rgsr.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                list[19].month8 = _list_rgsr.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                list[19].month9 = _list_rgsr.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                list[19].month10 = _list_rgsr.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                list[19].month11 = _list_rgsr.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                list[19].month12 = _list_rgsr.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                //序号22
                /*
                  * 收入-人工成本表：取对应月份的费用类别为‘管理人工’调整预算之和
                  * 年度目标：1-12月合计值
                  */
                _list_rgsr = list_rgsr.Where(p => p.costType == "管理人工").ToList();
                list[20].classify = "管理人工";
                list[20].goal = _list_rgsr.Sum(p => p.adjustBudget).ToString();
                list[20].month1 = _list_rgsr.Where(p => p.monthly == 1).Sum(p => p.adjustBudget).ToString();
                list[20].month2 = _list_rgsr.Where(p => p.monthly == 2).Sum(p => p.adjustBudget).ToString();
                list[20].month3 = _list_rgsr.Where(p => p.monthly == 3).Sum(p => p.adjustBudget).ToString();
                list[20].month4 = _list_rgsr.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                list[20].month5 = _list_rgsr.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                list[20].month6 = _list_rgsr.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                list[20].month7 = _list_rgsr.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                list[20].month8 = _list_rgsr.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                list[20].month9 = _list_rgsr.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                list[20].month10 = _list_rgsr.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                list[20].month11 = _list_rgsr.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                list[20].month12 = _list_rgsr.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                //序号23
                /*
                  * 收入-人工成本表：取对应月份的费用类别为‘制造人工’调整预算之和
                  * 年度目标：1-12月合计值
                  */
                _list_rgsr = list_rgsr.Where(p => p.costType == "制造人工").ToList();
                list[21].classify = "制造人工";
                list[21].goal = _list_rgsr.Sum(p => p.adjustBudget).ToString();
                list[21].month1 = _list_rgsr.Where(p => p.monthly == 1).Sum(p => p.adjustBudget).ToString();
                list[21].month2 = _list_rgsr.Where(p => p.monthly == 2).Sum(p => p.adjustBudget).ToString();
                list[21].month3 = _list_rgsr.Where(p => p.monthly == 3).Sum(p => p.adjustBudget).ToString();
                list[21].month4 = _list_rgsr.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                list[21].month5 = _list_rgsr.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                list[21].month6 = _list_rgsr.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                list[21].month7 = _list_rgsr.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                list[21].month8 = _list_rgsr.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                list[21].month9 = _list_rgsr.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                list[21].month10 = _list_rgsr.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                list[21].month11 = _list_rgsr.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                list[21].month12 = _list_rgsr.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                //序号24
                /*
                  * 收入-人工成本表：取对应月份的费用类别为‘直接人工’调整预算之和
                  * 年度目标：1-12月合计值
                  */
                _list_rgsr = list_rgsr.Where(p => p.costType == "直接人工").ToList();
                list[22].classify = "直接人工";
                list[22].goal = _list_rgsr.Sum(p => p.adjustBudget).ToString();
                list[22].month1 = _list_rgsr.Where(p => p.monthly == 1).Sum(p => p.adjustBudget).ToString();
                list[22].month2 = _list_rgsr.Where(p => p.monthly == 2).Sum(p => p.adjustBudget).ToString();
                list[22].month3 = _list_rgsr.Where(p => p.monthly == 3).Sum(p => p.adjustBudget).ToString();
                list[22].month4 = _list_rgsr.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                list[22].month5 = _list_rgsr.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                list[22].month6 = _list_rgsr.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                list[22].month7 = _list_rgsr.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                list[22].month8 = _list_rgsr.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                list[22].month9 = _list_rgsr.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                list[22].month10 = _list_rgsr.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                list[22].month11 = _list_rgsr.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                list[22].month12 = _list_rgsr.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                //序号25
                /*
                  * 收入-人工成本表：取对应月份的费用类别为‘直接人工BMI’调整预算之和
                  * 年度目标：1-12月合计值
                  */
                _list_rgsr = list_rgsr.Where(p => p.costType == "直接人工BMI").ToList();
                list[23].classify = "直接人工BMI";
                list[23].goal = _list_rgsr.Sum(p => p.adjustBudget).ToString();
                list[23].month1 = _list_rgsr.Where(p => p.monthly == 1).Sum(p => p.adjustBudget).ToString();
                list[23].month2 = _list_rgsr.Where(p => p.monthly == 2).Sum(p => p.adjustBudget).ToString();
                list[23].month3 = _list_rgsr.Where(p => p.monthly == 3).Sum(p => p.adjustBudget).ToString();
                list[23].month4 = _list_rgsr.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                list[23].month5 = _list_rgsr.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                list[23].month6 = _list_rgsr.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                list[23].month7 = _list_rgsr.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                list[23].month8 = _list_rgsr.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                list[23].month9 = _list_rgsr.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                list[23].month10 = _list_rgsr.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                list[23].month11 = _list_rgsr.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                list[23].month12 = _list_rgsr.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                //序号26
                /*
                  * 收入-人工成本表：取对应月份的费用类别为‘直接人工BPL’调整预算之和
                  * 年度目标：1-12月合计值
                  */
                _list_rgsr = list_rgsr.Where(p => p.costType == "直接人工BPL").ToList();
                list[24].classify = "直接人工BPL";
                list[24].goal = _list_rgsr.Sum(p => p.adjustBudget).ToString();
                list[24].month1 = _list_rgsr.Where(p => p.monthly == 1).Sum(p => p.adjustBudget).ToString();
                list[24].month2 = _list_rgsr.Where(p => p.monthly == 2).Sum(p => p.adjustBudget).ToString();
                list[24].month3 = _list_rgsr.Where(p => p.monthly == 3).Sum(p => p.adjustBudget).ToString();
                list[24].month4 = _list_rgsr.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                list[24].month5 = _list_rgsr.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                list[24].month6 = _list_rgsr.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                list[24].month7 = _list_rgsr.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                list[24].month8 = _list_rgsr.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                list[24].month9 = _list_rgsr.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                list[24].month10 = _list_rgsr.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                list[24].month11 = _list_rgsr.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                list[24].month12 = _list_rgsr.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                #endregion

                #region//序号27—33
                //序号27
                /*
                  * 支出-人工成本表：取对应月份的调整预算之和
                  * 年度目标：1-12月合计值
                  */
                list[25].classify = "收入-人工成本 (万元)";
                list[25].goal = list_rgzc.Sum(p => p.adjustBudget).ToString();
                list[25].month1 = list_rgzc.Where(p => p.monthly == 1).Sum(p => p.adjustBudget).ToString();
                list[25].month2 = list_rgzc.Where(p => p.monthly == 2).Sum(p => p.adjustBudget).ToString();
                list[25].month3 = list_rgzc.Where(p => p.monthly == 3).Sum(p => p.adjustBudget).ToString();
                list[25].month4 = list_rgzc.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                list[25].month5 = list_rgzc.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                list[25].month6 = list_rgzc.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                list[25].month7 = list_rgzc.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                list[25].month8 = list_rgzc.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                list[25].month9 = list_rgzc.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                list[25].month10 = list_rgzc.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                list[25].month11 = list_rgzc.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                list[25].month12 = list_rgzc.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                //序号28
                /*
                  * 支出-人工成本表：取对应月份的费用类别为‘市场人工’调整预算之和
                  * 年度目标：1-12月合计值
                  */
                var _list_rgzc = list_rgzc.Where(p => p.costType == "市场人工").ToList();
                list[26].classify = "其中：市场人工";
                list[26].goal = _list_rgzc.Sum(p => p.adjustBudget).ToString();
                list[26].month1 = _list_rgzc.Where(p => p.monthly == 1).Sum(p => p.adjustBudget).ToString();
                list[26].month2 = _list_rgzc.Where(p => p.monthly == 2).Sum(p => p.adjustBudget).ToString();
                list[26].month3 = _list_rgzc.Where(p => p.monthly == 3).Sum(p => p.adjustBudget).ToString();
                list[26].month4 = _list_rgzc.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                list[26].month5 = _list_rgzc.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                list[26].month6 = _list_rgzc.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                list[26].month7 = _list_rgzc.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                list[26].month8 = _list_rgzc.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                list[26].month9 = _list_rgzc.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                list[26].month10 = _list_rgzc.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                list[26].month11 = _list_rgzc.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                list[26].month12 = _list_rgzc.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                //序号29
                /*
                  * 支出-人工成本表：取对应月份的费用类别为‘管理人工’调整预算之和
                  * 年度目标：1-12月合计值
                  */
                _list_rgzc = list_rgzc.Where(p => p.costType == "管理人工").ToList();
                list[27].classify = "管理人工";
                list[27].goal = _list_rgzc.Sum(p => p.adjustBudget).ToString();
                list[27].month1 = _list_rgzc.Where(p => p.monthly == 1).Sum(p => p.adjustBudget).ToString();
                list[27].month2 = _list_rgzc.Where(p => p.monthly == 2).Sum(p => p.adjustBudget).ToString();
                list[27].month3 = _list_rgzc.Where(p => p.monthly == 3).Sum(p => p.adjustBudget).ToString();
                list[27].month4 = _list_rgzc.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                list[27].month5 = _list_rgzc.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                list[27].month6 = _list_rgzc.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                list[27].month7 = _list_rgzc.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                list[27].month8 = _list_rgzc.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                list[27].month9 = _list_rgzc.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                list[27].month10 = _list_rgzc.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                list[27].month11 = _list_rgzc.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                list[27].month12 = _list_rgzc.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                //序号30
                /*
                  * 支出-人工成本表：取对应月份的费用类别为‘制造人工’调整预算之和
                  * 年度目标：1-12月合计值
                  */
                _list_rgzc = list_rgzc.Where(p => p.costType == "制造人工").ToList();
                list[28].classify = "制造人工";
                list[28].goal = _list_rgzc.Sum(p => p.adjustBudget).ToString();
                list[28].month1 = _list_rgzc.Where(p => p.monthly == 1).Sum(p => p.adjustBudget).ToString();
                list[28].month2 = _list_rgzc.Where(p => p.monthly == 2).Sum(p => p.adjustBudget).ToString();
                list[28].month3 = _list_rgzc.Where(p => p.monthly == 3).Sum(p => p.adjustBudget).ToString();
                list[28].month4 = _list_rgzc.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                list[28].month5 = _list_rgzc.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                list[28].month6 = _list_rgzc.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                list[28].month7 = _list_rgzc.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                list[28].month8 = _list_rgzc.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                list[28].month9 = _list_rgzc.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                list[28].month10 = _list_rgzc.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                list[28].month11 = _list_rgzc.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                list[28].month12 = _list_rgzc.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                //序号31
                /*
                  * 支出-人工成本表：取对应月份的费用类别为‘直接人工’调整预算之和
                  * 年度目标：1-12月合计值
                  */
                _list_rgzc = list_rgzc.Where(p => p.costType == "直接人工").ToList();
                list[29].classify = "直接人工";
                list[29].goal = _list_rgzc.Sum(p => p.adjustBudget).ToString();
                list[29].month1 = _list_rgzc.Where(p => p.monthly == 1).Sum(p => p.adjustBudget).ToString();
                list[29].month2 = _list_rgzc.Where(p => p.monthly == 2).Sum(p => p.adjustBudget).ToString();
                list[29].month3 = _list_rgzc.Where(p => p.monthly == 3).Sum(p => p.adjustBudget).ToString();
                list[29].month4 = _list_rgzc.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                list[29].month5 = _list_rgzc.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                list[29].month6 = _list_rgzc.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                list[29].month7 = _list_rgzc.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                list[29].month8 = _list_rgzc.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                list[29].month9 = _list_rgzc.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                list[29].month10 = _list_rgzc.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                list[29].month11 = _list_rgzc.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                list[29].month12 = _list_rgzc.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                //序号32
                /*
                  * 支出-人工成本表：取对应月份的费用类别为‘直接人工BMI’调整预算之和
                  * 年度目标：1-12月合计值
                  */
                _list_rgzc = list_rgzc.Where(p => p.costType == "直接人工BMI").ToList();
                list[30].classify = "直接人工BMI";
                list[30].goal = _list_rgzc.Sum(p => p.adjustBudget).ToString();
                list[30].month1 = _list_rgzc.Where(p => p.monthly == 1).Sum(p => p.adjustBudget).ToString();
                list[30].month2 = _list_rgzc.Where(p => p.monthly == 2).Sum(p => p.adjustBudget).ToString();
                list[30].month3 = _list_rgzc.Where(p => p.monthly == 3).Sum(p => p.adjustBudget).ToString();
                list[30].month4 = _list_rgzc.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                list[30].month5 = _list_rgzc.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                list[30].month6 = _list_rgzc.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                list[30].month7 = _list_rgzc.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                list[30].month8 = _list_rgzc.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                list[30].month9 = _list_rgzc.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                list[30].month10 = _list_rgzc.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                list[30].month11 = _list_rgzc.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                list[30].month12 = _list_rgzc.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                //序号33
                /*
                  * 支出-人工成本表：取对应月份的费用类别为‘直接人工BPL’调整预算之和
                  * 年度目标：1-12月合计值
                  */
                _list_rgzc = list_rgzc.Where(p => p.costType == "直接人工BPL").ToList();
                list[31].classify = "直接人工BPL";
                list[31].goal = _list_rgzc.Sum(p => p.adjustBudget).ToString();
                list[31].month1 = _list_rgzc.Where(p => p.monthly == 1).Sum(p => p.adjustBudget).ToString();
                list[31].month2 = _list_rgzc.Where(p => p.monthly == 2).Sum(p => p.adjustBudget).ToString();
                list[31].month3 = _list_rgzc.Where(p => p.monthly == 3).Sum(p => p.adjustBudget).ToString();
                list[31].month4 = _list_rgzc.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                list[31].month5 = _list_rgzc.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                list[31].month6 = _list_rgzc.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                list[31].month7 = _list_rgzc.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                list[31].month8 = _list_rgzc.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                list[31].month9 = _list_rgzc.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                list[31].month10 = _list_rgzc.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                list[31].month11 = _list_rgzc.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                list[31].month12 = _list_rgzc.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                #endregion

                #region//序号34—40
                /*
                  * 盈亏： 收入 -  支出 
                  */
                //序号34
                list[32].classify = "盈亏（万）";
                list[32].goal = (double.Parse(list[18].goal) - double.Parse(list[25].goal)).ToString();
                list[32].month1 = (int.Parse(list[18].month1) - int.Parse(list[25].month1)).ToString();
                list[32].month2 = (int.Parse(list[18].month2) - int.Parse(list[25].month2)).ToString();
                list[32].month3 = (int.Parse(list[18].month3) - int.Parse(list[25].month3)).ToString();
                list[32].month4 = (int.Parse(list[18].month4) - int.Parse(list[25].month4)).ToString();
                list[32].month5 = (int.Parse(list[18].month5) - int.Parse(list[25].month5)).ToString();
                list[32].month6 = (int.Parse(list[18].month6) - int.Parse(list[25].month6)).ToString();
                list[32].month7 = (int.Parse(list[18].month7) - int.Parse(list[25].month7)).ToString();
                list[32].month8 = (int.Parse(list[18].month8) - int.Parse(list[25].month8)).ToString();
                list[32].month9 = (int.Parse(list[18].month9) - int.Parse(list[25].month9)).ToString();
                list[32].month10 = (int.Parse(list[18].month10) - int.Parse(list[25].month10)).ToString();
                list[32].month11 = (int.Parse(list[18].month11) - int.Parse(list[25].month11)).ToString();
                list[32].month12 = (int.Parse(list[18].month12) - int.Parse(list[25].month12)).ToString();

                //序号35
                list[33].classify = "其中：市场人工";
                list[33].goal = (double.Parse(list[19].goal) - double.Parse(list[26].goal)).ToString();
                list[33].month1 = (int.Parse(list[19].month1) - int.Parse(list[26].month1)).ToString();
                list[33].month2 = (int.Parse(list[19].month2) - int.Parse(list[26].month2)).ToString();
                list[33].month3 = (int.Parse(list[19].month3) - int.Parse(list[26].month3)).ToString();
                list[33].month4 = (int.Parse(list[19].month4) - int.Parse(list[26].month4)).ToString();
                list[33].month5 = (int.Parse(list[19].month5) - int.Parse(list[26].month5)).ToString();
                list[33].month6 = (int.Parse(list[19].month6) - int.Parse(list[26].month6)).ToString();
                list[33].month7 = (int.Parse(list[19].month7) - int.Parse(list[26].month7)).ToString();
                list[33].month8 = (int.Parse(list[19].month8) - int.Parse(list[26].month8)).ToString();
                list[33].month9 = (int.Parse(list[19].month9) - int.Parse(list[26].month9)).ToString();
                list[33].month10 = (int.Parse(list[19].month10) - int.Parse(list[26].month10)).ToString();
                list[33].month11 = (int.Parse(list[19].month11) - int.Parse(list[26].month11)).ToString();
                list[33].month12 = (int.Parse(list[19].month12) - int.Parse(list[26].month12)).ToString();

                //序号36
                list[34].classify = "管理人工";
                list[34].goal = (double.Parse(list[20].goal) - double.Parse(list[27].goal)).ToString();
                list[34].month1 = (int.Parse(list[20].month1) - int.Parse(list[27].month1)).ToString();
                list[34].month2 = (int.Parse(list[20].month2) - int.Parse(list[27].month2)).ToString();
                list[34].month3 = (int.Parse(list[20].month3) - int.Parse(list[27].month3)).ToString();
                list[34].month4 = (int.Parse(list[20].month4) - int.Parse(list[27].month4)).ToString();
                list[34].month5 = (int.Parse(list[20].month5) - int.Parse(list[27].month5)).ToString();
                list[34].month6 = (int.Parse(list[20].month6) - int.Parse(list[27].month6)).ToString();
                list[34].month7 = (int.Parse(list[20].month7) - int.Parse(list[27].month7)).ToString();
                list[34].month8 = (int.Parse(list[20].month8) - int.Parse(list[27].month8)).ToString();
                list[34].month9 = (int.Parse(list[20].month9) - int.Parse(list[27].month9)).ToString();
                list[34].month10 = (int.Parse(list[20].month10) - int.Parse(list[27].month10)).ToString();
                list[34].month11 = (int.Parse(list[20].month11) - int.Parse(list[27].month11)).ToString();
                list[34].month12 = (int.Parse(list[20].month12) - int.Parse(list[27].month12)).ToString();

                //序号37
                list[35].classify = "制造人工";
                list[35].goal = (double.Parse(list[21].goal) - double.Parse(list[28].goal)).ToString();
                list[35].month1 = (int.Parse(list[21].month1) - int.Parse(list[28].month1)).ToString();
                list[35].month2 = (int.Parse(list[21].month2) - int.Parse(list[28].month2)).ToString();
                list[35].month3 = (int.Parse(list[21].month3) - int.Parse(list[28].month3)).ToString();
                list[35].month4 = (int.Parse(list[21].month4) - int.Parse(list[28].month4)).ToString();
                list[35].month5 = (int.Parse(list[21].month5) - int.Parse(list[28].month5)).ToString();
                list[35].month6 = (int.Parse(list[21].month6) - int.Parse(list[28].month6)).ToString();
                list[35].month7 = (int.Parse(list[21].month7) - int.Parse(list[28].month7)).ToString();
                list[35].month8 = (int.Parse(list[21].month8) - int.Parse(list[28].month8)).ToString();
                list[35].month9 = (int.Parse(list[21].month9) - int.Parse(list[28].month9)).ToString();
                list[35].month10 = (int.Parse(list[21].month10) - int.Parse(list[28].month10)).ToString();
                list[35].month11 = (int.Parse(list[21].month11) - int.Parse(list[28].month11)).ToString();
                list[35].month12 = (int.Parse(list[21].month12) - int.Parse(list[28].month12)).ToString();

                //序号38
                list[36].classify = "直接人工";
                list[36].goal = (double.Parse(list[22].goal) - double.Parse(list[29].goal)).ToString();
                list[36].month1 = (int.Parse(list[22].month1) - int.Parse(list[29].month1)).ToString();
                list[36].month2 = (int.Parse(list[22].month2) - int.Parse(list[29].month2)).ToString();
                list[36].month3 = (int.Parse(list[22].month3) - int.Parse(list[29].month3)).ToString();
                list[36].month4 = (int.Parse(list[22].month4) - int.Parse(list[29].month4)).ToString();
                list[36].month5 = (int.Parse(list[22].month5) - int.Parse(list[29].month5)).ToString();
                list[36].month6 = (int.Parse(list[22].month6) - int.Parse(list[29].month6)).ToString();
                list[36].month7 = (int.Parse(list[22].month7) - int.Parse(list[29].month7)).ToString();
                list[36].month8 = (int.Parse(list[22].month8) - int.Parse(list[29].month8)).ToString();
                list[36].month9 = (int.Parse(list[22].month9) - int.Parse(list[29].month9)).ToString();
                list[36].month10 = (int.Parse(list[22].month10) - int.Parse(list[29].month10)).ToString();
                list[36].month11 = (int.Parse(list[22].month11) - int.Parse(list[29].month11)).ToString();
                list[36].month12 = (int.Parse(list[22].month12) - int.Parse(list[29].month12)).ToString();

                //序号39
                list[37].classify = "直接人工BMI";
                list[37].goal = (double.Parse(list[23].goal) - double.Parse(list[30].goal)).ToString();
                list[37].month1 = (int.Parse(list[23].month1) - int.Parse(list[30].month1)).ToString();
                list[37].month2 = (int.Parse(list[23].month2) - int.Parse(list[30].month2)).ToString();
                list[37].month3 = (int.Parse(list[23].month3) - int.Parse(list[30].month3)).ToString();
                list[37].month4 = (int.Parse(list[23].month4) - int.Parse(list[30].month4)).ToString();
                list[37].month5 = (int.Parse(list[23].month5) - int.Parse(list[30].month5)).ToString();
                list[37].month6 = (int.Parse(list[23].month6) - int.Parse(list[30].month6)).ToString();
                list[37].month7 = (int.Parse(list[23].month7) - int.Parse(list[30].month7)).ToString();
                list[37].month8 = (int.Parse(list[23].month8) - int.Parse(list[30].month8)).ToString();
                list[37].month9 = (int.Parse(list[23].month9) - int.Parse(list[30].month9)).ToString();
                list[37].month10 = (int.Parse(list[23].month10) - int.Parse(list[30].month10)).ToString();
                list[37].month11 = (int.Parse(list[23].month11) - int.Parse(list[30].month11)).ToString();
                list[37].month12 = (int.Parse(list[23].month12) - int.Parse(list[30].month12)).ToString();

                //序号40
                list[38].classify = "直接人工BPL";
                list[38].goal = (double.Parse(list[24].goal) - double.Parse(list[31].goal)).ToString();
                list[38].month1 = (int.Parse(list[24].month1) - int.Parse(list[31].month1)).ToString();
                list[38].month2 = (int.Parse(list[24].month2) - int.Parse(list[31].month2)).ToString();
                list[38].month3 = (int.Parse(list[24].month3) - int.Parse(list[31].month3)).ToString();
                list[38].month4 = (int.Parse(list[24].month4) - int.Parse(list[31].month4)).ToString();
                list[38].month5 = (int.Parse(list[24].month5) - int.Parse(list[31].month5)).ToString();
                list[38].month6 = (int.Parse(list[24].month6) - int.Parse(list[31].month6)).ToString();
                list[38].month7 = (int.Parse(list[24].month7) - int.Parse(list[31].month7)).ToString();
                list[38].month8 = (int.Parse(list[24].month8) - int.Parse(list[31].month8)).ToString();
                list[38].month9 = (int.Parse(list[24].month9) - int.Parse(list[31].month9)).ToString();
                list[38].month10 = (int.Parse(list[24].month10) - int.Parse(list[31].month10)).ToString();
                list[38].month11 = (int.Parse(list[24].month11) - int.Parse(list[31].month11)).ToString();
                list[38].month12 = (int.Parse(list[24].month12) - int.Parse(list[31].month12)).ToString();

                #endregion


                #region//序号7—9
                //序号7
                /* (序号3)  / (序号10)
                  * 年度目标：年度目标营收(序号3) 除以 年度目标总人数(序号10)
                  * 月份： 累计目标营收（和） 除以 人数月份（平均值）
                  * 小数点2位
                  */

                //序号8
                /* (序号27)  / (序号3)
                  * 年度目标：支出-人工成本 除以 营收
                  * 月份： 累计支出-人工成本（和） 除以 营收月份累计（和）
                  * 取百分比（*100+%）
                  */

                //序号9
                /* (序号4)  / (序号27)
                  * 年度目标：利润 除以 支出-人工成本
                  * 月份： 利润累计（和） 除以 累计支出-人工成本（和）
                  * 小数点2位
                  */

                list[5].classify = "人均产值(万元/人)";
                list[5].goal = string.Format("{0:N2}", double.Parse(list[1].goal) / int.Parse(list[8].goal));

                list[6].classify = "支出-人工成本占营收比";
                list[6].goal = string.Format("{0:P}", double.Parse(list[25].goal) / double.Parse(list[1].goal));

                list[7].classify = "劳动效率";
                list[7].goal = string.Format("{0:N2}", double.Parse(list[2].goal) / double.Parse(list[25].goal));

                double ysTotal = double.Parse(list[1].month1);  //营收
                int perTotal = int.Parse(list[8].month1);      //人数
                double outTotal = double.Parse(list[25].month1);//支出
                double lrTotal = double.Parse(list[2].month1);  //利润
                list[5].month1 = string.Format("{0:N2}", ysTotal / perTotal);
                list[6].month1 = string.Format("{0:P}", outTotal / ysTotal);
                list[7].month1 = string.Format("{0:N2}", lrTotal / outTotal);
                ysTotal += double.Parse(list[1].month2);
                perTotal += int.Parse(list[8].month2);
                outTotal += int.Parse(list[25].month2);
                lrTotal += int.Parse(list[2].month2);
                list[5].month2 = string.Format("{0:N2}", ysTotal / (perTotal / 2));
                list[6].month2 = string.Format("{0:P}", outTotal / ysTotal);
                list[7].month2 = string.Format("{0:N2}", lrTotal / outTotal);
                ysTotal += double.Parse(list[1].month3);
                perTotal += int.Parse(list[8].month3);
                outTotal += int.Parse(list[25].month3);
                lrTotal += int.Parse(list[2].month3);
                list[5].month3 = string.Format("{0:N2}", ysTotal / (perTotal / 3));
                list[6].month3 = string.Format("{0:P}", outTotal / ysTotal);
                list[7].month3 = string.Format("{0:N2}", lrTotal / outTotal);
                ysTotal += double.Parse(list[1].month4);
                perTotal += int.Parse(list[8].month4);
                outTotal += int.Parse(list[25].month4);
                lrTotal += int.Parse(list[2].month4);
                list[5].month4 = string.Format("{0:N2}", ysTotal / (perTotal / 4));
                list[6].month4 = string.Format("{0:P}", outTotal / ysTotal);
                list[7].month4 = string.Format("{0:N2}", lrTotal / outTotal);
                ysTotal += double.Parse(list[1].month5);
                perTotal += int.Parse(list[8].month5);
                outTotal += int.Parse(list[25].month5);
                lrTotal += int.Parse(list[2].month5);
                list[5].month5 = string.Format("{0:N2}", ysTotal / (perTotal / 5));
                list[6].month5 = string.Format("{0:P}", outTotal / ysTotal);
                list[7].month5 = string.Format("{0:N2}", lrTotal / outTotal);
                ysTotal += double.Parse(list[1].month6);
                perTotal += int.Parse(list[8].month6);
                outTotal += int.Parse(list[25].month6);
                lrTotal += int.Parse(list[2].month6);
                list[5].month6 = string.Format("{0:N2}", ysTotal / (perTotal / 6));
                list[6].month6 = string.Format("{0:P}", outTotal / ysTotal);
                list[7].month6 = string.Format("{0:N2}", lrTotal / outTotal);
                ysTotal += double.Parse(list[1].month7);
                perTotal += int.Parse(list[8].month7);
                outTotal += int.Parse(list[25].month7);
                lrTotal += int.Parse(list[2].month7);
                list[5].month7 = string.Format("{0:N2}", ysTotal / (perTotal / 7));
                list[6].month7 = string.Format("{0:P}", outTotal / ysTotal);
                list[7].month7 = string.Format("{0:N2}", lrTotal / outTotal);
                ysTotal += double.Parse(list[1].month8);
                perTotal += int.Parse(list[8].month8);
                outTotal += int.Parse(list[25].month8);
                lrTotal += int.Parse(list[2].month8);
                list[5].month8 = string.Format("{0:N2}", ysTotal / (perTotal / 8));
                list[6].month8 = string.Format("{0:P}", outTotal / ysTotal);
                list[7].month8 = string.Format("{0:N2}", lrTotal / outTotal);
                ysTotal += double.Parse(list[1].month9);
                perTotal += int.Parse(list[8].month9);
                outTotal += int.Parse(list[25].month9);
                lrTotal += int.Parse(list[2].month9);
                list[5].month9 = string.Format("{0:N2}", ysTotal / (perTotal / 9));
                list[6].month9 = string.Format("{0:P}", outTotal / ysTotal);
                list[7].month9 = string.Format("{0:N2}", lrTotal / outTotal);
                ysTotal += double.Parse(list[1].month10);
                perTotal += int.Parse(list[8].month10);
                outTotal += int.Parse(list[25].month10);
                lrTotal += int.Parse(list[2].month10);
                list[5].month10 = string.Format("{0:N2}", ysTotal / (perTotal / 10));
                list[6].month10 = string.Format("{0:P}", outTotal / ysTotal);
                list[7].month10 = string.Format("{0:N2}", lrTotal / outTotal);
                ysTotal += double.Parse(list[1].month11);
                perTotal += int.Parse(list[8].month11);
                outTotal += int.Parse(list[25].month11);
                lrTotal += int.Parse(list[2].month11);
                list[5].month11 = string.Format("{0:N2}", ysTotal / (perTotal / 11));
                list[6].month11 = string.Format("{0:P}", outTotal / ysTotal);
                list[7].month11 = string.Format("{0:N2}", lrTotal / outTotal);
                ysTotal += double.Parse(list[1].month12);
                perTotal += int.Parse(list[8].month12);
                outTotal += int.Parse(list[25].month12);
                lrTotal += int.Parse(list[2].month12);
                list[5].month12 = string.Format("{0:N2}", ysTotal / (perTotal / 12));
                list[6].month12 = string.Format("{0:P}", outTotal / ysTotal);
                list[7].month12 = string.Format("{0:N2}", lrTotal / outTotal);

                #endregion


                return list;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public class Model
        {
            public string classify;
            public string goal;
            public string month1;
            public string month2;
            public string month3;
            public string month4;
            public string month5;
            public string month6;
            public string month7;
            public string month8;
            public string month9;
            public string month10;
            public string month11;
            public string month12;

            public string yearLine;
            public string sj;
            public string yj;
            public string ys;
        }


        //// HRM项目进展
        [MyAuthAttribute]
        public ActionResult Hr_ModelPro()
        {
            HttpCookie cookie = GetCookie();
            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //表头
        public string GetColData()
        {

            int startMonth = int.Parse(Request["startMonth"]);
            int endMonth = int.Parse(Request["endMonth"]);

            var array = new JArray();

            var json1 = new JObject();
            json1.Add("field", "id");
            json1.Add("align", "center");
            json1.Add("valign", "middle");
            json1.Add("width", "100");
            json1.Add("title", "序号");
            array.Add(json1);
            var json2 = new JObject();
            json2.Add("field", "classify");
            json2.Add("align", "center");
            json2.Add("valign", "middle");
            json2.Add("width", "100");
            json2.Add("title", "项目/分类");
            array.Add(json2);
            var json3 = new JObject();
            json3.Add("field", "yearLine");
            json3.Add("align", "center");
            json3.Add("valign", "middle");
            json3.Add("width", "100");
            json3.Add("title", "年度保本线");
            array.Add(json3);
            var json4 = new JObject();
            json4.Add("field", "goal");
            json4.Add("align", "center");
            json4.Add("valign", "middle");
            json4.Add("width", "100");
            json4.Add("title", "年度目标");
            array.Add(json4);

            for (int i = 1; i <= 12; i++)
            {
                var _json = new JObject();
                _json.Add("field", "month" + i);
                _json.Add("align", "center");
                _json.Add("valign", "middle");
                _json.Add("title", i + "月");

                if (i < startMonth || i > endMonth)
                {
                    _json.Add("visible", false);
                }

                array.Add(_json);
            }

            var json5 = new JObject();
            json5.Add("field", "sj");
            json5.Add("align", "center");
            json5.Add("valign", "middle");
            json5.Add("width", "100");
            json5.Add("title", "年累实际");
            array.Add(json5);
            var json6 = new JObject();
            json6.Add("field", "yj");
            json6.Add("align", "center");
            json6.Add("valign", "middle");
            json6.Add("width", "100");
            json6.Add("title", "年累预计");
            array.Add(json6);
            var json7 = new JObject();
            json7.Add("field", "ys");
            json7.Add("align", "center");
            json7.Add("valign", "middle");
            json7.Add("width", "100");
            json7.Add("title", "年累预算");
            array.Add(json7);

            return array.ToString();
        }

        public string GetHr_ModelPro()
        {
            try
            {
                int yearly = int.Parse(Request["yearly"]);
                int startMonth = int.Parse(Request["startMonth"]);
                int endMonth = int.Parse(Request["endMonth"]);

                var cc = ((int)11 / 4).ToString();

                List<Model> list = SetModelProData(yearly, startMonth, endMonth);  // 获取数据

                JArray array = new JArray();
                int count = 0;
                foreach (var item in list)
                {
                    count++;

                    var json = new JObject();

                    json.Add("id", count);
                    json.Add("classify", item.classify);
                    json.Add("yearLine", item.yearLine);
                    json.Add("goal", item.goal);
                    json.Add("month1", item.month1);
                    json.Add("month2", item.month2);
                    json.Add("month3", item.month3);
                    json.Add("month4", item.month4);
                    json.Add("month5", item.month5);
                    json.Add("month6", item.month6);
                    json.Add("month7", item.month7);
                    json.Add("month8", item.month8);
                    json.Add("month9", item.month9);
                    json.Add("month10", item.month10);
                    json.Add("month11", item.month11);
                    json.Add("month12", item.month12);
                    json.Add("sj", item.sj);
                    json.Add("yj", item.yj);
                    json.Add("ys", item.ys);

                    array.Add(json);
                }

                return array.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<Model> SetModelProData(int yearly, int startMonth, int endMonth)
        {
            try
            {
                //组织单元
                var cookie = GetCookie();
                string midStr = cookie["comId"];
                int mid = int.Parse(midStr);

                var _nowDate = DateTime.Now.Date;  //当前日期
                var _nowYear = _nowDate.Year;     //当前年
                var _nowMonth = _nowDate.Month;   //当前月

                DateTime start = DateTime.Parse(yearly + "-01-01");
                DateTime end = DateTime.Parse(yearly + "-12-31");

                var list = new List<Model>();

                #region//初始返回结果对象 共63-2=61条
                for (int i = 0; i < 61; i++)
                {
                    var obj = new Model();
                    obj.classify = "";
                    obj.yearLine = "1";
                    obj.goal = "1";
                    obj.month1 = "1";
                    obj.month2 = "1";
                    obj.month3 = "1";
                    obj.month4 = "1";
                    obj.month5 = "1";
                    obj.month6 = "1";
                    obj.month7 = "1";
                    obj.month8 = "1";
                    obj.month9 = "1";
                    obj.month10 = "1";
                    obj.month11 = "1";
                    obj.month12 = "1";

                    obj.sj = "1";
                    obj.yj = "1";
                    obj.ys = "1";

                    list.Add(obj);
                }
                #endregion

                var dal = new dalPro();

                #region//调整预算，实际，项目进展 数据
                var list_tzys = new List<Hr_Midtzys>();

                var dt_tzys = dal.GetHr_Midtzys(start, end, mid);
                if (dt_tzys.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_tzys.Rows.Count; i++)
                    {
                        var obj1 = new Hr_Midtzys();

                        obj1.htAmount = double.Parse(dt_tzys.Rows[i]["htAmount"].ToString());
                        obj1.ysAmount = double.Parse(dt_tzys.Rows[i]["ysAmount"].ToString());
                        obj1.lrAmount = double.Parse(dt_tzys.Rows[i]["lrAmount"].ToString());
                        obj1.yield = double.Parse(dt_tzys.Rows[i]["yield"].ToString());
                        obj1.yieEffic = double.Parse(dt_tzys.Rows[i]["yieEffic"].ToString());
                        obj1.proTeams = int.Parse(dt_tzys.Rows[i]["proTeams"].ToString());
                        obj1.monthly = int.Parse(dt_tzys.Rows[i]["monthly"].ToString());

                        list_tzys.Add(obj1);
                    }
                }

                //分开取 1到12月 调整预算
                var list_tzys1 = list_tzys.Where(p => p.monthly == 1).ToList();
                var list_tzys2 = list_tzys.Where(p => p.monthly == 2).ToList();
                var list_tzys3 = list_tzys.Where(p => p.monthly == 3).ToList();
                var list_tzys4 = list_tzys.Where(p => p.monthly == 4).ToList();
                var list_tzys5 = list_tzys.Where(p => p.monthly == 5).ToList();
                var list_tzys6 = list_tzys.Where(p => p.monthly == 6).ToList();
                var list_tzys7 = list_tzys.Where(p => p.monthly == 7).ToList();
                var list_tzys8 = list_tzys.Where(p => p.monthly == 8).ToList();
                var list_tzys9 = list_tzys.Where(p => p.monthly == 9).ToList();
                var list_tzys10 = list_tzys.Where(p => p.monthly == 10).ToList();
                var list_tzys11 = list_tzys.Where(p => p.monthly == 11).ToList();
                var list_tzys12 = list_tzys.Where(p => p.monthly == 12).ToList();


                var list_fact = new List<Hr_Midtzys>();

                var dt_fact = dal.GetHr_Midfact(start, end, mid);
                if (dt_fact.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_fact.Rows.Count; i++)
                    {
                        var obj1 = new Hr_Midtzys();

                        obj1.htAmount = double.Parse(dt_fact.Rows[i]["htAmount"].ToString());
                        obj1.ysAmount = double.Parse(dt_fact.Rows[i]["ysAmount"].ToString());
                        obj1.lrAmount = double.Parse(dt_fact.Rows[i]["lrAmount"].ToString());
                        obj1.yield = double.Parse(dt_fact.Rows[i]["yield"].ToString());
                        obj1.yieEffic = double.Parse(dt_fact.Rows[i]["yieEffic"].ToString());
                        obj1.proTeams = int.Parse(dt_fact.Rows[i]["proTeams"].ToString());
                        obj1.monthly = int.Parse(dt_fact.Rows[i]["monthly"].ToString());

                        list_fact.Add(obj1);
                    }
                }

                //分开取 1到12月 调整预算
                var list_fact1 = list_fact.Where(p => p.monthly == 1).ToList();
                var list_fact2 = list_fact.Where(p => p.monthly == 2).ToList();
                var list_fact3 = list_fact.Where(p => p.monthly == 3).ToList();
                var list_fact4 = list_fact.Where(p => p.monthly == 4).ToList();
                var list_fact5 = list_fact.Where(p => p.monthly == 5).ToList();
                var list_fact6 = list_fact.Where(p => p.monthly == 6).ToList();
                var list_fact7 = list_fact.Where(p => p.monthly == 7).ToList();
                var list_fact8 = list_fact.Where(p => p.monthly == 8).ToList();
                var list_fact9 = list_fact.Where(p => p.monthly == 9).ToList();
                var list_fact10 = list_fact.Where(p => p.monthly == 10).ToList();
                var list_fact11 = list_fact.Where(p => p.monthly == 11).ToList();
                var list_fact12 = list_fact.Where(p => p.monthly == 12).ToList();

                var list_xmjz = new List<Hr_Midtzys>();

                var dt_xmjz = dal.GetHr_Midxmjz(start, end, mid);
                if (dt_xmjz.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_xmjz.Rows.Count; i++)
                    {
                        var obj1 = new Hr_Midtzys();

                        obj1.htAmount = double.Parse(dt_xmjz.Rows[i]["htAmount"].ToString());
                        obj1.ysAmount = double.Parse(dt_xmjz.Rows[i]["ysAmount"].ToString());
                        obj1.lrAmount = double.Parse(dt_xmjz.Rows[i]["lrAmount"].ToString());
                        obj1.yield = double.Parse(dt_xmjz.Rows[i]["proBudget"].ToString());
                        obj1.yieEffic = double.Parse(dt_xmjz.Rows[i]["yieEffic"].ToString());
                        obj1.proTeams = int.Parse(dt_xmjz.Rows[i]["proTeams"].ToString());
                        obj1.monthly = int.Parse(dt_xmjz.Rows[i]["monthly"].ToString());

                        list_xmjz.Add(obj1);
                    }
                }

                //分开取 1到12月 调整预算
                var list_xmjz1 = list_xmjz.Where(p => p.monthly == 1).ToList();
                var list_xmjz2 = list_xmjz.Where(p => p.monthly == 2).ToList();
                var list_xmjz3 = list_xmjz.Where(p => p.monthly == 3).ToList();
                var list_xmjz4 = list_xmjz.Where(p => p.monthly == 4).ToList();
                var list_xmjz5 = list_xmjz.Where(p => p.monthly == 5).ToList();
                var list_xmjz6 = list_xmjz.Where(p => p.monthly == 6).ToList();
                var list_xmjz7 = list_xmjz.Where(p => p.monthly == 7).ToList();
                var list_xmjz8 = list_xmjz.Where(p => p.monthly == 8).ToList();
                var list_xmjz9 = list_xmjz.Where(p => p.monthly == 9).ToList();
                var list_xmjz10 = list_xmjz.Where(p => p.monthly == 10).ToList();
                var list_xmjz11 = list_xmjz.Where(p => p.monthly == 11).ToList();
                var list_xmjz12 = list_xmjz.Where(p => p.monthly == 12).ToList();

                #endregion


                #region 序号2—5，6
                //序号2
                list[0].classify = "合同额(万)";
                list[0].yearLine = "1";
                list[0].goal = list_tzys.Sum(p => p.htAmount).ToString();

                //序号3
                list[1].classify = "营收(万)";
                list[1].yearLine = "1";
                list[1].goal = list_tzys.Sum(p => p.ysAmount).ToString();

                //序号4
                list[2].classify = "利润(万)";
                list[2].yearLine = "1";
                list[2].goal = list_tzys.Sum(p => p.lrAmount).ToString();

                //序号5
                list[3].classify = "产量(万立方)";
                list[3].yearLine = "1";
                list[3].goal = list_tzys.Sum(p => p.yield).ToString();

                //序号6
                list[4].classify = "综合产效（m³/8H/人）";
                list[4].yearLine = "1";
                list[4].goal = (list_tzys.Sum(p => p.yieEffic) / 12).ToString(); //平均值

                /*
                 * 小于当前月，取实际值
                 * 当前月到当前月+2，取项目进展值
                 * 余下月取调整预算值
                 */
                #region//月份
                if (1 == _nowMonth)
                {
                    list[0].month1 = list_xmjz1[0].htAmount.ToString();
                    list[0].month2 = list_xmjz2[0].htAmount.ToString();
                    list[0].month3 = list_xmjz3[0].htAmount.ToString();
                    list[0].month4 = list_tzys4[0].htAmount.ToString();
                    list[0].month5 = list_tzys5[0].htAmount.ToString();
                    list[0].month6 = list_tzys6[0].htAmount.ToString();
                    list[0].month7 = list_tzys7[0].htAmount.ToString();
                    list[0].month8 = list_tzys8[0].htAmount.ToString();
                    list[0].month9 = list_tzys9[0].htAmount.ToString();
                    list[0].month10 = list_tzys10[0].htAmount.ToString();
                    list[0].month11 = list_tzys11[0].htAmount.ToString();
                    list[0].month12 = list_tzys12[0].htAmount.ToString();

                    list[1].month1 = list_xmjz1[0].ysAmount.ToString();
                    list[1].month2 = list_xmjz2[0].ysAmount.ToString();
                    list[1].month3 = list_xmjz3[0].ysAmount.ToString();
                    list[1].month4 = list_tzys4[0].ysAmount.ToString();
                    list[1].month5 = list_tzys5[0].ysAmount.ToString();
                    list[1].month6 = list_tzys6[0].ysAmount.ToString();
                    list[1].month7 = list_tzys7[0].ysAmount.ToString();
                    list[1].month8 = list_tzys8[0].ysAmount.ToString();
                    list[1].month9 = list_tzys9[0].ysAmount.ToString();
                    list[1].month10 = list_tzys10[0].ysAmount.ToString();
                    list[1].month11 = list_tzys11[0].ysAmount.ToString();
                    list[1].month12 = list_tzys12[0].ysAmount.ToString();

                    list[2].month1 = list_xmjz1[0].lrAmount.ToString();
                    list[2].month2 = list_xmjz2[0].lrAmount.ToString();
                    list[2].month3 = list_xmjz3[0].lrAmount.ToString();
                    list[2].month4 = list_tzys4[0].lrAmount.ToString();
                    list[2].month5 = list_tzys5[0].lrAmount.ToString();
                    list[2].month6 = list_tzys6[0].lrAmount.ToString();
                    list[2].month7 = list_tzys7[0].lrAmount.ToString();
                    list[2].month8 = list_tzys8[0].lrAmount.ToString();
                    list[2].month9 = list_tzys9[0].lrAmount.ToString();
                    list[2].month10 = list_tzys10[0].lrAmount.ToString();
                    list[2].month11 = list_tzys11[0].lrAmount.ToString();
                    list[2].month12 = list_tzys12[0].lrAmount.ToString();

                    list[3].month1 = list_xmjz1[0].yield.ToString();
                    list[3].month2 = list_xmjz2[0].yield.ToString();
                    list[3].month3 = list_xmjz3[0].yield.ToString();
                    list[3].month4 = list_tzys4[0].yield.ToString();
                    list[3].month5 = list_tzys5[0].yield.ToString();
                    list[3].month6 = list_tzys6[0].yield.ToString();
                    list[3].month7 = list_tzys7[0].yield.ToString();
                    list[3].month8 = list_tzys8[0].yield.ToString();
                    list[3].month9 = list_tzys9[0].yield.ToString();
                    list[3].month10 = list_tzys10[0].yield.ToString();
                    list[3].month11 = list_tzys11[0].yield.ToString();
                    list[3].month12 = list_tzys12[0].yield.ToString();

                    list[4].month1 = list_xmjz1[0].yieEffic.ToString();
                    list[4].month2 = list_xmjz2[0].yieEffic.ToString();
                    list[4].month3 = list_xmjz3[0].yieEffic.ToString();
                    list[4].month4 = list_tzys4[0].yieEffic.ToString();
                    list[4].month5 = list_tzys5[0].yieEffic.ToString();
                    list[4].month6 = list_tzys6[0].yieEffic.ToString();
                    list[4].month7 = list_tzys7[0].yieEffic.ToString();
                    list[4].month8 = list_tzys8[0].yieEffic.ToString();
                    list[4].month9 = list_tzys9[0].yieEffic.ToString();
                    list[4].month10 = list_tzys10[0].yieEffic.ToString();
                    list[4].month11 = list_tzys11[0].yieEffic.ToString();
                    list[4].month12 = list_tzys12[0].yieEffic.ToString();
                }
                else if (2 == _nowMonth)
                {
                    list[0].month1 = list_fact1[0].htAmount.ToString();
                    list[0].month2 = list_xmjz2[0].htAmount.ToString();
                    list[0].month3 = list_xmjz3[0].htAmount.ToString();
                    list[0].month4 = list_xmjz4[0].htAmount.ToString();
                    list[0].month5 = list_tzys5[0].htAmount.ToString();
                    list[0].month6 = list_tzys6[0].htAmount.ToString();
                    list[0].month7 = list_tzys7[0].htAmount.ToString();
                    list[0].month8 = list_tzys8[0].htAmount.ToString();
                    list[0].month9 = list_tzys9[0].htAmount.ToString();
                    list[0].month10 = list_tzys10[0].htAmount.ToString();
                    list[0].month11 = list_tzys11[0].htAmount.ToString();
                    list[0].month12 = list_tzys12[0].htAmount.ToString();

                    list[1].month1 = list_fact1[0].ysAmount.ToString();
                    list[1].month2 = list_xmjz2[0].ysAmount.ToString();
                    list[1].month3 = list_xmjz3[0].ysAmount.ToString();
                    list[1].month4 = list_xmjz4[0].ysAmount.ToString();
                    list[1].month5 = list_tzys5[0].ysAmount.ToString();
                    list[1].month6 = list_tzys6[0].ysAmount.ToString();
                    list[1].month7 = list_tzys7[0].ysAmount.ToString();
                    list[1].month8 = list_tzys8[0].ysAmount.ToString();
                    list[1].month9 = list_tzys9[0].ysAmount.ToString();
                    list[1].month10 = list_tzys10[0].ysAmount.ToString();
                    list[1].month11 = list_tzys11[0].ysAmount.ToString();
                    list[1].month12 = list_tzys12[0].ysAmount.ToString();

                    list[2].month1 = list_fact1[0].lrAmount.ToString();
                    list[2].month2 = list_xmjz2[0].lrAmount.ToString();
                    list[2].month3 = list_xmjz3[0].lrAmount.ToString();
                    list[2].month4 = list_xmjz4[0].lrAmount.ToString();
                    list[2].month5 = list_tzys5[0].lrAmount.ToString();
                    list[2].month6 = list_tzys6[0].lrAmount.ToString();
                    list[2].month7 = list_tzys7[0].lrAmount.ToString();
                    list[2].month8 = list_tzys8[0].lrAmount.ToString();
                    list[2].month9 = list_tzys9[0].lrAmount.ToString();
                    list[2].month10 = list_tzys10[0].lrAmount.ToString();
                    list[2].month11 = list_tzys11[0].lrAmount.ToString();
                    list[2].month12 = list_tzys12[0].lrAmount.ToString();

                    list[3].month1 = list_fact1[0].yield.ToString();
                    list[3].month2 = list_xmjz2[0].yield.ToString();
                    list[3].month3 = list_xmjz3[0].yield.ToString();
                    list[3].month4 = list_xmjz4[0].yield.ToString();
                    list[3].month5 = list_tzys5[0].yield.ToString();
                    list[3].month6 = list_tzys6[0].yield.ToString();
                    list[3].month7 = list_tzys7[0].yield.ToString();
                    list[3].month8 = list_tzys8[0].yield.ToString();
                    list[3].month9 = list_tzys9[0].yield.ToString();
                    list[3].month10 = list_tzys10[0].yield.ToString();
                    list[3].month11 = list_tzys11[0].yield.ToString();
                    list[3].month12 = list_tzys12[0].yield.ToString();

                    list[4].month1 = list_fact1[0].yieEffic.ToString();
                    list[4].month2 = list_xmjz2[0].yieEffic.ToString();
                    list[4].month3 = list_xmjz3[0].yieEffic.ToString();
                    list[4].month4 = list_xmjz4[0].yieEffic.ToString();
                    list[4].month5 = list_tzys5[0].yieEffic.ToString();
                    list[4].month6 = list_tzys6[0].yieEffic.ToString();
                    list[4].month7 = list_tzys7[0].yieEffic.ToString();
                    list[4].month8 = list_tzys8[0].yieEffic.ToString();
                    list[4].month9 = list_tzys9[0].yieEffic.ToString();
                    list[4].month10 = list_tzys10[0].yieEffic.ToString();
                    list[4].month11 = list_tzys11[0].yieEffic.ToString();
                    list[4].month12 = list_tzys12[0].yieEffic.ToString();
                }
                else if (3 == _nowMonth)
                {
                    list[0].month1 = list_fact1[0].htAmount.ToString();
                    list[0].month2 = list_fact2[0].htAmount.ToString();
                    list[0].month3 = list_xmjz3[0].htAmount.ToString();
                    list[0].month4 = list_xmjz4[0].htAmount.ToString();
                    list[0].month5 = list_xmjz5[0].htAmount.ToString();
                    list[0].month6 = list_tzys6[0].htAmount.ToString();
                    list[0].month7 = list_tzys7[0].htAmount.ToString();
                    list[0].month8 = list_tzys8[0].htAmount.ToString();
                    list[0].month9 = list_tzys9[0].htAmount.ToString();
                    list[0].month10 = list_tzys10[0].htAmount.ToString();
                    list[0].month11 = list_tzys11[0].htAmount.ToString();
                    list[0].month12 = list_tzys12[0].htAmount.ToString();

                    list[1].month1 = list_fact1[0].ysAmount.ToString();
                    list[1].month2 = list_fact2[0].ysAmount.ToString();
                    list[1].month3 = list_xmjz3[0].ysAmount.ToString();
                    list[1].month4 = list_xmjz4[0].ysAmount.ToString();
                    list[1].month5 = list_xmjz5[0].ysAmount.ToString();
                    list[1].month6 = list_tzys6[0].ysAmount.ToString();
                    list[1].month7 = list_tzys7[0].ysAmount.ToString();
                    list[1].month8 = list_tzys8[0].ysAmount.ToString();
                    list[1].month9 = list_tzys9[0].ysAmount.ToString();
                    list[1].month10 = list_tzys10[0].ysAmount.ToString();
                    list[1].month11 = list_tzys11[0].ysAmount.ToString();
                    list[1].month12 = list_tzys12[0].ysAmount.ToString();

                    list[2].month1 = list_fact1[0].lrAmount.ToString();
                    list[2].month2 = list_fact2[0].lrAmount.ToString();
                    list[2].month3 = list_xmjz3[0].lrAmount.ToString();
                    list[2].month4 = list_xmjz4[0].lrAmount.ToString();
                    list[2].month5 = list_xmjz5[0].lrAmount.ToString();
                    list[2].month6 = list_tzys6[0].lrAmount.ToString();
                    list[2].month7 = list_tzys7[0].lrAmount.ToString();
                    list[2].month8 = list_tzys8[0].lrAmount.ToString();
                    list[2].month9 = list_tzys9[0].lrAmount.ToString();
                    list[2].month10 = list_tzys10[0].lrAmount.ToString();
                    list[2].month11 = list_tzys11[0].lrAmount.ToString();
                    list[2].month12 = list_tzys12[0].lrAmount.ToString();

                    list[3].month1 = list_fact1[0].yield.ToString();
                    list[3].month2 = list_fact2[0].yield.ToString();
                    list[3].month3 = list_xmjz3[0].yield.ToString();
                    list[3].month4 = list_xmjz4[0].yield.ToString();
                    list[3].month5 = list_xmjz5[0].yield.ToString();
                    list[3].month6 = list_tzys6[0].yield.ToString();
                    list[3].month7 = list_tzys7[0].yield.ToString();
                    list[3].month8 = list_tzys8[0].yield.ToString();
                    list[3].month9 = list_tzys9[0].yield.ToString();
                    list[3].month10 = list_tzys10[0].yield.ToString();
                    list[3].month11 = list_tzys11[0].yield.ToString();
                    list[3].month12 = list_tzys12[0].yield.ToString();

                    list[4].month1 = list_fact1[0].yieEffic.ToString();
                    list[4].month2 = list_fact2[0].yieEffic.ToString();
                    list[4].month3 = list_xmjz3[0].yieEffic.ToString();
                    list[4].month4 = list_xmjz4[0].yieEffic.ToString();
                    list[4].month5 = list_xmjz5[0].yieEffic.ToString();
                    list[4].month6 = list_tzys6[0].yieEffic.ToString();
                    list[4].month7 = list_tzys7[0].yieEffic.ToString();
                    list[4].month8 = list_tzys8[0].yieEffic.ToString();
                    list[4].month9 = list_tzys9[0].yieEffic.ToString();
                    list[4].month10 = list_tzys10[0].yieEffic.ToString();
                    list[4].month11 = list_tzys11[0].yieEffic.ToString();
                    list[4].month12 = list_tzys12[0].yieEffic.ToString();
                }
                else if (4 == _nowMonth)
                {
                    list[0].month1 = list_fact1[0].htAmount.ToString();
                    list[0].month2 = list_fact2[0].htAmount.ToString();
                    list[0].month3 = list_fact3[0].htAmount.ToString();
                    list[0].month4 = list_xmjz4[0].htAmount.ToString();
                    list[0].month5 = list_xmjz5[0].htAmount.ToString();
                    list[0].month6 = list_xmjz6[0].htAmount.ToString();
                    list[0].month7 = list_tzys7[0].htAmount.ToString();
                    list[0].month8 = list_tzys8[0].htAmount.ToString();
                    list[0].month9 = list_tzys9[0].htAmount.ToString();
                    list[0].month10 = list_tzys10[0].htAmount.ToString();
                    list[0].month11 = list_tzys11[0].htAmount.ToString();
                    list[0].month12 = list_tzys12[0].htAmount.ToString();

                    list[1].month1 = list_fact1[0].ysAmount.ToString();
                    list[1].month2 = list_fact2[0].ysAmount.ToString();
                    list[1].month3 = list_fact3[0].ysAmount.ToString();
                    list[1].month4 = list_xmjz4[0].ysAmount.ToString();
                    list[1].month5 = list_xmjz5[0].ysAmount.ToString();
                    list[1].month6 = list_xmjz6[0].ysAmount.ToString();
                    list[1].month7 = list_tzys7[0].ysAmount.ToString();
                    list[1].month8 = list_tzys8[0].ysAmount.ToString();
                    list[1].month9 = list_tzys9[0].ysAmount.ToString();
                    list[1].month10 = list_tzys10[0].ysAmount.ToString();
                    list[1].month11 = list_tzys11[0].ysAmount.ToString();
                    list[1].month12 = list_tzys12[0].ysAmount.ToString();

                    list[2].month1 = list_fact1[0].lrAmount.ToString();
                    list[2].month2 = list_fact2[0].lrAmount.ToString();
                    list[2].month3 = list_fact3[0].lrAmount.ToString();
                    list[2].month4 = list_xmjz4[0].lrAmount.ToString();
                    list[2].month5 = list_xmjz5[0].lrAmount.ToString();
                    list[2].month6 = list_xmjz6[0].lrAmount.ToString();
                    list[2].month7 = list_tzys7[0].lrAmount.ToString();
                    list[2].month8 = list_tzys8[0].lrAmount.ToString();
                    list[2].month9 = list_tzys9[0].lrAmount.ToString();
                    list[2].month10 = list_tzys10[0].lrAmount.ToString();
                    list[2].month11 = list_tzys11[0].lrAmount.ToString();
                    list[2].month12 = list_tzys12[0].lrAmount.ToString();

                    list[3].month1 = list_fact1[0].yield.ToString();
                    list[3].month2 = list_fact2[0].yield.ToString();
                    list[3].month3 = list_fact3[0].yield.ToString();
                    list[3].month4 = list_xmjz4[0].yield.ToString();
                    list[3].month5 = list_xmjz5[0].yield.ToString();
                    list[3].month6 = list_xmjz6[0].yield.ToString();
                    list[3].month7 = list_tzys7[0].yield.ToString();
                    list[3].month8 = list_tzys8[0].yield.ToString();
                    list[3].month9 = list_tzys9[0].yield.ToString();
                    list[3].month10 = list_tzys10[0].yield.ToString();
                    list[3].month11 = list_tzys11[0].yield.ToString();
                    list[3].month12 = list_tzys12[0].yield.ToString();

                    list[4].month1 = list_fact1[0].yieEffic.ToString();
                    list[4].month2 = list_fact2[0].yieEffic.ToString();
                    list[4].month3 = list_fact3[0].yieEffic.ToString();
                    list[4].month4 = list_xmjz4[0].yieEffic.ToString();
                    list[4].month5 = list_xmjz5[0].yieEffic.ToString();
                    list[4].month6 = list_xmjz6[0].yieEffic.ToString();
                    list[4].month7 = list_tzys7[0].yieEffic.ToString();
                    list[4].month8 = list_tzys8[0].yieEffic.ToString();
                    list[4].month9 = list_tzys9[0].yieEffic.ToString();
                    list[4].month10 = list_tzys10[0].yieEffic.ToString();
                    list[4].month11 = list_tzys11[0].yieEffic.ToString();
                    list[4].month12 = list_tzys12[0].yieEffic.ToString();
                }
                else if (5 == _nowMonth)
                {
                    list[0].month1 = list_fact1[0].htAmount.ToString();
                    list[0].month2 = list_fact2[0].htAmount.ToString();
                    list[0].month3 = list_fact3[0].htAmount.ToString();
                    list[0].month4 = list_fact4[0].htAmount.ToString();
                    list[0].month5 = list_xmjz5[0].htAmount.ToString();
                    list[0].month6 = list_xmjz6[0].htAmount.ToString();
                    list[0].month7 = list_xmjz7[0].htAmount.ToString();
                    list[0].month8 = list_tzys8[0].htAmount.ToString();
                    list[0].month9 = list_tzys9[0].htAmount.ToString();
                    list[0].month10 = list_tzys10[0].htAmount.ToString();
                    list[0].month11 = list_tzys11[0].htAmount.ToString();
                    list[0].month12 = list_tzys12[0].htAmount.ToString();

                    list[1].month1 = list_fact1[0].ysAmount.ToString();
                    list[1].month2 = list_fact2[0].ysAmount.ToString();
                    list[1].month3 = list_fact3[0].ysAmount.ToString();
                    list[1].month4 = list_fact4[0].ysAmount.ToString();
                    list[1].month5 = list_xmjz5[0].ysAmount.ToString();
                    list[1].month6 = list_xmjz6[0].ysAmount.ToString();
                    list[1].month7 = list_xmjz7[0].ysAmount.ToString();
                    list[1].month8 = list_tzys8[0].ysAmount.ToString();
                    list[1].month9 = list_tzys9[0].ysAmount.ToString();
                    list[1].month10 = list_tzys10[0].ysAmount.ToString();
                    list[1].month11 = list_tzys11[0].ysAmount.ToString();
                    list[1].month12 = list_tzys12[0].ysAmount.ToString();

                    list[2].month1 = list_fact1[0].lrAmount.ToString();
                    list[2].month2 = list_fact2[0].lrAmount.ToString();
                    list[2].month3 = list_fact3[0].lrAmount.ToString();
                    list[2].month4 = list_fact4[0].lrAmount.ToString();
                    list[2].month5 = list_xmjz5[0].lrAmount.ToString();
                    list[2].month6 = list_xmjz6[0].lrAmount.ToString();
                    list[2].month7 = list_xmjz7[0].lrAmount.ToString();
                    list[2].month8 = list_tzys8[0].lrAmount.ToString();
                    list[2].month9 = list_tzys9[0].lrAmount.ToString();
                    list[2].month10 = list_tzys10[0].lrAmount.ToString();
                    list[2].month11 = list_tzys11[0].lrAmount.ToString();
                    list[2].month12 = list_tzys12[0].lrAmount.ToString();

                    list[3].month1 = list_fact1[0].yield.ToString();
                    list[3].month2 = list_fact2[0].yield.ToString();
                    list[3].month3 = list_fact3[0].yield.ToString();
                    list[3].month4 = list_fact4[0].yield.ToString();
                    list[3].month5 = list_xmjz5[0].yield.ToString();
                    list[3].month6 = list_xmjz6[0].yield.ToString();
                    list[3].month7 = list_xmjz7[0].yield.ToString();
                    list[3].month8 = list_tzys8[0].yield.ToString();
                    list[3].month9 = list_tzys9[0].yield.ToString();
                    list[3].month10 = list_tzys10[0].yield.ToString();
                    list[3].month11 = list_tzys11[0].yield.ToString();
                    list[3].month12 = list_tzys12[0].yield.ToString();

                    list[4].month1 = list_fact1[0].yieEffic.ToString();
                    list[4].month2 = list_fact2[0].yieEffic.ToString();
                    list[4].month3 = list_fact3[0].yieEffic.ToString();
                    list[4].month4 = list_fact4[0].yieEffic.ToString();
                    list[4].month5 = list_xmjz5[0].yieEffic.ToString();
                    list[4].month6 = list_xmjz6[0].yieEffic.ToString();
                    list[4].month7 = list_xmjz7[0].yieEffic.ToString();
                    list[4].month8 = list_tzys8[0].yieEffic.ToString();
                    list[4].month9 = list_tzys9[0].yieEffic.ToString();
                    list[4].month10 = list_tzys10[0].yieEffic.ToString();
                    list[4].month11 = list_tzys11[0].yieEffic.ToString();
                    list[4].month12 = list_tzys12[0].yieEffic.ToString();
                }
                else if (6 == _nowMonth)
                {
                    list[0].month1 = list_fact1[0].htAmount.ToString();
                    list[0].month2 = list_fact2[0].htAmount.ToString();
                    list[0].month3 = list_fact3[0].htAmount.ToString();
                    list[0].month4 = list_fact4[0].htAmount.ToString();
                    list[0].month5 = list_fact5[0].htAmount.ToString();
                    list[0].month6 = list_xmjz6[0].htAmount.ToString();
                    list[0].month7 = list_xmjz7[0].htAmount.ToString();
                    list[0].month8 = list_xmjz8[0].htAmount.ToString();
                    list[0].month9 = list_tzys9[0].htAmount.ToString();
                    list[0].month10 = list_tzys10[0].htAmount.ToString();
                    list[0].month11 = list_tzys11[0].htAmount.ToString();
                    list[0].month12 = list_tzys12[0].htAmount.ToString();

                    list[1].month1 = list_fact1[0].ysAmount.ToString();
                    list[1].month2 = list_fact2[0].ysAmount.ToString();
                    list[1].month3 = list_fact3[0].ysAmount.ToString();
                    list[1].month4 = list_fact4[0].ysAmount.ToString();
                    list[1].month5 = list_fact5[0].ysAmount.ToString();
                    list[1].month6 = list_xmjz6[0].ysAmount.ToString();
                    list[1].month7 = list_xmjz7[0].ysAmount.ToString();
                    list[1].month8 = list_xmjz8[0].ysAmount.ToString();
                    list[1].month9 = list_tzys9[0].ysAmount.ToString();
                    list[1].month10 = list_tzys10[0].ysAmount.ToString();
                    list[1].month11 = list_tzys11[0].ysAmount.ToString();
                    list[1].month12 = list_tzys12[0].ysAmount.ToString();

                    list[2].month1 = list_fact1[0].lrAmount.ToString();
                    list[2].month2 = list_fact2[0].lrAmount.ToString();
                    list[2].month3 = list_fact3[0].lrAmount.ToString();
                    list[2].month4 = list_fact4[0].lrAmount.ToString();
                    list[2].month5 = list_fact5[0].lrAmount.ToString();
                    list[2].month6 = list_xmjz6[0].lrAmount.ToString();
                    list[2].month7 = list_xmjz7[0].lrAmount.ToString();
                    list[2].month8 = list_xmjz8[0].lrAmount.ToString();
                    list[2].month9 = list_tzys9[0].lrAmount.ToString();
                    list[2].month10 = list_tzys10[0].lrAmount.ToString();
                    list[2].month11 = list_tzys11[0].lrAmount.ToString();
                    list[2].month12 = list_tzys12[0].lrAmount.ToString();

                    list[3].month1 = list_fact1[0].yield.ToString();
                    list[3].month2 = list_fact2[0].yield.ToString();
                    list[3].month3 = list_fact3[0].yield.ToString();
                    list[3].month4 = list_fact4[0].yield.ToString();
                    list[3].month5 = list_fact5[0].yield.ToString();
                    list[3].month6 = list_xmjz6[0].yield.ToString();
                    list[3].month7 = list_xmjz7[0].yield.ToString();
                    list[3].month8 = list_xmjz8[0].yield.ToString();
                    list[3].month9 = list_tzys9[0].yield.ToString();
                    list[3].month10 = list_tzys10[0].yield.ToString();
                    list[3].month11 = list_tzys11[0].yield.ToString();
                    list[3].month12 = list_tzys12[0].yield.ToString();

                    list[4].month1 = list_fact1[0].yieEffic.ToString();
                    list[4].month2 = list_fact2[0].yieEffic.ToString();
                    list[4].month3 = list_fact3[0].yieEffic.ToString();
                    list[4].month4 = list_fact4[0].yieEffic.ToString();
                    list[4].month5 = list_fact5[0].yieEffic.ToString();
                    list[4].month6 = list_xmjz6[0].yieEffic.ToString();
                    list[4].month7 = list_xmjz7[0].yieEffic.ToString();
                    list[4].month8 = list_xmjz8[0].yieEffic.ToString();
                    list[4].month9 = list_tzys9[0].yieEffic.ToString();
                    list[4].month10 = list_tzys10[0].yieEffic.ToString();
                    list[4].month11 = list_tzys11[0].yieEffic.ToString();
                    list[4].month12 = list_tzys12[0].yieEffic.ToString();
                }
                else if (7 == _nowMonth)
                {
                    list[0].month1 = list_fact1[0].htAmount.ToString();
                    list[0].month2 = list_fact2[0].htAmount.ToString();
                    list[0].month3 = list_fact3[0].htAmount.ToString();
                    list[0].month4 = list_fact4[0].htAmount.ToString();
                    list[0].month5 = list_fact5[0].htAmount.ToString();
                    list[0].month6 = list_fact6[0].htAmount.ToString();
                    list[0].month7 = list_xmjz7[0].htAmount.ToString();
                    list[0].month8 = list_xmjz8[0].htAmount.ToString();
                    list[0].month9 = list_xmjz9[0].htAmount.ToString();
                    list[0].month10 = list_tzys10[0].htAmount.ToString();
                    list[0].month11 = list_tzys11[0].htAmount.ToString();
                    list[0].month12 = list_tzys12[0].htAmount.ToString();

                    list[1].month1 = list_fact1[0].ysAmount.ToString();
                    list[1].month2 = list_fact2[0].ysAmount.ToString();
                    list[1].month3 = list_fact3[0].ysAmount.ToString();
                    list[1].month4 = list_fact4[0].ysAmount.ToString();
                    list[1].month5 = list_fact5[0].ysAmount.ToString();
                    list[1].month6 = list_fact6[0].ysAmount.ToString();
                    list[1].month7 = list_xmjz7[0].ysAmount.ToString();
                    list[1].month8 = list_xmjz8[0].ysAmount.ToString();
                    list[1].month9 = list_xmjz9[0].ysAmount.ToString();
                    list[1].month10 = list_tzys10[0].ysAmount.ToString();
                    list[1].month11 = list_tzys11[0].ysAmount.ToString();
                    list[1].month12 = list_tzys12[0].ysAmount.ToString();

                    list[2].month1 = list_fact1[0].lrAmount.ToString();
                    list[2].month2 = list_fact2[0].lrAmount.ToString();
                    list[2].month3 = list_fact3[0].lrAmount.ToString();
                    list[2].month4 = list_fact4[0].lrAmount.ToString();
                    list[2].month5 = list_fact5[0].lrAmount.ToString();
                    list[2].month6 = list_fact6[0].lrAmount.ToString();
                    list[2].month7 = list_xmjz7[0].lrAmount.ToString();
                    list[2].month8 = list_xmjz8[0].lrAmount.ToString();
                    list[2].month9 = list_xmjz9[0].lrAmount.ToString();
                    list[2].month10 = list_tzys10[0].lrAmount.ToString();
                    list[2].month11 = list_tzys11[0].lrAmount.ToString();
                    list[2].month12 = list_tzys12[0].lrAmount.ToString();

                    list[3].month1 = list_fact1[0].yield.ToString();
                    list[3].month2 = list_fact2[0].yield.ToString();
                    list[3].month3 = list_fact3[0].yield.ToString();
                    list[3].month4 = list_fact4[0].yield.ToString();
                    list[3].month5 = list_fact5[0].yield.ToString();
                    list[3].month6 = list_fact6[0].yield.ToString();
                    list[3].month7 = list_xmjz7[0].yield.ToString();
                    list[3].month8 = list_xmjz8[0].yield.ToString();
                    list[3].month9 = list_xmjz9[0].yield.ToString();
                    list[3].month10 = list_tzys10[0].yield.ToString();
                    list[3].month11 = list_tzys11[0].yield.ToString();
                    list[3].month12 = list_tzys12[0].yield.ToString();

                    list[4].month1 = list_fact1[0].yieEffic.ToString();
                    list[4].month2 = list_fact2[0].yieEffic.ToString();
                    list[4].month3 = list_fact3[0].yieEffic.ToString();
                    list[4].month4 = list_fact4[0].yieEffic.ToString();
                    list[4].month5 = list_fact5[0].yieEffic.ToString();
                    list[4].month6 = list_fact6[0].yieEffic.ToString();
                    list[4].month7 = list_xmjz7[0].yieEffic.ToString();
                    list[4].month8 = list_xmjz8[0].yieEffic.ToString();
                    list[4].month9 = list_xmjz9[0].yieEffic.ToString();
                    list[4].month10 = list_tzys10[0].yieEffic.ToString();
                    list[4].month11 = list_tzys11[0].yieEffic.ToString();
                    list[4].month12 = list_tzys12[0].yieEffic.ToString();
                }
                else if (8 == _nowMonth)
                {
                    list[0].month1 = list_fact1[0].htAmount.ToString();
                    list[0].month2 = list_fact2[0].htAmount.ToString();
                    list[0].month3 = list_fact3[0].htAmount.ToString();
                    list[0].month4 = list_fact4[0].htAmount.ToString();
                    list[0].month5 = list_fact5[0].htAmount.ToString();
                    list[0].month6 = list_fact6[0].htAmount.ToString();
                    list[0].month7 = list_fact7[0].htAmount.ToString();
                    list[0].month8 = list_xmjz8[0].htAmount.ToString();
                    list[0].month9 = list_xmjz9[0].htAmount.ToString();
                    list[0].month10 = list_xmjz10[0].htAmount.ToString();
                    list[0].month11 = list_tzys11[0].htAmount.ToString();
                    list[0].month12 = list_tzys12[0].htAmount.ToString();

                    list[1].month1 = list_fact1[0].ysAmount.ToString();
                    list[1].month2 = list_fact2[0].ysAmount.ToString();
                    list[1].month3 = list_fact3[0].ysAmount.ToString();
                    list[1].month4 = list_fact4[0].ysAmount.ToString();
                    list[1].month5 = list_fact5[0].ysAmount.ToString();
                    list[1].month6 = list_fact6[0].ysAmount.ToString();
                    list[1].month7 = list_fact7[0].ysAmount.ToString();
                    list[1].month8 = list_xmjz8[0].ysAmount.ToString();
                    list[1].month9 = list_xmjz9[0].ysAmount.ToString();
                    list[1].month10 = list_xmjz10[0].ysAmount.ToString();
                    list[1].month11 = list_tzys11[0].ysAmount.ToString();
                    list[1].month12 = list_tzys12[0].ysAmount.ToString();

                    list[2].month1 = list_fact1[0].lrAmount.ToString();
                    list[2].month2 = list_fact2[0].lrAmount.ToString();
                    list[2].month3 = list_fact3[0].lrAmount.ToString();
                    list[2].month4 = list_fact4[0].lrAmount.ToString();
                    list[2].month5 = list_fact5[0].lrAmount.ToString();
                    list[2].month6 = list_fact6[0].lrAmount.ToString();
                    list[2].month7 = list_fact7[0].lrAmount.ToString();
                    list[2].month8 = list_xmjz8[0].lrAmount.ToString();
                    list[2].month9 = list_xmjz9[0].lrAmount.ToString();
                    list[2].month10 = list_xmjz10[0].lrAmount.ToString();
                    list[2].month11 = list_tzys11[0].lrAmount.ToString();
                    list[2].month12 = list_tzys12[0].lrAmount.ToString();

                    list[3].month1 = list_fact1[0].yield.ToString();
                    list[3].month2 = list_fact2[0].yield.ToString();
                    list[3].month3 = list_fact3[0].yield.ToString();
                    list[3].month4 = list_fact4[0].yield.ToString();
                    list[3].month5 = list_fact5[0].yield.ToString();
                    list[3].month6 = list_fact6[0].yield.ToString();
                    list[3].month7 = list_fact7[0].yield.ToString();
                    list[3].month8 = list_xmjz8[0].yield.ToString();
                    list[3].month9 = list_xmjz9[0].yield.ToString();
                    list[3].month10 = list_xmjz10[0].yield.ToString();
                    list[3].month11 = list_tzys11[0].yield.ToString();
                    list[3].month12 = list_tzys12[0].yield.ToString();

                    list[4].month1 = list_fact1[0].yieEffic.ToString();
                    list[4].month2 = list_fact2[0].yieEffic.ToString();
                    list[4].month3 = list_fact3[0].yieEffic.ToString();
                    list[4].month4 = list_fact4[0].yieEffic.ToString();
                    list[4].month5 = list_fact5[0].yieEffic.ToString();
                    list[4].month6 = list_fact6[0].yieEffic.ToString();
                    list[4].month7 = list_fact7[0].yieEffic.ToString();
                    list[4].month8 = list_xmjz8[0].yieEffic.ToString();
                    list[4].month9 = list_xmjz9[0].yieEffic.ToString();
                    list[4].month10 = list_xmjz10[0].yieEffic.ToString();
                    list[4].month11 = list_tzys11[0].yieEffic.ToString();
                    list[4].month12 = list_tzys12[0].yieEffic.ToString();
                }
                else if (9 == _nowMonth)
                {
                    list[0].month1 = list_fact1[0].htAmount.ToString();
                    list[0].month2 = list_fact2[0].htAmount.ToString();
                    list[0].month3 = list_fact3[0].htAmount.ToString();
                    list[0].month4 = list_fact4[0].htAmount.ToString();
                    list[0].month5 = list_fact5[0].htAmount.ToString();
                    list[0].month6 = list_fact6[0].htAmount.ToString();
                    list[0].month7 = list_fact7[0].htAmount.ToString();
                    list[0].month8 = list_fact8[0].htAmount.ToString();
                    list[0].month9 = list_xmjz9[0].htAmount.ToString();
                    list[0].month10 = list_xmjz10[0].htAmount.ToString();
                    list[0].month11 = list_xmjz11[0].htAmount.ToString();
                    list[0].month12 = list_tzys12[0].htAmount.ToString();

                    list[1].month1 = list_fact1[0].ysAmount.ToString();
                    list[1].month2 = list_fact2[0].ysAmount.ToString();
                    list[1].month3 = list_fact3[0].ysAmount.ToString();
                    list[1].month4 = list_fact4[0].ysAmount.ToString();
                    list[1].month5 = list_fact5[0].ysAmount.ToString();
                    list[1].month6 = list_fact6[0].ysAmount.ToString();
                    list[1].month7 = list_fact7[0].ysAmount.ToString();
                    list[1].month8 = list_fact8[0].ysAmount.ToString();
                    list[1].month9 = list_xmjz9[0].ysAmount.ToString();
                    list[1].month10 = list_xmjz10[0].ysAmount.ToString();
                    list[1].month11 = list_xmjz11[0].ysAmount.ToString();
                    list[1].month12 = list_tzys12[0].ysAmount.ToString();

                    list[2].month1 = list_fact1[0].lrAmount.ToString();
                    list[2].month2 = list_fact2[0].lrAmount.ToString();
                    list[2].month3 = list_fact3[0].lrAmount.ToString();
                    list[2].month4 = list_fact4[0].lrAmount.ToString();
                    list[2].month5 = list_fact5[0].lrAmount.ToString();
                    list[2].month6 = list_fact6[0].lrAmount.ToString();
                    list[2].month7 = list_fact7[0].lrAmount.ToString();
                    list[2].month8 = list_fact8[0].lrAmount.ToString();
                    list[2].month9 = list_xmjz9[0].lrAmount.ToString();
                    list[2].month10 = list_xmjz10[0].lrAmount.ToString();
                    list[2].month11 = list_xmjz11[0].lrAmount.ToString();
                    list[2].month12 = list_tzys12[0].lrAmount.ToString();

                    list[3].month1 = list_fact1[0].yield.ToString();
                    list[3].month2 = list_fact2[0].yield.ToString();
                    list[3].month3 = list_fact3[0].yield.ToString();
                    list[3].month4 = list_fact4[0].yield.ToString();
                    list[3].month5 = list_fact5[0].yield.ToString();
                    list[3].month6 = list_fact6[0].yield.ToString();
                    list[3].month7 = list_fact7[0].yield.ToString();
                    list[3].month8 = list_fact8[0].yield.ToString();
                    list[3].month9 = list_xmjz9[0].yield.ToString();
                    list[3].month10 = list_xmjz10[0].yield.ToString();
                    list[3].month11 = list_xmjz11[0].yield.ToString();
                    list[3].month12 = list_tzys12[0].yield.ToString();

                    list[4].month1 = list_fact1[0].yieEffic.ToString();
                    list[4].month2 = list_fact2[0].yieEffic.ToString();
                    list[4].month3 = list_fact3[0].yieEffic.ToString();
                    list[4].month4 = list_fact4[0].yieEffic.ToString();
                    list[4].month5 = list_fact5[0].yieEffic.ToString();
                    list[4].month6 = list_fact6[0].yieEffic.ToString();
                    list[4].month7 = list_fact7[0].yieEffic.ToString();
                    list[4].month8 = list_fact8[0].yieEffic.ToString();
                    list[4].month9 = list_xmjz9[0].yieEffic.ToString();
                    list[4].month10 = list_xmjz10[0].yieEffic.ToString();
                    list[4].month11 = list_xmjz11[0].yieEffic.ToString();
                    list[4].month12 = list_tzys12[0].yieEffic.ToString();
                }
                else if (10 == _nowMonth)
                {
                    list[0].month1 = list_fact1[0].htAmount.ToString();
                    list[0].month2 = list_fact2[0].htAmount.ToString();
                    list[0].month3 = list_fact3[0].htAmount.ToString();
                    list[0].month4 = list_fact4[0].htAmount.ToString();
                    list[0].month5 = list_fact5[0].htAmount.ToString();
                    list[0].month6 = list_fact6[0].htAmount.ToString();
                    list[0].month7 = list_fact7[0].htAmount.ToString();
                    list[0].month8 = list_fact8[0].htAmount.ToString();
                    list[0].month9 = list_fact9[0].htAmount.ToString();
                    list[0].month10 = list_xmjz10[0].htAmount.ToString();
                    list[0].month11 = list_xmjz11[0].htAmount.ToString();
                    list[0].month12 = list_xmjz12[0].htAmount.ToString();

                    list[1].month1 = list_fact1[0].ysAmount.ToString();
                    list[1].month2 = list_fact2[0].ysAmount.ToString();
                    list[1].month3 = list_fact3[0].ysAmount.ToString();
                    list[1].month4 = list_fact4[0].ysAmount.ToString();
                    list[1].month5 = list_fact5[0].ysAmount.ToString();
                    list[1].month6 = list_fact6[0].ysAmount.ToString();
                    list[1].month7 = list_fact7[0].ysAmount.ToString();
                    list[1].month8 = list_fact8[0].ysAmount.ToString();
                    list[1].month9 = list_fact9[0].ysAmount.ToString();
                    list[1].month10 = list_xmjz10[0].ysAmount.ToString();
                    list[1].month11 = list_xmjz11[0].ysAmount.ToString();
                    list[1].month12 = list_xmjz12[0].ysAmount.ToString();

                    list[2].month1 = list_fact1[0].lrAmount.ToString();
                    list[2].month2 = list_fact2[0].lrAmount.ToString();
                    list[2].month3 = list_fact3[0].lrAmount.ToString();
                    list[2].month4 = list_fact4[0].lrAmount.ToString();
                    list[2].month5 = list_fact5[0].lrAmount.ToString();
                    list[2].month6 = list_fact6[0].lrAmount.ToString();
                    list[2].month7 = list_fact7[0].lrAmount.ToString();
                    list[2].month8 = list_fact8[0].lrAmount.ToString();
                    list[2].month9 = list_fact9[0].lrAmount.ToString();
                    list[2].month10 = list_xmjz10[0].lrAmount.ToString();
                    list[2].month11 = list_xmjz11[0].lrAmount.ToString();
                    list[2].month12 = list_xmjz12[0].lrAmount.ToString();

                    list[3].month1 = list_fact1[0].yield.ToString();
                    list[3].month2 = list_fact2[0].yield.ToString();
                    list[3].month3 = list_fact3[0].yield.ToString();
                    list[3].month4 = list_fact4[0].yield.ToString();
                    list[3].month5 = list_fact5[0].yield.ToString();
                    list[3].month6 = list_fact6[0].yield.ToString();
                    list[3].month7 = list_fact7[0].yield.ToString();
                    list[3].month8 = list_fact8[0].yield.ToString();
                    list[3].month9 = list_fact9[0].yield.ToString();
                    list[3].month10 = list_xmjz10[0].yield.ToString();
                    list[3].month11 = list_xmjz11[0].yield.ToString();
                    list[3].month12 = list_xmjz12[0].yield.ToString();

                    list[4].month1 = list_fact1[0].yieEffic.ToString();
                    list[4].month2 = list_fact2[0].yieEffic.ToString();
                    list[4].month3 = list_fact3[0].yieEffic.ToString();
                    list[4].month4 = list_fact4[0].yieEffic.ToString();
                    list[4].month5 = list_fact5[0].yieEffic.ToString();
                    list[4].month6 = list_fact6[0].yieEffic.ToString();
                    list[4].month7 = list_fact7[0].yieEffic.ToString();
                    list[4].month8 = list_fact8[0].yieEffic.ToString();
                    list[4].month9 = list_fact9[0].yieEffic.ToString();
                    list[4].month10 = list_xmjz10[0].yieEffic.ToString();
                    list[4].month11 = list_xmjz11[0].yieEffic.ToString();
                    list[4].month12 = list_xmjz12[0].yieEffic.ToString();
                }
                else if (11 == _nowMonth)
                {
                    list[0].month1 = list_fact1[0].htAmount.ToString();
                    list[0].month2 = list_fact2[0].htAmount.ToString();
                    list[0].month3 = list_fact3[0].htAmount.ToString();
                    list[0].month4 = list_fact4[0].htAmount.ToString();
                    list[0].month5 = list_fact5[0].htAmount.ToString();
                    list[0].month6 = list_fact6[0].htAmount.ToString();
                    list[0].month7 = list_fact7[0].htAmount.ToString();
                    list[0].month8 = list_fact8[0].htAmount.ToString();
                    list[0].month9 = list_fact9[0].htAmount.ToString();
                    list[0].month10 = list_fact10[0].htAmount.ToString();
                    list[0].month11 = list_xmjz11[0].htAmount.ToString();
                    list[0].month12 = list_xmjz12[0].htAmount.ToString();

                    list[1].month1 = list_fact1[0].ysAmount.ToString();
                    list[1].month2 = list_fact2[0].ysAmount.ToString();
                    list[1].month3 = list_fact3[0].ysAmount.ToString();
                    list[1].month4 = list_fact4[0].ysAmount.ToString();
                    list[1].month5 = list_fact5[0].ysAmount.ToString();
                    list[1].month6 = list_fact6[0].ysAmount.ToString();
                    list[1].month7 = list_fact7[0].ysAmount.ToString();
                    list[1].month8 = list_fact8[0].ysAmount.ToString();
                    list[1].month9 = list_fact9[0].ysAmount.ToString();
                    list[1].month10 = list_fact10[0].ysAmount.ToString();
                    list[1].month11 = list_xmjz11[0].ysAmount.ToString();
                    list[1].month12 = list_xmjz12[0].ysAmount.ToString();

                    list[2].month1 = list_fact1[0].lrAmount.ToString();
                    list[2].month2 = list_fact2[0].lrAmount.ToString();
                    list[2].month3 = list_fact3[0].lrAmount.ToString();
                    list[2].month4 = list_fact4[0].lrAmount.ToString();
                    list[2].month5 = list_fact5[0].lrAmount.ToString();
                    list[2].month6 = list_fact6[0].lrAmount.ToString();
                    list[2].month7 = list_fact7[0].lrAmount.ToString();
                    list[2].month8 = list_fact8[0].lrAmount.ToString();
                    list[2].month9 = list_fact9[0].lrAmount.ToString();
                    list[2].month10 = list_fact10[0].lrAmount.ToString();
                    list[2].month11 = list_xmjz11[0].lrAmount.ToString();
                    list[2].month12 = list_xmjz12[0].lrAmount.ToString();

                    list[3].month1 = list_fact1[0].yield.ToString();
                    list[3].month2 = list_fact2[0].yield.ToString();
                    list[3].month3 = list_fact3[0].yield.ToString();
                    list[3].month4 = list_fact4[0].yield.ToString();
                    list[3].month5 = list_fact5[0].yield.ToString();
                    list[3].month6 = list_fact6[0].yield.ToString();
                    list[3].month7 = list_fact7[0].yield.ToString();
                    list[3].month8 = list_fact8[0].yield.ToString();
                    list[3].month9 = list_fact9[0].yield.ToString();
                    list[3].month10 = list_fact10[0].yield.ToString();
                    list[3].month11 = list_xmjz11[0].yield.ToString();
                    list[3].month12 = list_xmjz12[0].yield.ToString();

                    list[4].month1 = list_fact1[0].yieEffic.ToString();
                    list[4].month2 = list_fact2[0].yieEffic.ToString();
                    list[4].month3 = list_fact3[0].yieEffic.ToString();
                    list[4].month4 = list_fact4[0].yieEffic.ToString();
                    list[4].month5 = list_fact5[0].yieEffic.ToString();
                    list[4].month6 = list_fact6[0].yieEffic.ToString();
                    list[4].month7 = list_fact7[0].yieEffic.ToString();
                    list[4].month8 = list_fact8[0].yieEffic.ToString();
                    list[4].month9 = list_fact9[0].yieEffic.ToString();
                    list[4].month10 = list_fact10[0].yieEffic.ToString();
                    list[4].month11 = list_xmjz11[0].yieEffic.ToString();
                    list[4].month12 = list_xmjz12[0].yieEffic.ToString();
                }
                else
                {
                    list[0].month1 = list_fact1[0].htAmount.ToString();
                    list[0].month2 = list_fact2[0].htAmount.ToString();
                    list[0].month3 = list_fact3[0].htAmount.ToString();
                    list[0].month4 = list_fact4[0].htAmount.ToString();
                    list[0].month5 = list_fact5[0].htAmount.ToString();
                    list[0].month6 = list_fact6[0].htAmount.ToString();
                    list[0].month7 = list_fact7[0].htAmount.ToString();
                    list[0].month8 = list_fact8[0].htAmount.ToString();
                    list[0].month9 = list_fact9[0].htAmount.ToString();
                    list[0].month10 = list_fact10[0].htAmount.ToString();
                    list[0].month11 = list_fact11[0].htAmount.ToString();
                    list[0].month12 = list_xmjz12[0].htAmount.ToString();

                    list[1].month1 = list_fact1[0].ysAmount.ToString();
                    list[1].month2 = list_fact2[0].ysAmount.ToString();
                    list[1].month3 = list_fact3[0].ysAmount.ToString();
                    list[1].month4 = list_fact4[0].ysAmount.ToString();
                    list[1].month5 = list_fact5[0].ysAmount.ToString();
                    list[1].month6 = list_fact6[0].ysAmount.ToString();
                    list[1].month7 = list_fact7[0].ysAmount.ToString();
                    list[1].month8 = list_fact8[0].ysAmount.ToString();
                    list[1].month9 = list_fact9[0].ysAmount.ToString();
                    list[1].month10 = list_fact10[0].ysAmount.ToString();
                    list[1].month11 = list_fact11[0].ysAmount.ToString();
                    list[1].month12 = list_xmjz12[0].ysAmount.ToString();

                    list[2].month1 = list_fact1[0].lrAmount.ToString();
                    list[2].month2 = list_fact2[0].lrAmount.ToString();
                    list[2].month3 = list_fact3[0].lrAmount.ToString();
                    list[2].month4 = list_fact4[0].lrAmount.ToString();
                    list[2].month5 = list_fact5[0].lrAmount.ToString();
                    list[2].month6 = list_fact6[0].lrAmount.ToString();
                    list[2].month7 = list_fact7[0].lrAmount.ToString();
                    list[2].month8 = list_fact8[0].lrAmount.ToString();
                    list[2].month9 = list_fact9[0].lrAmount.ToString();
                    list[2].month10 = list_fact10[0].lrAmount.ToString();
                    list[2].month11 = list_fact11[0].lrAmount.ToString();
                    list[2].month12 = list_xmjz12[0].lrAmount.ToString();

                    list[3].month1 = list_fact1[0].yield.ToString();
                    list[3].month2 = list_fact2[0].yield.ToString();
                    list[3].month3 = list_fact3[0].yield.ToString();
                    list[3].month4 = list_fact4[0].yield.ToString();
                    list[3].month5 = list_fact5[0].yield.ToString();
                    list[3].month6 = list_fact6[0].yield.ToString();
                    list[3].month7 = list_fact7[0].yield.ToString();
                    list[3].month8 = list_fact8[0].yield.ToString();
                    list[3].month9 = list_fact9[0].yield.ToString();
                    list[3].month10 = list_fact10[0].yield.ToString();
                    list[3].month11 = list_fact11[0].yield.ToString();
                    list[3].month12 = list_xmjz12[0].yield.ToString();

                    list[4].month1 = list_fact1[0].yieEffic.ToString();
                    list[4].month2 = list_fact2[0].yieEffic.ToString();
                    list[4].month3 = list_fact3[0].yieEffic.ToString();
                    list[4].month4 = list_fact4[0].yieEffic.ToString();
                    list[4].month5 = list_fact5[0].yieEffic.ToString();
                    list[4].month6 = list_fact6[0].yieEffic.ToString();
                    list[4].month7 = list_fact7[0].yieEffic.ToString();
                    list[4].month8 = list_fact8[0].yieEffic.ToString();
                    list[4].month9 = list_fact9[0].yieEffic.ToString();
                    list[4].month10 = list_fact10[0].yieEffic.ToString();
                    list[4].month11 = list_fact11[0].yieEffic.ToString();
                    list[4].month12 = list_xmjz12[0].yieEffic.ToString();
                }
                #endregion

                /*
                 * 年累实际：显示的月份的‘实际值’之和
                 * 年累预算：显示的月份的‘调整预算值’之和
                 * 年累预计：显示的月份的之和
                 */

                list[0].sj = list_fact.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.htAmount).ToString();
                list[0].ys = list_tzys.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.htAmount).ToString();
                list[1].sj = list_fact.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.ysAmount).ToString();
                list[1].ys = list_tzys.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.ysAmount).ToString();
                list[2].sj = list_fact.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.lrAmount).ToString();
                list[2].ys = list_tzys.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.lrAmount).ToString();
                list[3].sj = list_fact.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.yield).ToString();
                list[3].ys = list_tzys.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.yield).ToString();
                list[4].sj = (list_fact.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.yieEffic) / (endMonth - startMonth + 1)).ToString(); //平均值
                list[4].ys = (list_tzys.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.yieEffic) / (endMonth - startMonth + 1)).ToString();

                double _yj0 = 0;
                double _yj1 = 0;
                double _yj2 = 0;
                double _yj3 = 0;
                double _yj4 = 0;

                if (startMonth <= 1 && endMonth >= 1)
                {
                    _yj0 += double.Parse(list[0].month1);
                    _yj1 += double.Parse(list[1].month1);
                    _yj2 += double.Parse(list[2].month1);
                    _yj3 += double.Parse(list[3].month1);
                    _yj4 += double.Parse(list[4].month1);
                }
                if (startMonth <= 2 && endMonth >= 2)
                {
                    _yj0 += double.Parse(list[0].month2);
                    _yj1 += double.Parse(list[1].month2);
                    _yj2 += double.Parse(list[2].month2);
                    _yj3 += double.Parse(list[3].month2);
                    _yj4 += double.Parse(list[4].month2);
                }
                if (startMonth <= 3 && endMonth >= 3)
                {
                    _yj0 += double.Parse(list[0].month3);
                    _yj1 += double.Parse(list[1].month3);
                    _yj2 += double.Parse(list[2].month3);
                    _yj3 += double.Parse(list[3].month3);
                    _yj4 += double.Parse(list[4].month3);
                }
                if (startMonth <= 4 && endMonth >= 4)
                {
                    _yj0 += double.Parse(list[0].month4);
                    _yj1 += double.Parse(list[1].month4);
                    _yj2 += double.Parse(list[2].month4);
                    _yj3 += double.Parse(list[3].month4);
                    _yj4 += double.Parse(list[4].month4);
                }
                if (startMonth <= 5 && endMonth >= 5)
                {
                    _yj0 += double.Parse(list[0].month5);
                    _yj1 += double.Parse(list[1].month5);
                    _yj2 += double.Parse(list[2].month5);
                    _yj3 += double.Parse(list[3].month5);
                    _yj4 += double.Parse(list[4].month5);
                }
                if (startMonth <= 6 && endMonth >= 6)
                {
                    _yj0 += double.Parse(list[0].month6);
                    _yj1 += double.Parse(list[1].month6);
                    _yj2 += double.Parse(list[2].month6);
                    _yj3 += double.Parse(list[3].month6);
                    _yj4 += double.Parse(list[4].month6);
                }
                if (startMonth <= 7 && endMonth >= 7)
                {
                    _yj0 += double.Parse(list[0].month7);
                    _yj1 += double.Parse(list[1].month7);
                    _yj2 += double.Parse(list[2].month7);
                    _yj3 += double.Parse(list[3].month7);
                    _yj4 += double.Parse(list[4].month7);
                }
                if (startMonth <= 8 && endMonth >= 8)
                {
                    _yj0 += double.Parse(list[0].month8);
                    _yj1 += double.Parse(list[1].month8);
                    _yj2 += double.Parse(list[2].month8);
                    _yj3 += double.Parse(list[3].month8);
                    _yj4 += double.Parse(list[4].month8);
                }
                if (startMonth <= 9 && endMonth >= 9)
                {
                    _yj0 += double.Parse(list[0].month9);
                    _yj1 += double.Parse(list[1].month9);
                    _yj2 += double.Parse(list[2].month9);
                    _yj3 += double.Parse(list[3].month9);
                    _yj4 += double.Parse(list[4].month9);
                }
                if (startMonth <= 10 && endMonth >= 10)
                {
                    _yj0 += double.Parse(list[0].month10);
                    _yj1 += double.Parse(list[1].month10);
                    _yj2 += double.Parse(list[2].month10);
                    _yj3 += double.Parse(list[3].month10);
                    _yj4 += double.Parse(list[4].month10);
                }
                if (startMonth <= 11 && endMonth >= 11)
                {
                    _yj0 += double.Parse(list[0].month11);
                    _yj1 += double.Parse(list[1].month11);
                    _yj2 += double.Parse(list[2].month11);
                    _yj3 += double.Parse(list[3].month11);
                    _yj4 += double.Parse(list[4].month11);
                }
                if (startMonth <= 12 && endMonth == 12)
                {
                    _yj0 += double.Parse(list[0].month12);
                    _yj1 += double.Parse(list[1].month12);
                    _yj2 += double.Parse(list[2].month12);
                    _yj3 += double.Parse(list[3].month12);
                    _yj4 += double.Parse(list[4].month12);
                }

                list[0].yj = _yj0.ToString();
                list[1].yj = _yj1.ToString();
                list[2].yj = _yj2.ToString();
                list[3].yj = _yj3.ToString();
                list[4].yj = (_yj4 / (endMonth - startMonth + 1)).ToString(); //平均值

                #endregion


                #region// 人数数据
                //市场人数
                var list_htrs = new List<Hr_Midhtrs>();

                var dt_htrs = dal.GetHr_Midhtrs(start, end, mid);
                if (dt_htrs.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_htrs.Rows.Count; i++)
                    {
                        var obj2 = new Hr_Midhtrs();

                        obj2.core_tzys = int.Parse(dt_htrs.Rows[i]["core_tzys"].ToString());
                        obj2.bone_tzys = int.Parse(dt_htrs.Rows[i]["bone_tzys"].ToString());
                        obj2.monthly = int.Parse(dt_htrs.Rows[i]["monthly"].ToString());

                        list_htrs.Add(obj2);
                    }
                }

                //非市场人数
                var list_ysrs = new List<Hr_Midysrs>();

                var dt_ysrs = dal.GetHr_Midysrs(start, end, mid);
                if (dt_ysrs.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_ysrs.Rows.Count; i++)
                    {
                        var obj3 = new Hr_Midysrs();

                        obj3.coreQuota = int.Parse(dt_ysrs.Rows[i]["coreQuota"].ToString());
                        obj3.boneQuota = int.Parse(dt_ysrs.Rows[i]["boneQuota"].ToString());
                        obj3.floatQuota = int.Parse(dt_ysrs.Rows[i]["floatQuota"].ToString());
                        obj3.floattzys = int.Parse(dt_ysrs.Rows[i]["floattzys"].ToString());
                        obj3.coreActual = int.Parse(dt_ysrs.Rows[i]["coreActual"].ToString());
                        obj3.boneActual = int.Parse(dt_ysrs.Rows[i]["boneActual"].ToString());
                        obj3.floatActual = int.Parse(dt_ysrs.Rows[i]["floatActual"].ToString());
                        obj3.monthly = int.Parse(dt_ysrs.Rows[i]["monthly"].ToString());
                        obj3.postName = dt_ysrs.Rows[i]["postName"].ToString();
                        obj3.postLevel = dt_ysrs.Rows[i]["postLevel"].ToString();

                        list_ysrs.Add(obj3);
                    }
                }
                #endregion

                #region//序号14-15
                //序号14
                /* a + b
                  * 市场人数表：取公司的 调整预算核心人数之和 a
                  * 非市场人数表：取公司 核心定额人数之和 b
                  * 年度目标：取平均值（四舍五入取整数）
                  */
                list[12].classify = "其中：核心";
                list[12].yearLine = "1";
                list[12].goal = (list_htrs.Sum(p => p.core_tzys) / 12 + list_ysrs.Sum(p => p.coreQuota) / 12).ToString();
                list[12].month1 = (list_htrs.Where(p => p.monthly == 1).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 1).Sum(p => p.coreQuota)).ToString();
                list[12].month2 = (list_htrs.Where(p => p.monthly == 2).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 2).Sum(p => p.coreQuota)).ToString();
                list[12].month3 = (list_htrs.Where(p => p.monthly == 3).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 3).Sum(p => p.coreQuota)).ToString();
                list[12].month4 = (list_htrs.Where(p => p.monthly == 4).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 4).Sum(p => p.coreQuota)).ToString();
                list[12].month5 = (list_htrs.Where(p => p.monthly == 5).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 5).Sum(p => p.coreQuota)).ToString();
                list[12].month6 = (list_htrs.Where(p => p.monthly == 6).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 6).Sum(p => p.coreQuota)).ToString();
                list[12].month7 = (list_htrs.Where(p => p.monthly == 7).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 7).Sum(p => p.coreQuota)).ToString();
                list[12].month8 = (list_htrs.Where(p => p.monthly == 8).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 8).Sum(p => p.coreQuota)).ToString();
                list[12].month9 = (list_htrs.Where(p => p.monthly == 9).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 9).Sum(p => p.coreQuota)).ToString();
                list[12].month10 = (list_htrs.Where(p => p.monthly == 10).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 10).Sum(p => p.coreQuota)).ToString();
                list[12].month11 = (list_htrs.Where(p => p.monthly == 11).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 11).Sum(p => p.coreQuota)).ToString();
                list[12].month12 = (list_htrs.Where(p => p.monthly == 12).Sum(p => p.core_tzys) + list_ysrs.Where(p => p.monthly == 12).Sum(p => p.coreQuota)).ToString();

                /*
                  * 年累实际：显示的月份中有 实际值月份 的人数（月度平均值）【1月到当前月-1】
                  * 年累预计：显示的月份的人数（月度平均值）
                  * 年累预算：显示的月份的‘调整预算值’人数（月度平均值）
                  */

                double _core_tzys = list_htrs.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.core_tzys);
                double _coreQuota = list_ysrs.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.coreQuota);

                list[12].sj = ((_core_tzys+ _coreQuota) / (_nowMonth - startMonth)).ToString();

                _core_tzys = list_htrs.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.core_tzys);
                _coreQuota = list_ysrs.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.coreQuota);

                list[12].yj = ((_core_tzys + _coreQuota) / (endMonth - startMonth + 1)).ToString();
                list[12].ys = ((_core_tzys + _coreQuota) / (endMonth - startMonth + 1)).ToString();

                //序号15
                /* a + b
                  * 市场人数表：取公司的 调整预算骨干人数之和 a
                  * 非市场人数表：取公司 骨干定额人数之和 b
                  * 年度目标：取平均值（四舍五入取整数）
                  */
                list[13].classify = "骨干";
                list[13].yearLine = "1";
                list[13].goal = (list_htrs.Sum(p => p.bone_tzys) / 12 + list_ysrs.Sum(p => p.boneQuota) / 12).ToString();
                list[13].month1 = (list_htrs.Where(p => p.monthly == 1).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 1).Sum(p => p.boneQuota)).ToString();
                list[13].month2 = (list_htrs.Where(p => p.monthly == 2).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 2).Sum(p => p.boneQuota)).ToString();
                list[13].month3 = (list_htrs.Where(p => p.monthly == 3).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 3).Sum(p => p.boneQuota)).ToString();
                list[13].month4 = (list_htrs.Where(p => p.monthly == 4).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 4).Sum(p => p.boneQuota)).ToString();
                list[13].month5 = (list_htrs.Where(p => p.monthly == 5).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 5).Sum(p => p.boneQuota)).ToString();
                list[13].month6 = (list_htrs.Where(p => p.monthly == 6).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 6).Sum(p => p.boneQuota)).ToString();
                list[13].month7 = (list_htrs.Where(p => p.monthly == 7).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 7).Sum(p => p.boneQuota)).ToString();
                list[13].month8 = (list_htrs.Where(p => p.monthly == 8).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 8).Sum(p => p.boneQuota)).ToString();
                list[13].month9 = (list_htrs.Where(p => p.monthly == 9).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 9).Sum(p => p.boneQuota)).ToString();
                list[13].month10 = (list_htrs.Where(p => p.monthly == 10).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 10).Sum(p => p.boneQuota)).ToString();
                list[13].month11 = (list_htrs.Where(p => p.monthly == 11).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 11).Sum(p => p.boneQuota)).ToString();
                list[13].month12 = (list_htrs.Where(p => p.monthly == 12).Sum(p => p.bone_tzys) + list_ysrs.Where(p => p.monthly == 12).Sum(p => p.boneQuota)).ToString();

                double _bone_tzys = list_htrs.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.bone_tzys);
                double _boneQuota = list_ysrs.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.boneQuota);

                list[13].sj = ((_bone_tzys + _boneQuota) / (_nowMonth - startMonth)).ToString();

                _bone_tzys = list_htrs.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.bone_tzys);
                _boneQuota = list_ysrs.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.boneQuota);

                list[13].yj = ((_bone_tzys + _boneQuota) / (endMonth - startMonth + 1)).ToString();
                list[13].ys = ((_bone_tzys + _boneQuota) / (endMonth - startMonth + 1)).ToString();

                #endregion

                #region//序号16
                /* 
                  * 年度目标：   调整预算表：取公司的项目组数（月度平均值）
                  * 月份：       实际表：取公司当月的项目组数
                  * 年度目标：取平均值（四舍五入取整数）
                  * 月份： 岗位工资定额表：取公司当月岗位归属为核心的人数
                  */
                list[14].classify = "销售团队数";
                list[14].yearLine = "1";
                list[14].goal = (list_tzys.Sum(p => p.proTeams) / 12).ToString(); //平均值
                list[14].month1 = list_xmjz1[0].proTeams.ToString();
                list[14].month2 = list_xmjz2[0].proTeams.ToString();
                list[14].month3 = list_xmjz3[0].proTeams.ToString();
                list[14].month4 = list_xmjz4[0].proTeams.ToString();
                list[14].month5 = list_xmjz5[0].proTeams.ToString();
                list[14].month6 = list_xmjz6[0].proTeams.ToString();
                list[14].month7 = list_xmjz7[0].proTeams.ToString();
                list[14].month8 = list_xmjz8[0].proTeams.ToString();
                list[14].month9 = list_xmjz9[0].proTeams.ToString();
                list[14].month10 = list_xmjz10[0].proTeams.ToString();
                list[14].month11 = list_xmjz11[0].proTeams.ToString();
                list[14].month12 = list_xmjz12[0].proTeams.ToString();

                double _sum14 = list_fact.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.proTeams);
                list[14].sj = (_sum14 / (_nowMonth - startMonth)).ToString();
                _sum14 = list_xmjz.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.proTeams);
                list[14].yj = (_sum14 / (endMonth - startMonth + 1)).ToString();
                _sum14 = list_tzys.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.proTeams);
                list[14].yj = (_sum14 / (endMonth - startMonth + 1)).ToString();
                #endregion

                #region//序号17
                /* 
                  * 市场人数表：取公司当月的调整预算（核心 + 骨干）人数之和
                  * 年度目标：市场人数表：取公司的调整预算（核心 + 骨干）人数之和（月度平均值）
                  */
                list[15].classify = "其中：营销中心";
                list[15].yearLine = "1";
                list[15].goal = (list_htrs.Sum(p => p.core_tzys + p.bone_tzys) / 12).ToString();
                list[15].month1 = (list_htrs.Where(p => p.monthly == 1).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[15].month2 = (list_htrs.Where(p => p.monthly == 2).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[15].month3 = (list_htrs.Where(p => p.monthly == 3).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[15].month4 = (list_htrs.Where(p => p.monthly == 4).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[15].month5 = (list_htrs.Where(p => p.monthly == 5).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[15].month6 = (list_htrs.Where(p => p.monthly == 6).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[15].month7 = (list_htrs.Where(p => p.monthly == 7).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[15].month8 = (list_htrs.Where(p => p.monthly == 8).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[15].month9 = (list_htrs.Where(p => p.monthly == 9).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[15].month10 = (list_htrs.Where(p => p.monthly == 10).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[15].month11 = (list_htrs.Where(p => p.monthly == 11).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                list[15].month12 = (list_htrs.Where(p => p.monthly == 12).Sum(p => p.core_tzys + p.bone_tzys)).ToString();
                /*
                  * 年累实际：显示的月份中有实际值月份的人数（月度平均值）【1月到当前月-1】
                  * 年累预计：显示的月份的人数（月度平均值）
                  * 年累预算：显示的月份的‘调整预算值’人数（月度平均值）
                  */
                double _sum_htrs = list_htrs.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.core_tzys + p.bone_tzys);
                list[15].sj = (_sum_htrs / (_nowMonth - startMonth)).ToString();

                _sum_htrs = list_htrs.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.core_tzys + p.bone_tzys);
                list[15].yj = (_sum_htrs / (endMonth - startMonth + 1)).ToString();
                list[15].ys = (_sum_htrs / (endMonth - startMonth + 1)).ToString();
                #endregion

                #region//序号18—22
                //序号18
                /* 
                  * 月份：当月为x，月份＜x           非市场人数表：取公司岗位对应职级为A-F当月（核心定额 + 骨干定额+浮动定额）人数之和
                  *             x≤月份≤（x+2）  非市场人数表：取公司岗位对应职级为A-F当月（核心定额 + 骨干定额+浮动预测）人数之和
                  *             月份＞（x+2）     非市场人数表：取公司岗位对应职级为A-F当月（核心定额 + 骨干定额+浮动调整预算）人数之和
                  * 年度目标：非市场人数表：取公司岗位对应职级为A-F（核心定额 + 骨干定额+浮动调整预算）人数之和（月度平均值）
                  */
                var _list_ysrs16 = list_ysrs.Where(p => p.postLevel == "A" || p.postLevel == "B" || p.postLevel == "C"
                     || p.postLevel == "D" || p.postLevel == "E" || p.postLevel == "F").ToList();

                list[16].classify = "干部";
                list[16].yearLine = "1";
                list[16].goal = (_list_ysrs16.Sum(p => p.coreQuota + p.boneQuota + p.floattzys) / 12).ToString();

                //序号19
                /* 
                  * 月份：当月为x，月份＜x           非市场人数表：取公司岗位对应职级为OP当月（核心定额 + 骨干定额+浮动定额）人数之和
                  *             x≤月份≤（x+2）  非市场人数表：取公司岗位对应职级为OP当月（核心定额 + 骨干定额+浮动预测）人数之和
                  *             月份＞（x+2）     非市场人数表：取公司岗位对应职级为OP当月（核心定额 + 骨干定额+浮动调整预算）人数之和
                  * 年度目标：非市场人数表：取公司岗位对应职级为OP（核心定额 + 骨干定额+浮动调整预算）人数之和（月度平均值）
                  */
                var _list_ysrs17 = list_ysrs.Where(p => p.postLevel == "OP").ToList();

                list[17].classify = "OP";
                list[17].yearLine = "1";
                list[17].goal = (_list_ysrs17.Sum(p => p.coreQuota + p.boneQuota + p.floattzys) / 12).ToString();

                //序号20
                /* 
                  * 月份：当月为x，月份＜x           非市场人数表：取公司岗位对应职级为OO当月（核心定额 + 骨干定额+浮动定额）人数之和
                  *             x≤月份≤（x+2）  非市场人数表：取公司岗位对应职级为OO当月（核心定额 + 骨干定额+浮动预测）人数之和
                  *             月份＞（x+2）     非市场人数表：取公司岗位对应职级为OO当月（核心定额 + 骨干定额+浮动调整预算）人数之和
                  * 年度目标：非市场人数表：取公司岗位对应职级为OP（核心定额 + 骨干定额+浮动调整预算）人数之和（月度平均值）
                  */
                var _list_ysrs18 = list_ysrs.Where(p => p.postLevel == "OO").ToList();

                list[18].classify = "OO";
                list[18].yearLine = "1";
                list[18].goal = (_list_ysrs18.Sum(p => p.coreQuota + p.boneQuota + p.floattzys) / 12).ToString();

                //序号21
                /* 
                  * 月份：当月为x，月份＜x           非市场人数表：取公司BMI供应链部门中岗位对应职级为OO当月（核心定额 + 骨干定额+浮动定额）人数之和
                  *             x≤月份≤（x+2）  非市场人数表：取公司BMI供应链部门中岗位对应职级为OO当月（核心定额 + 骨干定额+浮动预测）人数之和
                  *             月份＞（x+2）     非市场人数表：取公司BMI供应链部门中岗位对应职级为OO当月（核心定额 + 骨干定额+浮动调整预算）人数之和
                  * 年度目标：非市场人数表：取公司BMI供应链部门中岗位对应职级为OO（核心定额 + 骨干定额+浮动调整预算）人数之和（月度平均值）
                  */
                var _list_ysrs19 = list_ysrs.Where(p => p.postName == "BMI供应链").ToList();

                list[19].classify = "BMI";
                list[19].yearLine = "1";
                list[19].goal = (_list_ysrs19.Sum(p => p.coreQuota + p.boneQuota + p.floattzys) / 12).ToString();

                //序号22
                /* 
                  * 月份：当月为x，月份＜x           非市场人数表：取公司BPL生产部门中岗位对应职级为OO当月（核心定额 + 骨干定额+浮动定额）人数之和
                  *             x≤月份≤（x+2）  非市场人数表：取公司BPL生产部门中岗位对应职级为OO当月（核心定额 + 骨干定额+浮动预测）人数之和
                  *             月份＞（x+2）     非市场人数表：取公司BPL生产部门中岗位对应职级为OO当月（核心定额 + 骨干定额+浮动调整预算）人数之和
                  * 年度目标：非市场人数表：取公司BPL生产部门中岗位对应职级为OO（核心定额 + 骨干定额+浮动调整预算）人数之和（月度平均值）
                  */
                var _list_ysrs20 = list_ysrs.Where(p => p.postName == "BPL生产").ToList();

                list[20].classify = "BPL";
                list[20].yearLine = "1";
                list[20].goal = (_list_ysrs20.Sum(p => p.coreQuota + p.boneQuota + p.floattzys) / 12).ToString();

                #region //月份
                switch (_nowMonth)
                {
                    case 1:
                        list[16].month1 = (_list_ysrs16.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month2 = (_list_ysrs16.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month3 = (_list_ysrs16.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month4 = (_list_ysrs16.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month5 = (_list_ysrs16.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month6 = (_list_ysrs16.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month7 = (_list_ysrs16.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month8 = (_list_ysrs16.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month9 = (_list_ysrs16.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month10 = (_list_ysrs16.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month11 = (_list_ysrs16.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month12 = (_list_ysrs16.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[17].month1 = (_list_ysrs17.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month2 = (_list_ysrs17.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month3 = (_list_ysrs17.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month4 = (_list_ysrs17.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month5 = (_list_ysrs17.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month6 = (_list_ysrs17.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month7 = (_list_ysrs17.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month8 = (_list_ysrs17.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month9 = (_list_ysrs17.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month10 = (_list_ysrs17.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month11 = (_list_ysrs17.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month12 = (_list_ysrs17.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[18].month1 = (_list_ysrs18.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month2 = (_list_ysrs18.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month3 = (_list_ysrs18.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month4 = (_list_ysrs18.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month5 = (_list_ysrs18.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month6 = (_list_ysrs18.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month7 = (_list_ysrs18.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month8 = (_list_ysrs18.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month9 = (_list_ysrs18.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month10 = (_list_ysrs18.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month11 = (_list_ysrs18.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month12 = (_list_ysrs18.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[19].month1 = (_list_ysrs19.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month2 = (_list_ysrs19.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month3 = (_list_ysrs19.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month4 = (_list_ysrs19.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month5 = (_list_ysrs19.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month6 = (_list_ysrs19.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month7 = (_list_ysrs19.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month8 = (_list_ysrs19.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month9 = (_list_ysrs19.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month10 = (_list_ysrs19.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month11 = (_list_ysrs19.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month12 = (_list_ysrs19.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[20].month1 = (_list_ysrs20.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month2 = (_list_ysrs20.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month3 = (_list_ysrs20.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month4 = (_list_ysrs20.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month5 = (_list_ysrs20.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month6 = (_list_ysrs20.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month7 = (_list_ysrs20.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month8 = (_list_ysrs20.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month9 = (_list_ysrs20.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month10 = (_list_ysrs20.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month11 = (_list_ysrs20.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month12 = (_list_ysrs20.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        break;
                    case 2:
                        list[16].month1 = (_list_ysrs16.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month2 = (_list_ysrs16.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month3 = (_list_ysrs16.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month4 = (_list_ysrs16.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month5 = (_list_ysrs16.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month6 = (_list_ysrs16.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month7 = (_list_ysrs16.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month8 = (_list_ysrs16.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month9 = (_list_ysrs16.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month10 = (_list_ysrs16.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month11 = (_list_ysrs16.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month12 = (_list_ysrs16.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[17].month1 = (_list_ysrs17.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month2 = (_list_ysrs17.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month3 = (_list_ysrs17.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month4 = (_list_ysrs17.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month5 = (_list_ysrs17.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month6 = (_list_ysrs17.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month7 = (_list_ysrs17.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month8 = (_list_ysrs17.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month9 = (_list_ysrs17.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month10 = (_list_ysrs17.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month11 = (_list_ysrs17.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month12 = (_list_ysrs17.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[18].month1 = (_list_ysrs18.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month2 = (_list_ysrs18.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month3 = (_list_ysrs18.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month4 = (_list_ysrs18.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month5 = (_list_ysrs18.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month6 = (_list_ysrs18.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month7 = (_list_ysrs18.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month8 = (_list_ysrs18.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month9 = (_list_ysrs18.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month10 = (_list_ysrs18.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month11 = (_list_ysrs18.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month12 = (_list_ysrs18.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[19].month1 = (_list_ysrs19.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month2 = (_list_ysrs19.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month3 = (_list_ysrs19.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month4 = (_list_ysrs19.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month5 = (_list_ysrs19.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month6 = (_list_ysrs19.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month7 = (_list_ysrs19.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month8 = (_list_ysrs19.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month9 = (_list_ysrs19.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month10 = (_list_ysrs19.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month11 = (_list_ysrs19.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month12 = (_list_ysrs19.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[20].month1 = (_list_ysrs20.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month2 = (_list_ysrs20.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month3 = (_list_ysrs20.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month4 = (_list_ysrs20.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month5 = (_list_ysrs20.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month6 = (_list_ysrs20.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month7 = (_list_ysrs20.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month8 = (_list_ysrs20.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month9 = (_list_ysrs20.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month10 = (_list_ysrs20.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month11 = (_list_ysrs20.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month12 = (_list_ysrs20.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        break;
                    case 3:
                        list[16].month1 = (_list_ysrs16.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month2 = (_list_ysrs16.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month3 = (_list_ysrs16.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month4 = (_list_ysrs16.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month5 = (_list_ysrs16.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month6 = (_list_ysrs16.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month7 = (_list_ysrs16.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month8 = (_list_ysrs16.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month9 = (_list_ysrs16.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month10 = (_list_ysrs16.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month11 = (_list_ysrs16.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month12 = (_list_ysrs16.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[17].month1 = (_list_ysrs17.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month2 = (_list_ysrs17.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month3 = (_list_ysrs17.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month4 = (_list_ysrs17.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month5 = (_list_ysrs17.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month6 = (_list_ysrs17.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month7 = (_list_ysrs17.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month8 = (_list_ysrs17.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month9 = (_list_ysrs17.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month10 = (_list_ysrs17.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month11 = (_list_ysrs17.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month12 = (_list_ysrs17.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[18].month1 = (_list_ysrs18.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month2 = (_list_ysrs18.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month3 = (_list_ysrs18.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month4 = (_list_ysrs18.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month5 = (_list_ysrs18.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month6 = (_list_ysrs18.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month7 = (_list_ysrs18.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month8 = (_list_ysrs18.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month9 = (_list_ysrs18.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month10 = (_list_ysrs18.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month11 = (_list_ysrs18.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month12 = (_list_ysrs18.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[19].month1 = (_list_ysrs19.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month2 = (_list_ysrs19.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month3 = (_list_ysrs19.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month4 = (_list_ysrs19.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month5 = (_list_ysrs19.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month6 = (_list_ysrs19.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month7 = (_list_ysrs19.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month8 = (_list_ysrs19.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month9 = (_list_ysrs19.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month10 = (_list_ysrs19.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month11 = (_list_ysrs19.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month12 = (_list_ysrs19.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[20].month1 = (_list_ysrs20.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month2 = (_list_ysrs20.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month3 = (_list_ysrs20.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month4 = (_list_ysrs20.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month5 = (_list_ysrs20.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month6 = (_list_ysrs20.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month7 = (_list_ysrs20.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month8 = (_list_ysrs20.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month9 = (_list_ysrs20.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month10 = (_list_ysrs20.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month11 = (_list_ysrs20.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month12 = (_list_ysrs20.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        break;
                    case 4:
                        list[16].month1 = (_list_ysrs16.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month2 = (_list_ysrs16.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month3 = (_list_ysrs16.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month4 = (_list_ysrs16.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month5 = (_list_ysrs16.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month6 = (_list_ysrs16.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month7 = (_list_ysrs16.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month8 = (_list_ysrs16.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month9 = (_list_ysrs16.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month10 = (_list_ysrs16.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month11 = (_list_ysrs16.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month12 = (_list_ysrs16.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[17].month1 = (_list_ysrs17.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month2 = (_list_ysrs17.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month3 = (_list_ysrs17.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month4 = (_list_ysrs17.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month5 = (_list_ysrs17.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month6 = (_list_ysrs17.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month7 = (_list_ysrs17.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month8 = (_list_ysrs17.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month9 = (_list_ysrs17.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month10 = (_list_ysrs17.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month11 = (_list_ysrs17.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month12 = (_list_ysrs17.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[18].month1 = (_list_ysrs18.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month2 = (_list_ysrs18.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month3 = (_list_ysrs18.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month4 = (_list_ysrs18.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month5 = (_list_ysrs18.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month6 = (_list_ysrs18.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month7 = (_list_ysrs18.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month8 = (_list_ysrs18.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month9 = (_list_ysrs18.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month10 = (_list_ysrs18.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month11 = (_list_ysrs18.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month12 = (_list_ysrs18.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[19].month1 = (_list_ysrs19.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month2 = (_list_ysrs19.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month3 = (_list_ysrs19.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month4 = (_list_ysrs19.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month5 = (_list_ysrs19.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month6 = (_list_ysrs19.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month7 = (_list_ysrs19.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month8 = (_list_ysrs19.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month9 = (_list_ysrs19.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month10 = (_list_ysrs19.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month11 = (_list_ysrs19.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month12 = (_list_ysrs19.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[20].month1 = (_list_ysrs20.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month2 = (_list_ysrs20.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month3 = (_list_ysrs20.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month4 = (_list_ysrs20.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month5 = (_list_ysrs20.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month6 = (_list_ysrs20.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month7 = (_list_ysrs20.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month8 = (_list_ysrs20.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month9 = (_list_ysrs20.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month10 = (_list_ysrs20.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month11 = (_list_ysrs20.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month12 = (_list_ysrs20.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        break;
                    case 5:
                        list[16].month1 = (_list_ysrs16.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month2 = (_list_ysrs16.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month3 = (_list_ysrs16.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month4 = (_list_ysrs16.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month5 = (_list_ysrs16.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month6 = (_list_ysrs16.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month7 = (_list_ysrs16.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month8 = (_list_ysrs16.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month9 = (_list_ysrs16.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month10 = (_list_ysrs16.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month11 = (_list_ysrs16.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month12 = (_list_ysrs16.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[17].month1 = (_list_ysrs17.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month2 = (_list_ysrs17.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month3 = (_list_ysrs17.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month4 = (_list_ysrs17.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month5 = (_list_ysrs17.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month6 = (_list_ysrs17.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month7 = (_list_ysrs17.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month8 = (_list_ysrs17.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month9 = (_list_ysrs17.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month10 = (_list_ysrs17.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month11 = (_list_ysrs17.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month12 = (_list_ysrs17.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[18].month1 = (_list_ysrs18.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month2 = (_list_ysrs18.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month3 = (_list_ysrs18.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month4 = (_list_ysrs18.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month5 = (_list_ysrs18.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month6 = (_list_ysrs18.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month7 = (_list_ysrs18.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month8 = (_list_ysrs18.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month9 = (_list_ysrs18.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month10 = (_list_ysrs18.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month11 = (_list_ysrs18.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month12 = (_list_ysrs18.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[19].month1 = (_list_ysrs19.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month2 = (_list_ysrs19.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month3 = (_list_ysrs19.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month4 = (_list_ysrs19.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month5 = (_list_ysrs19.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month6 = (_list_ysrs19.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month7 = (_list_ysrs19.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month8 = (_list_ysrs19.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month9 = (_list_ysrs19.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month10 = (_list_ysrs19.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month11 = (_list_ysrs19.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month12 = (_list_ysrs19.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[20].month1 = (_list_ysrs20.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month2 = (_list_ysrs20.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month3 = (_list_ysrs20.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month4 = (_list_ysrs20.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month5 = (_list_ysrs20.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month6 = (_list_ysrs20.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month7 = (_list_ysrs20.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month8 = (_list_ysrs20.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month9 = (_list_ysrs20.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month10 = (_list_ysrs20.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month11 = (_list_ysrs20.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month12 = (_list_ysrs20.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        break;
                    case 6:
                        list[16].month1 = (_list_ysrs16.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month2 = (_list_ysrs16.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month3 = (_list_ysrs16.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month4 = (_list_ysrs16.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month5 = (_list_ysrs16.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month6 = (_list_ysrs16.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month7 = (_list_ysrs16.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month8 = (_list_ysrs16.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month9 = (_list_ysrs16.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month10 = (_list_ysrs16.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month11 = (_list_ysrs16.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month12 = (_list_ysrs16.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[17].month1 = (_list_ysrs17.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month2 = (_list_ysrs17.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month3 = (_list_ysrs17.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month4 = (_list_ysrs17.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month5 = (_list_ysrs17.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month6 = (_list_ysrs17.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month7 = (_list_ysrs17.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month8 = (_list_ysrs17.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month9 = (_list_ysrs17.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month10 = (_list_ysrs17.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month11 = (_list_ysrs17.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month12 = (_list_ysrs17.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[18].month1 = (_list_ysrs18.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month2 = (_list_ysrs18.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month3 = (_list_ysrs18.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month4 = (_list_ysrs18.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month5 = (_list_ysrs18.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month6 = (_list_ysrs18.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month7 = (_list_ysrs18.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month8 = (_list_ysrs18.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month9 = (_list_ysrs18.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month10 = (_list_ysrs18.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month11 = (_list_ysrs18.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month12 = (_list_ysrs18.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[19].month1 = (_list_ysrs19.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month2 = (_list_ysrs19.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month3 = (_list_ysrs19.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month4 = (_list_ysrs19.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month5 = (_list_ysrs19.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month6 = (_list_ysrs19.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month7 = (_list_ysrs19.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month8 = (_list_ysrs19.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month9 = (_list_ysrs19.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month10 = (_list_ysrs19.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month11 = (_list_ysrs19.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month12 = (_list_ysrs19.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[20].month1 = (_list_ysrs20.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month2 = (_list_ysrs20.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month3 = (_list_ysrs20.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month4 = (_list_ysrs20.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month5 = (_list_ysrs20.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month6 = (_list_ysrs20.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month7 = (_list_ysrs20.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month8 = (_list_ysrs20.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month9 = (_list_ysrs20.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month10 = (_list_ysrs20.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month11 = (_list_ysrs20.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month12 = (_list_ysrs20.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        break;
                    case 7:
                        list[16].month1 = (_list_ysrs16.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month2 = (_list_ysrs16.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month3 = (_list_ysrs16.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month4 = (_list_ysrs16.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month5 = (_list_ysrs16.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month6 = (_list_ysrs16.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month7 = (_list_ysrs16.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month8 = (_list_ysrs16.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month9 = (_list_ysrs16.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month10 = (_list_ysrs16.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month11 = (_list_ysrs16.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month12 = (_list_ysrs16.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[17].month1 = (_list_ysrs17.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month2 = (_list_ysrs17.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month3 = (_list_ysrs17.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month4 = (_list_ysrs17.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month5 = (_list_ysrs17.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month6 = (_list_ysrs17.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month7 = (_list_ysrs17.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month8 = (_list_ysrs17.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month9 = (_list_ysrs17.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month10 = (_list_ysrs17.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month11 = (_list_ysrs17.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month12 = (_list_ysrs17.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[18].month1 = (_list_ysrs18.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month2 = (_list_ysrs18.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month3 = (_list_ysrs18.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month4 = (_list_ysrs18.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month5 = (_list_ysrs18.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month6 = (_list_ysrs18.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month7 = (_list_ysrs18.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month8 = (_list_ysrs18.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month9 = (_list_ysrs18.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month10 = (_list_ysrs18.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month11 = (_list_ysrs18.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month12 = (_list_ysrs18.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[19].month1 = (_list_ysrs19.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month2 = (_list_ysrs19.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month3 = (_list_ysrs19.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month4 = (_list_ysrs19.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month5 = (_list_ysrs19.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month6 = (_list_ysrs19.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month7 = (_list_ysrs19.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month8 = (_list_ysrs19.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month9 = (_list_ysrs19.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month10 = (_list_ysrs19.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month11 = (_list_ysrs19.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month12 = (_list_ysrs19.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[20].month1 = (_list_ysrs20.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month2 = (_list_ysrs20.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month3 = (_list_ysrs20.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month4 = (_list_ysrs20.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month5 = (_list_ysrs20.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month6 = (_list_ysrs20.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month7 = (_list_ysrs20.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month8 = (_list_ysrs20.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month9 = (_list_ysrs20.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month10 = (_list_ysrs20.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month11 = (_list_ysrs20.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month12 = (_list_ysrs20.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        break;
                    case 8:
                        list[16].month1 = (_list_ysrs16.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month2 = (_list_ysrs16.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month3 = (_list_ysrs16.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month4 = (_list_ysrs16.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month5 = (_list_ysrs16.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month6 = (_list_ysrs16.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month7 = (_list_ysrs16.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month8 = (_list_ysrs16.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month9 = (_list_ysrs16.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month10 = (_list_ysrs16.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month11 = (_list_ysrs16.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[16].month12 = (_list_ysrs16.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[17].month1 = (_list_ysrs17.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month2 = (_list_ysrs17.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month3 = (_list_ysrs17.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month4 = (_list_ysrs17.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month5 = (_list_ysrs17.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month6 = (_list_ysrs17.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month7 = (_list_ysrs17.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month8 = (_list_ysrs17.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month9 = (_list_ysrs17.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month10 = (_list_ysrs17.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month11 = (_list_ysrs17.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[17].month12 = (_list_ysrs17.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[18].month1 = (_list_ysrs18.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month2 = (_list_ysrs18.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month3 = (_list_ysrs18.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month4 = (_list_ysrs18.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month5 = (_list_ysrs18.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month6 = (_list_ysrs18.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month7 = (_list_ysrs18.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month8 = (_list_ysrs18.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month9 = (_list_ysrs18.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month10 = (_list_ysrs18.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month11 = (_list_ysrs18.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[18].month12 = (_list_ysrs18.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[19].month1 = (_list_ysrs19.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month2 = (_list_ysrs19.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month3 = (_list_ysrs19.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month4 = (_list_ysrs19.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month5 = (_list_ysrs19.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month6 = (_list_ysrs19.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month7 = (_list_ysrs19.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month8 = (_list_ysrs19.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month9 = (_list_ysrs19.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month10 = (_list_ysrs19.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month11 = (_list_ysrs19.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[19].month12 = (_list_ysrs19.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[20].month1 = (_list_ysrs20.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month2 = (_list_ysrs20.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month3 = (_list_ysrs20.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month4 = (_list_ysrs20.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month5 = (_list_ysrs20.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month6 = (_list_ysrs20.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month7 = (_list_ysrs20.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month8 = (_list_ysrs20.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month9 = (_list_ysrs20.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month10 = (_list_ysrs20.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month11 = (_list_ysrs20.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        list[20].month12 = (_list_ysrs20.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        break;
                    case 9:
                        list[16].month1 = (_list_ysrs16.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month2 = (_list_ysrs16.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month3 = (_list_ysrs16.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month4 = (_list_ysrs16.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month5 = (_list_ysrs16.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month6 = (_list_ysrs16.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month7 = (_list_ysrs16.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month8 = (_list_ysrs16.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month9 = (_list_ysrs16.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month10 = (_list_ysrs16.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month11 = (_list_ysrs16.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month12 = (_list_ysrs16.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[17].month1 = (_list_ysrs17.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month2 = (_list_ysrs17.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month3 = (_list_ysrs17.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month4 = (_list_ysrs17.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month5 = (_list_ysrs17.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month6 = (_list_ysrs17.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month7 = (_list_ysrs17.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month8 = (_list_ysrs17.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month9 = (_list_ysrs17.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month10 = (_list_ysrs17.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month11 = (_list_ysrs17.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month12 = (_list_ysrs17.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[18].month1 = (_list_ysrs18.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month2 = (_list_ysrs18.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month3 = (_list_ysrs18.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month4 = (_list_ysrs18.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month5 = (_list_ysrs18.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month6 = (_list_ysrs18.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month7 = (_list_ysrs18.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month8 = (_list_ysrs18.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month9 = (_list_ysrs18.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month10 = (_list_ysrs18.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month11 = (_list_ysrs18.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month12 = (_list_ysrs18.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[19].month1 = (_list_ysrs19.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month2 = (_list_ysrs19.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month3 = (_list_ysrs19.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month4 = (_list_ysrs19.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month5 = (_list_ysrs19.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month6 = (_list_ysrs19.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month7 = (_list_ysrs19.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month8 = (_list_ysrs19.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month9 = (_list_ysrs19.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month10 = (_list_ysrs19.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month11 = (_list_ysrs19.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month12 = (_list_ysrs19.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();

                        list[20].month1 = (_list_ysrs20.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month2 = (_list_ysrs20.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month3 = (_list_ysrs20.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month4 = (_list_ysrs20.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month5 = (_list_ysrs20.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month6 = (_list_ysrs20.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month7 = (_list_ysrs20.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month8 = (_list_ysrs20.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month9 = (_list_ysrs20.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month10 = (_list_ysrs20.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month11 = (_list_ysrs20.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month12 = (_list_ysrs20.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floattzys)).ToString();
                        break;
                    case 10:
                        list[16].month1 = (_list_ysrs16.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month2 = (_list_ysrs16.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month3 = (_list_ysrs16.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month4 = (_list_ysrs16.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month5 = (_list_ysrs16.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month6 = (_list_ysrs16.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month7 = (_list_ysrs16.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month8 = (_list_ysrs16.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month9 = (_list_ysrs16.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month10 = (_list_ysrs16.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month11 = (_list_ysrs16.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month12 = (_list_ysrs16.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();

                        list[17].month1 = (_list_ysrs17.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month2 = (_list_ysrs17.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month3 = (_list_ysrs17.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month4 = (_list_ysrs17.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month5 = (_list_ysrs17.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month6 = (_list_ysrs17.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month7 = (_list_ysrs17.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month8 = (_list_ysrs17.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month9 = (_list_ysrs17.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month10 = (_list_ysrs17.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month11 = (_list_ysrs17.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month12 = (_list_ysrs17.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();

                        list[18].month1 = (_list_ysrs18.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month2 = (_list_ysrs18.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month3 = (_list_ysrs18.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month4 = (_list_ysrs18.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month5 = (_list_ysrs18.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month6 = (_list_ysrs18.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month7 = (_list_ysrs18.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month8 = (_list_ysrs18.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month9 = (_list_ysrs18.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month10 = (_list_ysrs18.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month11 = (_list_ysrs18.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month12 = (_list_ysrs18.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();

                        list[19].month1 = (_list_ysrs19.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month2 = (_list_ysrs19.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month3 = (_list_ysrs19.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month4 = (_list_ysrs19.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month5 = (_list_ysrs19.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month6 = (_list_ysrs19.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month7 = (_list_ysrs19.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month8 = (_list_ysrs19.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month9 = (_list_ysrs19.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month10 = (_list_ysrs19.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month11 = (_list_ysrs19.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month12 = (_list_ysrs19.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();

                        list[20].month1 = (_list_ysrs20.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month2 = (_list_ysrs20.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month3 = (_list_ysrs20.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month4 = (_list_ysrs20.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month5 = (_list_ysrs20.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month6 = (_list_ysrs20.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month7 = (_list_ysrs20.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month8 = (_list_ysrs20.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month9 = (_list_ysrs20.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month10 = (_list_ysrs20.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month11 = (_list_ysrs20.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month12 = (_list_ysrs20.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        break;
                    case 11:
                        list[16].month1 = (_list_ysrs16.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month2 = (_list_ysrs16.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month3 = (_list_ysrs16.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month4 = (_list_ysrs16.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month5 = (_list_ysrs16.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month6 = (_list_ysrs16.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month7 = (_list_ysrs16.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month8 = (_list_ysrs16.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month9 = (_list_ysrs16.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month10 = (_list_ysrs16.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month11 = (_list_ysrs16.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[16].month12 = (_list_ysrs16.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();

                        list[17].month1 = (_list_ysrs17.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month2 = (_list_ysrs17.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month3 = (_list_ysrs17.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month4 = (_list_ysrs17.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month5 = (_list_ysrs17.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month6 = (_list_ysrs17.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month7 = (_list_ysrs17.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month8 = (_list_ysrs17.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month9 = (_list_ysrs17.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month10 = (_list_ysrs17.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month11 = (_list_ysrs17.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[17].month12 = (_list_ysrs17.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();

                        list[18].month1 = (_list_ysrs18.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month2 = (_list_ysrs18.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month3 = (_list_ysrs18.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month4 = (_list_ysrs18.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month5 = (_list_ysrs18.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month6 = (_list_ysrs18.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month7 = (_list_ysrs18.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month8 = (_list_ysrs18.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month9 = (_list_ysrs18.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month10 = (_list_ysrs18.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month11 = (_list_ysrs18.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[18].month12 = (_list_ysrs18.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();

                        list[19].month1 = (_list_ysrs19.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month2 = (_list_ysrs19.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month3 = (_list_ysrs19.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month4 = (_list_ysrs19.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month5 = (_list_ysrs19.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month6 = (_list_ysrs19.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month7 = (_list_ysrs19.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month8 = (_list_ysrs19.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month9 = (_list_ysrs19.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month10 = (_list_ysrs19.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month11 = (_list_ysrs19.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[19].month12 = (_list_ysrs19.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();

                        list[20].month1 = (_list_ysrs20.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month2 = (_list_ysrs20.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month3 = (_list_ysrs20.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month4 = (_list_ysrs20.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month5 = (_list_ysrs20.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month6 = (_list_ysrs20.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month7 = (_list_ysrs20.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month8 = (_list_ysrs20.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month9 = (_list_ysrs20.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month10 = (_list_ysrs20.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month11 = (_list_ysrs20.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        list[20].month12 = (_list_ysrs20.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        break;
                    case 12:
                        list[16].month1 = (_list_ysrs16.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month2 = (_list_ysrs16.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month3 = (_list_ysrs16.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month4 = (_list_ysrs16.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month5 = (_list_ysrs16.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month6 = (_list_ysrs16.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month7 = (_list_ysrs16.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month8 = (_list_ysrs16.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month9 = (_list_ysrs16.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month10 = (_list_ysrs16.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month11 = (_list_ysrs16.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[16].month12 = (_list_ysrs16.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();

                        list[17].month1 = (_list_ysrs17.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month2 = (_list_ysrs17.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month3 = (_list_ysrs17.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month4 = (_list_ysrs17.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month5 = (_list_ysrs17.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month6 = (_list_ysrs17.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month7 = (_list_ysrs17.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month8 = (_list_ysrs17.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month9 = (_list_ysrs17.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month10 = (_list_ysrs17.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month11 = (_list_ysrs17.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[17].month12 = (_list_ysrs17.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();

                        list[18].month1 = (_list_ysrs18.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month2 = (_list_ysrs18.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month3 = (_list_ysrs18.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month4 = (_list_ysrs18.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month5 = (_list_ysrs18.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month6 = (_list_ysrs18.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month7 = (_list_ysrs18.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month8 = (_list_ysrs18.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month9 = (_list_ysrs18.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month10 = (_list_ysrs18.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month11 = (_list_ysrs18.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[18].month12 = (_list_ysrs18.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();

                        list[19].month1 = (_list_ysrs19.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month2 = (_list_ysrs19.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month3 = (_list_ysrs19.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month4 = (_list_ysrs19.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month5 = (_list_ysrs19.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month6 = (_list_ysrs19.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month7 = (_list_ysrs19.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month8 = (_list_ysrs19.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month9 = (_list_ysrs19.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month10 = (_list_ysrs19.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month11 = (_list_ysrs19.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[19].month12 = (_list_ysrs19.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();

                        list[20].month1 = (_list_ysrs20.Where(p => p.monthly == 1).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month2 = (_list_ysrs20.Where(p => p.monthly == 2).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month3 = (_list_ysrs20.Where(p => p.monthly == 3).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month4 = (_list_ysrs20.Where(p => p.monthly == 4).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month5 = (_list_ysrs20.Where(p => p.monthly == 5).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month6 = (_list_ysrs20.Where(p => p.monthly == 6).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month7 = (_list_ysrs20.Where(p => p.monthly == 7).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month8 = (_list_ysrs20.Where(p => p.monthly == 8).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month9 = (_list_ysrs20.Where(p => p.monthly == 9).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month10 = (_list_ysrs20.Where(p => p.monthly == 10).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month11 = (_list_ysrs20.Where(p => p.monthly == 11).Sum(p => p.coreQuota + p.boneQuota + p.floatActual)).ToString();
                        list[20].month12 = (_list_ysrs20.Where(p => p.monthly == 12).Sum(p => p.coreQuota + p.boneQuota + p.floatFore)).ToString();
                        break;

                }
                #endregion

                /*
                  * 年累实际：显示的月份中有实际值月份的人数（月度平均值）【1月到当前月-1】
                  * 年累预计：显示的月份的人数（月度平均值）
                  * 年累预算：显示的月份的‘调整预算值’人数（月度平均值）
                  */
                double _sum16 = _list_ysrs16.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.coreQuota + p.boneQuota + p.floatActual);
                list[16].sj = (_sum16 / (_nowMonth - startMonth)).ToString();

                _sum16 += _list_ysrs16.Where(p => p.monthly >= _nowMonth && p.monthly <= (_nowMonth+2)).Sum(p => p.coreQuota + p.boneQuota + p.floatFore);
                _sum16 += _list_ysrs16.Where(p => p.monthly >= (_nowMonth + 2) && p.monthly <= endMonth).Sum(p => p.coreQuota + p.boneQuota + p.floattzys);
                list[16].yj = (_sum16 / (endMonth - startMonth + 1)).ToString();

                _sum16 = _list_ysrs16.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.coreQuota + p.boneQuota + p.floattzys);
                list[16].ys = (_sum16 / (endMonth - startMonth + 1)).ToString();

                
                double _sum17 = _list_ysrs17.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.coreQuota + p.boneQuota + p.floatActual);
                list[17].sj = (_sum17 / (_nowMonth - startMonth)).ToString();

                _sum17 += _list_ysrs17.Where(p => p.monthly >= _nowMonth && p.monthly <= (_nowMonth + 2)).Sum(p => p.coreQuota + p.boneQuota + p.floatFore);
                _sum17 += _list_ysrs17.Where(p => p.monthly >= (_nowMonth + 2) && p.monthly <= endMonth).Sum(p => p.coreQuota + p.boneQuota + p.floattzys);
                list[17].yj = (_sum17 / (endMonth - startMonth + 1)).ToString();

                _sum17 = _list_ysrs17.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.coreQuota + p.boneQuota + p.floattzys);
                list[17].ys = (_sum17 / (endMonth - startMonth + 1)).ToString();


                double _sum18 = _list_ysrs18.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.coreQuota + p.boneQuota + p.floatActual);
                list[18].sj = (_sum18 / (_nowMonth - startMonth)).ToString();

                _sum18 += _list_ysrs18.Where(p => p.monthly >= _nowMonth && p.monthly <= (_nowMonth + 2)).Sum(p => p.coreQuota + p.boneQuota + p.floatFore);
                _sum18 += _list_ysrs18.Where(p => p.monthly >= (_nowMonth + 2) && p.monthly <= endMonth).Sum(p => p.coreQuota + p.boneQuota + p.floattzys);
                list[18].yj = (_sum18 / (endMonth - startMonth + 1)).ToString();

                _sum18 = _list_ysrs18.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.coreQuota + p.boneQuota + p.floattzys);
                list[18].ys = (_sum18 / (endMonth - startMonth + 1)).ToString();


                double _sum19 = _list_ysrs19.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.coreQuota + p.boneQuota + p.floatActual);
                list[19].sj = (_sum19 / (_nowMonth - startMonth)).ToString();

                _sum19 += _list_ysrs19.Where(p => p.monthly >= _nowMonth && p.monthly <= (_nowMonth + 2)).Sum(p => p.coreQuota + p.boneQuota + p.floatFore);
                _sum19 += _list_ysrs19.Where(p => p.monthly >= (_nowMonth + 2) && p.monthly <= endMonth).Sum(p => p.coreQuota + p.boneQuota + p.floattzys);
                list[19].yj = (_sum19 / (endMonth - startMonth + 1)).ToString();

                _sum19 = _list_ysrs19.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.coreQuota + p.boneQuota + p.floattzys);
                list[19].ys = (_sum19 / (endMonth - startMonth + 1)).ToString();


                double _sum20 = _list_ysrs20.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.coreQuota + p.boneQuota + p.floatActual);
                list[20].sj = (_sum20 / (_nowMonth - startMonth)).ToString();

                _sum20 += _list_ysrs20.Where(p => p.monthly >= _nowMonth && p.monthly <= (_nowMonth + 2)).Sum(p => p.coreQuota + p.boneQuota + p.floatFore);
                _sum20 += _list_ysrs20.Where(p => p.monthly >= (_nowMonth + 2) && p.monthly <= endMonth).Sum(p => p.coreQuota + p.boneQuota + p.floattzys);
                list[20].yj = (_sum20 / (endMonth - startMonth + 1)).ToString();

                _sum20 = _list_ysrs20.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.coreQuota + p.boneQuota + p.floattzys);
                list[20].ys = (_sum20 / (endMonth - startMonth + 1)).ToString();

                #endregion

                #region//序号13
                /*
                  * Cyber-总人数 = 营销中心 + 干部 + OP + OO
                  *  = 17 + 18 + 19 + 20
                  */
                list[11].classify = "Cyber-总人数";
                list[11].yearLine = "1";
                list[11].goal = (int.Parse(list[15].goal) + int.Parse(list[16].goal) + int.Parse(list[17].goal) + int.Parse(list[18].goal)).ToString();
                list[11].month1 = (int.Parse(list[15].month1) + int.Parse(list[16].month1) + int.Parse(list[17].month1) + int.Parse(list[18].month1)).ToString();
                list[11].month2 = (int.Parse(list[15].month2) + int.Parse(list[16].month2) + int.Parse(list[17].month2) + int.Parse(list[18].month2)).ToString();
                list[11].month3 = (int.Parse(list[15].month3) + int.Parse(list[16].month3) + int.Parse(list[17].month3) + int.Parse(list[18].month3)).ToString();
                list[11].month4 = (int.Parse(list[15].month4) + int.Parse(list[16].month4) + int.Parse(list[17].month4) + int.Parse(list[18].month4)).ToString();
                list[11].month5 = (int.Parse(list[15].month5) + int.Parse(list[16].month5) + int.Parse(list[17].month5) + int.Parse(list[18].month5)).ToString();
                list[11].month6 = (int.Parse(list[15].month6) + int.Parse(list[16].month6) + int.Parse(list[17].month6) + int.Parse(list[18].month6)).ToString();
                list[11].month7 = (int.Parse(list[15].month7) + int.Parse(list[16].month7) + int.Parse(list[17].month7) + int.Parse(list[18].month7)).ToString();
                list[11].month8 = (int.Parse(list[15].month8) + int.Parse(list[16].month8) + int.Parse(list[17].month8) + int.Parse(list[18].month8)).ToString();
                list[11].month9 = (int.Parse(list[15].month9) + int.Parse(list[16].month9) + int.Parse(list[17].month9) + int.Parse(list[18].month9)).ToString();
                list[11].month10 = (int.Parse(list[15].month10) + int.Parse(list[16].month10) + int.Parse(list[17].month10) + int.Parse(list[18].month10)).ToString();
                list[11].month11 = (int.Parse(list[15].month11) + int.Parse(list[16].month11) + int.Parse(list[17].month11) + int.Parse(list[18].month11)).ToString();
                list[11].month12 = (int.Parse(list[15].month12) + int.Parse(list[16].month12) + int.Parse(list[17].month12) + int.Parse(list[18].month12)).ToString();
                list[11].sj = (int.Parse(list[15].sj) + int.Parse(list[16].sj) + int.Parse(list[17].sj) + int.Parse(list[18].sj)).ToString();
                list[11].yj = (int.Parse(list[15].yj) + int.Parse(list[16].yj) + int.Parse(list[17].yj) + int.Parse(list[18].yj)).ToString();
                list[11].ys = (int.Parse(list[15].ys) + int.Parse(list[16].ys) + int.Parse(list[17].ys) + int.Parse(list[18].ys)).ToString();

                #endregion

                #region//岗位工资定额 数据 
                //岗位工资定额
                var list_bhr = new List<Hr_Bhr_fact>();

                var dt_bhr = dal.GetHr_Bhr_factByMonth2(yearly, 0, midStr);
                if (dt_bhr.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_bhr.Rows.Count; i++)
                    {
                        var obj4 = new Hr_Bhr_fact();

                        obj4.ruleDeptCode = dt_htrs.Rows[i]["ruleDeptCode"].ToString();
                        obj4.rulePostCode = dt_htrs.Rows[i]["rulePostCode"].ToString();
                        obj4.wage = double.Parse(dt_htrs.Rows[i]["wage"].ToString());
                        obj4.postLevel = dt_htrs.Rows[i]["postLevel"].ToString();
                        obj4.postType = dt_htrs.Rows[i]["postType"].ToString();
                        obj4.monthly = int.Parse(dt_htrs.Rows[i]["monthly"].ToString());

                        list_bhr.Add(obj4);
                    }
                }
                #endregion

                #region//序号24—25
                //序号24
                /* 
                  * 市场人数表：取公司的 调整预算核心人数之和 a
                  * 非市场人数表：取公司 核心定额人数之和 b
                  * 年度目标：取平均值（四舍五入取整数）
                  * 月份： 岗位工资定额表：取公司当月岗位归属为核心的人数
                  */
                list[21].classify = "其中：核心";
                list[21].yearLine = "1";
                list[21].goal = (list_htrs.Sum(p => p.core_tzys) / 12 + list_ysrs.Sum(p => p.coreQuota) / 12).ToString();
                list[21].month1 = list_bhr.Where(p => p.monthly == 1 && p.postType == "核心").Count().ToString();
                list[21].month2 = list_bhr.Where(p => p.monthly == 2 && p.postType == "核心").Count().ToString();
                list[21].month3 = list_bhr.Where(p => p.monthly == 3 && p.postType == "核心").Count().ToString();
                list[21].month4 = list_bhr.Where(p => p.monthly == 4 && p.postType == "核心").Count().ToString();
                list[21].month5 = list_bhr.Where(p => p.monthly == 5 && p.postType == "核心").Count().ToString();
                list[21].month6 = list_bhr.Where(p => p.monthly == 6 && p.postType == "核心").Count().ToString();
                list[21].month7 = list_bhr.Where(p => p.monthly == 7 && p.postType == "核心").Count().ToString();
                list[21].month8 = list_bhr.Where(p => p.monthly == 8 && p.postType == "核心").Count().ToString();
                list[21].month9 = list_bhr.Where(p => p.monthly == 9 && p.postType == "核心").Count().ToString();
                list[21].month10 = list_bhr.Where(p => p.monthly == 10 && p.postType == "核心").Count().ToString();
                list[21].month11 = list_bhr.Where(p => p.monthly == 11 && p.postType == "核心").Count().ToString();
                list[21].month12 = list_bhr.Where(p => p.monthly == 12 && p.postType == "核心").Count().ToString();

                //序号25
                /* 
                  * 市场人数表：取公司的 调整预算骨干人数之和 a
                  * 非市场人数表：取公司 骨干定额人数之和 b
                  * 年度目标：取平均值（四舍五入取整数）
                  * 月份： 岗位工资定额表：取公司当月岗位归属为骨干的人数
                  */
                list[22].classify = "骨干";
                list[22].yearLine = "1";
                list[22].goal = (list_htrs.Sum(p => p.bone_tzys) / 12 + list_ysrs.Sum(p => p.boneQuota) / 12).ToString();
                list[22].month1 = list_bhr.Where(p => p.monthly == 1 && p.postType == "骨干").Count().ToString();
                list[22].month2 = list_bhr.Where(p => p.monthly == 2 && p.postType == "骨干").Count().ToString();
                list[22].month3 = list_bhr.Where(p => p.monthly == 3 && p.postType == "骨干").Count().ToString();
                list[22].month4 = list_bhr.Where(p => p.monthly == 4 && p.postType == "骨干").Count().ToString();
                list[22].month5 = list_bhr.Where(p => p.monthly == 5 && p.postType == "骨干").Count().ToString();
                list[22].month6 = list_bhr.Where(p => p.monthly == 6 && p.postType == "骨干").Count().ToString();
                list[22].month7 = list_bhr.Where(p => p.monthly == 7 && p.postType == "骨干").Count().ToString();
                list[22].month8 = list_bhr.Where(p => p.monthly == 8 && p.postType == "骨干").Count().ToString();
                list[22].month9 = list_bhr.Where(p => p.monthly == 9 && p.postType == "骨干").Count().ToString();
                list[22].month10 = list_bhr.Where(p => p.monthly == 10 && p.postType == "骨干").Count().ToString();
                list[22].month11 = list_bhr.Where(p => p.monthly == 11 && p.postType == "骨干").Count().ToString();
                list[22].month12 = list_bhr.Where(p => p.monthly == 12 && p.postType == "骨干").Count().ToString();

                /*
                  * 年累实际：显示的月份中有 实际值月份 的人数（月度平均值）【1月到当前月-1】
                  * 年累预计：显示的月份的人数（月度平均值）
                  * 年累预算：显示的月份的‘调整预算值’人数（月度平均值）
                  */
                double _sum21 = list_bhr.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth && p.postType == "核心").Count();
                list[21].sj = (_sum21 / (_nowMonth - startMonth)).ToString();
                _sum21 = list_bhr.Where(p => p.monthly >= startMonth && p.monthly <= endMonth && p.postType == "核心").Count();
                list[21].yj = (_sum21 / (_nowMonth - startMonth)).ToString();

                _core_tzys = list_htrs.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.core_tzys);
                _coreQuota = list_ysrs.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.coreQuota);                
                list[21].ys = ((_core_tzys + _coreQuota) / (endMonth - startMonth + 1)).ToString();

                double _sum22 = list_bhr.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth && p.postType == "骨干").Count();
                list[22].sj = (_sum22 / (_nowMonth - startMonth)).ToString();
                _sum22 = list_bhr.Where(p => p.monthly >= startMonth && p.monthly <= endMonth && p.postType == "骨干").Count();
                list[22].yj = (_sum22 / (_nowMonth - startMonth)).ToString();

                _bone_tzys = list_htrs.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.bone_tzys);
                _boneQuota = list_ysrs.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.boneQuota);                
                list[22].ys = ((_bone_tzys + _boneQuota) / (endMonth - startMonth + 1)).ToString();

                #endregion
                
                #region//序号26
                /* 
                  * 年度目标：   调整预算表：取公司的项目组数（月度平均值）
                  * 月份：       实际表：取公司当月的项目组数
                  * 年度目标：取平均值（四舍五入取整数）
                  * 月份： 岗位工资定额表：取公司当月岗位归属为核心的人数
                  */
                list[24].classify = "销售团队数";
                list[24].yearLine = "1";
                list[24].goal = (list_tzys.Sum(p => p.proTeams) / 12).ToString(); //平均值
                list[24].month1 = list_fact1[0].proTeams.ToString();
                list[24].month2 = list_fact2[0].proTeams.ToString();
                list[24].month3 = list_fact3[0].proTeams.ToString();
                list[24].month4 = list_fact4[0].proTeams.ToString();
                list[24].month5 = list_fact5[0].proTeams.ToString();
                list[24].month6 = list_fact6[0].proTeams.ToString();
                list[24].month7 = list_fact7[0].proTeams.ToString();
                list[24].month8 = list_fact8[0].proTeams.ToString();
                list[24].month9 = list_fact9[0].proTeams.ToString();
                list[24].month10 = list_fact10[0].proTeams.ToString();
                list[24].month11 = list_fact11[0].proTeams.ToString();
                list[24].month12 = list_fact12[0].proTeams.ToString();

                double _sum24 = list_fact.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p=>p.proTeams);
                list[24].sj = (_sum24 / (_nowMonth - startMonth)).ToString();
                _sum24 = list_fact.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.proTeams);
                list[24].yj = (_sum24 / (endMonth - startMonth + 1)).ToString();
                _sum24 = list_tzys.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.proTeams);
                list[24].yj = (_sum24 / (endMonth - startMonth + 1)).ToString();
                #endregion

                #region//序号27
                /* 
                  * 年度目标：   市场人数表：取公司的调整预算（核心 + 骨干）人数之和（月度平均值）
                  * 月份：       岗位工资定额表：取公司当月市场部的人数之和
                  */

                //市场部
                var list_bhr1 = list_bhr.Where(p => p.easDeptName == "营销中心" || p.easDeptName == "综合部" || (p.easDeptName.StartsWith("项目") && p.easDeptName.EndsWith("部"))).ToList();

                list[25].classify = "其中：营销中心";
                list[25].yearLine = "1";
                list[25].goal = (list_htrs.Sum(p => p.core_tzys + p.bone_tzys) / 12).ToString();
                list[25].month1 = list_bhr1.Where(p => p.monthly == 1).Count().ToString();
                list[25].month2 = list_bhr1.Where(p => p.monthly == 2).Count().ToString();
                list[25].month3 = list_bhr1.Where(p => p.monthly == 3).Count().ToString();
                list[25].month4 = list_bhr1.Where(p => p.monthly == 4).Count().ToString();
                list[25].month5 = list_bhr1.Where(p => p.monthly == 5).Count().ToString();
                list[25].month6 = list_bhr1.Where(p => p.monthly == 6).Count().ToString();
                list[25].month7 = list_bhr1.Where(p => p.monthly == 7).Count().ToString();
                list[25].month8 = list_bhr1.Where(p => p.monthly == 8).Count().ToString();
                list[25].month9 = list_bhr1.Where(p => p.monthly == 9).Count().ToString();
                list[25].month10 = list_bhr1.Where(p => p.monthly == 10).Count().ToString();
                list[25].month11 = list_bhr1.Where(p => p.monthly == 11).Count().ToString();
                list[25].month12 = list_bhr1.Where(p => p.monthly == 12).Count().ToString();

                double _sum25 = list_bhr1.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Count();
                list[25].sj = (_sum25 / (_nowMonth - startMonth)).ToString();
                _sum25 = list_bhr1.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Count();
                list[25].yj = (_sum25 / (endMonth - startMonth + 1)).ToString();
                _sum25 = list_htrs.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.core_tzys + p.bone_tzys);
                list[25].yj = (_sum25 / (endMonth - startMonth + 1)).ToString();
                #endregion

                #region//序号28—32
                //序号28
                /* 
                  * 月份：岗位工资定额表：取公司当月非市场部职级为A-F的人数之和
                  * 年度目标：非市场人数表：取公司岗位对应职级为A-F（核心定额 + 骨干定额+浮动调整预算）人数之和（月度平均值）
                  */
                list[26].classify = "干部";
                list[26].yearLine = "1";
                list[26].goal = (_list_ysrs16.Sum(p => p.coreQuota + p.boneQuota + p.floatQuota) / 12).ToString();
                list[26].month1 = (_list_ysrs16.Where(p => p.monthly == 1).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[26].month2 = (_list_ysrs16.Where(p => p.monthly == 2).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[26].month3 = (_list_ysrs16.Where(p => p.monthly == 3).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[26].month4 = (_list_ysrs16.Where(p => p.monthly == 4).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[26].month5 = (_list_ysrs16.Where(p => p.monthly == 5).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[26].month6 = (_list_ysrs16.Where(p => p.monthly == 6).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[26].month7 = (_list_ysrs16.Where(p => p.monthly == 7).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[26].month8 = (_list_ysrs16.Where(p => p.monthly == 8).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[26].month9 = (_list_ysrs16.Where(p => p.monthly == 9).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[26].month10 = (_list_ysrs16.Where(p => p.monthly == 10).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[26].month11 = (_list_ysrs16.Where(p => p.monthly == 11).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[26].month12 = (_list_ysrs16.Where(p => p.monthly == 12).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();

                double _sum26 = _list_ysrs16.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.coreActual + p.boneActual + p.floatActual);
                list[26].sj = (_sum26 / (_nowMonth - startMonth)).ToString();
                _sum26 = _list_ysrs16.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.coreActual + p.boneActual + p.floatActual);
                list[26].yj = (_sum26 / (endMonth - startMonth + 1)).ToString();
                _sum26 = _list_ysrs16.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.coreQuota + p.boneQuota + p.floattzys);
                list[26].yj = (_sum26 / (endMonth - startMonth + 1)).ToString();

                //序号29
                /* 
                  * 月份：       岗位工资定额表：取公司当月职级为OP的人数之和
                  * 年度目标：   非市场人数表：取公司岗位对应职级为OP的定额（核心 + 骨干+浮动）人数之和（月度平均值）
                  */
                list[27].classify = "OP";
                list[27].yearLine = "1";
                list[27].goal = (_list_ysrs17.Sum(p => p.coreQuota + p.boneQuota + p.floatQuota) / 12).ToString();
                list[27].month1 = (_list_ysrs17.Where(p => p.monthly == 1).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[27].month2 = (_list_ysrs17.Where(p => p.monthly == 2).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[27].month3 = (_list_ysrs17.Where(p => p.monthly == 3).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[27].month4 = (_list_ysrs17.Where(p => p.monthly == 4).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[27].month5 = (_list_ysrs17.Where(p => p.monthly == 5).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[27].month6 = (_list_ysrs17.Where(p => p.monthly == 6).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[27].month7 = (_list_ysrs17.Where(p => p.monthly == 7).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[27].month8 = (_list_ysrs17.Where(p => p.monthly == 8).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[27].month9 = (_list_ysrs17.Where(p => p.monthly == 9).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[27].month10 = (_list_ysrs17.Where(p => p.monthly == 10).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[27].month11 = (_list_ysrs17.Where(p => p.monthly == 11).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[27].month12 = (_list_ysrs17.Where(p => p.monthly == 12).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();

                double _sum27 = _list_ysrs17.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.coreActual + p.boneActual + p.floatActual);
                list[27].sj = (_sum27 / (_nowMonth - startMonth)).ToString();
                _sum27 = _list_ysrs17.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.coreActual + p.boneActual + p.floatActual);
                list[27].yj = (_sum27 / (endMonth - startMonth + 1)).ToString();
                _sum27 = _list_ysrs17.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.coreQuota + p.boneQuota + p.floattzys);
                list[27].yj = (_sum27 / (endMonth - startMonth + 1)).ToString();

                //序号30
                /* 
                  * 月份：       岗位工资定额表：取公司当月职级为OO的人数之和
                  * 年度目标：   非市场人数表：取公司岗位对应职级为OO的定额（核心 + 骨干+浮动）人数之和（月度平均值）
                  */
                list[28].classify = "OO";
                list[28].yearLine = "1";
                list[28].goal = (_list_ysrs18.Sum(p => p.coreQuota + p.boneQuota + p.floatQuota) / 12).ToString();
                list[28].month1 = (_list_ysrs18.Where(p => p.monthly == 1).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[28].month2 = (_list_ysrs18.Where(p => p.monthly == 2).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[28].month3 = (_list_ysrs18.Where(p => p.monthly == 3).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[28].month4 = (_list_ysrs18.Where(p => p.monthly == 4).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[28].month5 = (_list_ysrs18.Where(p => p.monthly == 5).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[28].month6 = (_list_ysrs18.Where(p => p.monthly == 6).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[28].month7 = (_list_ysrs18.Where(p => p.monthly == 7).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[28].month8 = (_list_ysrs18.Where(p => p.monthly == 8).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[28].month9 = (_list_ysrs18.Where(p => p.monthly == 9).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[28].month10 = (_list_ysrs18.Where(p => p.monthly == 10).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[28].month11 = (_list_ysrs18.Where(p => p.monthly == 11).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[28].month12 = (_list_ysrs18.Where(p => p.monthly == 12).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();

                double _sum28 = _list_ysrs18.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.coreActual + p.boneActual + p.floatActual);
                list[28].sj = (_sum28 / (_nowMonth - startMonth)).ToString();
                _sum28 = _list_ysrs18.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.coreActual + p.boneActual + p.floatActual);
                list[28].yj = (_sum28 / (endMonth - startMonth + 1)).ToString();
                _sum28 = _list_ysrs18.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.coreQuota + p.boneQuota + p.floattzys);
                list[28].yj = (_sum28 / (endMonth - startMonth + 1)).ToString();

                //序号31
                /* 
                  * 月份：       岗位工资定额表：取公司当月BMI供应链部门中职级为OO的人数之和
                  * 年度目标：   非市场人数表：取公司BMI供应链部门下岗位对应职级为OO的定额（核心 + 骨干+浮动）人数之和（月度平均值）
                  */
                list[29].classify = "BMI";
                list[29].yearLine = "1";
                list[29].goal = (_list_ysrs19.Sum(p => p.coreQuota + p.boneQuota + p.floatQuota) / 12).ToString();
                list[29].month1 = (_list_ysrs19.Where(p => p.monthly == 1).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[29].month2 = (_list_ysrs19.Where(p => p.monthly == 2).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[29].month3 = (_list_ysrs19.Where(p => p.monthly == 3).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[29].month4 = (_list_ysrs19.Where(p => p.monthly == 4).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[29].month5 = (_list_ysrs19.Where(p => p.monthly == 5).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[29].month6 = (_list_ysrs19.Where(p => p.monthly == 6).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[29].month7 = (_list_ysrs19.Where(p => p.monthly == 7).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[29].month8 = (_list_ysrs19.Where(p => p.monthly == 8).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[29].month9 = (_list_ysrs19.Where(p => p.monthly == 9).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[29].month10 = (_list_ysrs19.Where(p => p.monthly == 10).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[29].month11 = (_list_ysrs19.Where(p => p.monthly == 11).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[29].month12 = (_list_ysrs19.Where(p => p.monthly == 12).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();

                double _sum29 = _list_ysrs19.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.coreActual + p.boneActual + p.floatActual);
                list[29].sj = (_sum29 / (_nowMonth - startMonth)).ToString();
                _sum29 = _list_ysrs19.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.coreActual + p.boneActual + p.floatActual);
                list[29].yj = (_sum29 / (endMonth - startMonth + 1)).ToString();
                _sum29 = _list_ysrs19.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.coreQuota + p.boneQuota + p.floattzys);
                list[29].yj = (_sum29 / (endMonth - startMonth + 1)).ToString();

                //序号32
                /* 
                  * 月份：       岗位工资定额表：取公司当月BPL生产部门中职级为OO的人数之和
                  * 年度目标：   非市场人数表：取公司BPL生产部门下岗位对应职级为OO的定额（核心 + 骨干+浮动）人数之和（月度平均值）
                  */
                list[30].classify = "BPL";
                list[30].yearLine = "1";
                list[30].goal = (_list_ysrs20.Sum(p => p.coreQuota + p.boneQuota + p.floatQuota) / 12).ToString();
                list[30].month1 = (_list_ysrs20.Where(p => p.monthly == 1).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[30].month2 = (_list_ysrs20.Where(p => p.monthly == 2).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[30].month3 = (_list_ysrs20.Where(p => p.monthly == 3).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[30].month4 = (_list_ysrs20.Where(p => p.monthly == 4).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[30].month5 = (_list_ysrs20.Where(p => p.monthly == 5).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[30].month6 = (_list_ysrs20.Where(p => p.monthly == 6).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[30].month7 = (_list_ysrs20.Where(p => p.monthly == 7).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[30].month8 = (_list_ysrs20.Where(p => p.monthly == 8).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[30].month9 = (_list_ysrs20.Where(p => p.monthly == 9).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[30].month10 = (_list_ysrs20.Where(p => p.monthly == 10).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[30].month11 = (_list_ysrs20.Where(p => p.monthly == 11).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();
                list[30].month12 = (_list_ysrs20.Where(p => p.monthly == 12).Sum(p => p.coreActual + p.boneActual + p.floatActual)).ToString();

                double _sum30 = _list_ysrs20.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.coreActual + p.boneActual + p.floatActual);
                list[30].sj = (_sum30 / (_nowMonth - startMonth)).ToString();
                _sum30 = _list_ysrs20.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.coreActual + p.boneActual + p.floatActual);
                list[30].yj = (_sum30 / (endMonth - startMonth + 1)).ToString();
                _sum30 = _list_ysrs20.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.coreQuota + p.boneQuota + p.floattzys);
                list[30].yj = (_sum30 / (endMonth - startMonth + 1)).ToString();
                #endregion

                #region//序号23
                /*
                  * Physical-总人数 = 营销中心 + 干部 + OP + OO
                  *  = 27 + 28 + 29 + 30
                  */
                list[21].classify = "Physical-总人数";
                list[21].yearLine = "1";
                list[21].goal = (int.Parse(list[25].goal) + int.Parse(list[26].goal) + int.Parse(list[27].goal) + int.Parse(list[28].goal)).ToString();
                list[21].month1 = (int.Parse(list[25].month1) + int.Parse(list[26].month1) + int.Parse(list[27].month1) + int.Parse(list[28].month1)).ToString();
                list[21].month2 = (int.Parse(list[25].month2) + int.Parse(list[26].month2) + int.Parse(list[27].month2) + int.Parse(list[28].month2)).ToString();
                list[21].month3 = (int.Parse(list[25].month3) + int.Parse(list[26].month3) + int.Parse(list[27].month3) + int.Parse(list[28].month3)).ToString();
                list[21].month4 = (int.Parse(list[25].month4) + int.Parse(list[26].month4) + int.Parse(list[27].month4) + int.Parse(list[28].month4)).ToString();
                list[21].month5 = (int.Parse(list[25].month5) + int.Parse(list[26].month5) + int.Parse(list[27].month5) + int.Parse(list[28].month5)).ToString();
                list[21].month6 = (int.Parse(list[25].month6) + int.Parse(list[26].month6) + int.Parse(list[27].month6) + int.Parse(list[28].month6)).ToString();
                list[21].month7 = (int.Parse(list[25].month7) + int.Parse(list[26].month7) + int.Parse(list[27].month7) + int.Parse(list[28].month7)).ToString();
                list[21].month8 = (int.Parse(list[25].month8) + int.Parse(list[26].month8) + int.Parse(list[27].month8) + int.Parse(list[28].month8)).ToString();
                list[21].month9 = (int.Parse(list[25].month9) + int.Parse(list[26].month9) + int.Parse(list[27].month9) + int.Parse(list[28].month9)).ToString();
                list[21].month10 = (int.Parse(list[25].month10) + int.Parse(list[26].month10) + int.Parse(list[27].month10) + int.Parse(list[28].month10)).ToString();
                list[21].month11 = (int.Parse(list[25].month11) + int.Parse(list[26].month11) + int.Parse(list[27].month11) + int.Parse(list[28].month11)).ToString();
                list[21].month12 = (int.Parse(list[25].month12) + int.Parse(list[26].month12) + int.Parse(list[27].month12) + int.Parse(list[28].month12)).ToString();
                list[21].sj = (int.Parse(list[25].sj) + int.Parse(list[26].sj) + int.Parse(list[27].sj) + int.Parse(list[28].sj)).ToString();
                list[21].yj = (int.Parse(list[25].yj) + int.Parse(list[26].yj) + int.Parse(list[27].yj) + int.Parse(list[28].yj)).ToString();
                list[21].ys = (int.Parse(list[25].ys) + int.Parse(list[26].ys) + int.Parse(list[27].ys) + int.Parse(list[28].ys)).ToString();

                #endregion

                #region//序号33—42
                //序号33
                //13行 减去 23行
                list[31].classify = "人员总需求";
                list[31].yearLine = (int.Parse(list[11].yearLine) - int.Parse(list[21].yearLine)).ToString();
                list[31].goal = (int.Parse(list[11].goal) - int.Parse(list[21].goal)).ToString();
                list[31].month1 = (int.Parse(list[11].month1) - int.Parse(list[21].month1)).ToString();
                list[31].month2 = (int.Parse(list[11].month2) - int.Parse(list[21].month2)).ToString();
                list[31].month3 = (int.Parse(list[11].month3) - int.Parse(list[21].month3)).ToString();
                list[31].month4 = (int.Parse(list[11].month4) - int.Parse(list[21].month4)).ToString();
                list[31].month5 = (int.Parse(list[11].month5) - int.Parse(list[21].month5)).ToString();
                list[31].month6 = (int.Parse(list[11].month6) - int.Parse(list[21].month6)).ToString();
                list[31].month7 = (int.Parse(list[11].month7) - int.Parse(list[21].month7)).ToString();
                list[31].month8 = (int.Parse(list[11].month8) - int.Parse(list[21].month8)).ToString();
                list[31].month9 = (int.Parse(list[11].month9) - int.Parse(list[21].month9)).ToString();
                list[31].month10 = (int.Parse(list[11].month10) - int.Parse(list[21].month10)).ToString();
                list[31].month11 = (int.Parse(list[11].month11) - int.Parse(list[21].month11)).ToString();
                list[31].month12 = (int.Parse(list[11].month12) - int.Parse(list[21].month12)).ToString();
                list[31].sj = (int.Parse(list[11].sj) - int.Parse(list[21].sj)).ToString();
                list[31].yj = (int.Parse(list[11].yj) - int.Parse(list[21].yj)).ToString();
                list[31].ys = (int.Parse(list[11].ys) - int.Parse(list[21].ys)).ToString();

                //序号34
                //14行 减去 24行
                list[32].classify = "其中：核心";
                list[32].yearLine = (int.Parse(list[12].yearLine) - int.Parse(list[22].yearLine)).ToString();
                list[32].goal = (int.Parse(list[12].goal) - int.Parse(list[22].goal)).ToString();
                list[32].month1 = (int.Parse(list[12].month1) - int.Parse(list[22].month1)).ToString();
                list[32].month2 = (int.Parse(list[12].month2) - int.Parse(list[22].month2)).ToString();
                list[32].month3 = (int.Parse(list[12].month3) - int.Parse(list[22].month3)).ToString();
                list[32].month4 = (int.Parse(list[12].month4) - int.Parse(list[22].month4)).ToString();
                list[32].month5 = (int.Parse(list[12].month5) - int.Parse(list[22].month5)).ToString();
                list[32].month6 = (int.Parse(list[12].month6) - int.Parse(list[22].month6)).ToString();
                list[32].month7 = (int.Parse(list[12].month7) - int.Parse(list[22].month7)).ToString();
                list[32].month8 = (int.Parse(list[12].month8) - int.Parse(list[22].month8)).ToString();
                list[32].month9 = (int.Parse(list[12].month9) - int.Parse(list[22].month9)).ToString();
                list[32].month10 = (int.Parse(list[12].month10) - int.Parse(list[22].month10)).ToString();
                list[32].month11 = (int.Parse(list[12].month11) - int.Parse(list[22].month11)).ToString();
                list[32].month12 = (int.Parse(list[12].month12) - int.Parse(list[22].month12)).ToString();
                list[32].sj = (int.Parse(list[12].sj) - int.Parse(list[22].sj)).ToString();
                list[32].yj = (int.Parse(list[12].yj) - int.Parse(list[22].yj)).ToString();
                list[32].ys = (int.Parse(list[12].ys) - int.Parse(list[22].ys)).ToString();

                //序号35
                //15行 减去 25行
                list[33].classify = "骨干";
                list[33].yearLine = (int.Parse(list[13].yearLine) - int.Parse(list[23].yearLine)).ToString();
                list[33].goal = (int.Parse(list[13].goal) - int.Parse(list[23].goal)).ToString();
                list[33].month1 = (int.Parse(list[13].month1) - int.Parse(list[23].month1)).ToString();
                list[33].month2 = (int.Parse(list[13].month2) - int.Parse(list[23].month2)).ToString();
                list[33].month3 = (int.Parse(list[13].month3) - int.Parse(list[23].month3)).ToString();
                list[33].month4 = (int.Parse(list[13].month4) - int.Parse(list[23].month4)).ToString();
                list[33].month5 = (int.Parse(list[13].month5) - int.Parse(list[23].month5)).ToString();
                list[33].month6 = (int.Parse(list[13].month6) - int.Parse(list[23].month6)).ToString();
                list[33].month7 = (int.Parse(list[13].month7) - int.Parse(list[23].month7)).ToString();
                list[33].month8 = (int.Parse(list[13].month8) - int.Parse(list[23].month8)).ToString();
                list[33].month9 = (int.Parse(list[13].month9) - int.Parse(list[23].month9)).ToString();
                list[33].month10 = (int.Parse(list[13].month10) - int.Parse(list[23].month10)).ToString();
                list[33].month11 = (int.Parse(list[13].month11) - int.Parse(list[23].month11)).ToString();
                list[33].month12 = (int.Parse(list[13].month12) - int.Parse(list[23].month12)).ToString();
                list[33].sj = (int.Parse(list[13].sj) - int.Parse(list[23].sj)).ToString();
                list[33].yj = (int.Parse(list[13].yj) - int.Parse(list[23].yj)).ToString();
                list[33].ys = (int.Parse(list[13].ys) - int.Parse(list[23].ys)).ToString();

                //序号36
                //16行 减去 26行
                list[34].classify = "销售团队数";
                list[34].yearLine = (int.Parse(list[14].yearLine) - int.Parse(list[24].yearLine)).ToString();
                list[34].goal = (int.Parse(list[14].goal) - int.Parse(list[24].goal)).ToString();
                list[34].month1 = (int.Parse(list[14].month1) - int.Parse(list[24].month1)).ToString();
                list[34].month2 = (int.Parse(list[14].month2) - int.Parse(list[24].month2)).ToString();
                list[34].month3 = (int.Parse(list[14].month3) - int.Parse(list[24].month3)).ToString();
                list[34].month4 = (int.Parse(list[14].month4) - int.Parse(list[24].month4)).ToString();
                list[34].month5 = (int.Parse(list[14].month5) - int.Parse(list[24].month5)).ToString();
                list[34].month6 = (int.Parse(list[14].month6) - int.Parse(list[24].month6)).ToString();
                list[34].month7 = (int.Parse(list[14].month7) - int.Parse(list[24].month7)).ToString();
                list[34].month8 = (int.Parse(list[14].month8) - int.Parse(list[24].month8)).ToString();
                list[34].month9 = (int.Parse(list[14].month9) - int.Parse(list[24].month9)).ToString();
                list[34].month10 = (int.Parse(list[14].month10) - int.Parse(list[24].month10)).ToString();
                list[34].month11 = (int.Parse(list[14].month11) - int.Parse(list[24].month11)).ToString();
                list[34].month12 = (int.Parse(list[14].month12) - int.Parse(list[24].month12)).ToString();
                list[34].sj = (int.Parse(list[14].sj) - int.Parse(list[24].sj)).ToString();
                list[34].yj = (int.Parse(list[14].yj) - int.Parse(list[24].yj)).ToString();
                list[34].ys = (int.Parse(list[14].ys) - int.Parse(list[24].ys)).ToString();

                //序号37
                //17行 减去 27行
                list[35].classify = "其中：营销中心";
                list[35].yearLine = (int.Parse(list[15].yearLine) - int.Parse(list[25].yearLine)).ToString();
                list[35].goal = (int.Parse(list[15].goal) - int.Parse(list[25].goal)).ToString();
                list[35].month1 = (int.Parse(list[15].month1) - int.Parse(list[25].month1)).ToString();
                list[35].month2 = (int.Parse(list[15].month2) - int.Parse(list[25].month2)).ToString();
                list[35].month3 = (int.Parse(list[15].month3) - int.Parse(list[25].month3)).ToString();
                list[35].month4 = (int.Parse(list[15].month4) - int.Parse(list[25].month4)).ToString();
                list[35].month5 = (int.Parse(list[15].month5) - int.Parse(list[25].month5)).ToString();
                list[35].month6 = (int.Parse(list[15].month6) - int.Parse(list[25].month6)).ToString();
                list[35].month7 = (int.Parse(list[15].month7) - int.Parse(list[25].month7)).ToString();
                list[35].month8 = (int.Parse(list[15].month8) - int.Parse(list[25].month8)).ToString();
                list[35].month9 = (int.Parse(list[15].month9) - int.Parse(list[25].month9)).ToString();
                list[35].month10 = (int.Parse(list[15].month10) - int.Parse(list[25].month10)).ToString();
                list[35].month11 = (int.Parse(list[15].month11) - int.Parse(list[25].month11)).ToString();
                list[35].month12 = (int.Parse(list[15].month12) - int.Parse(list[25].month12)).ToString();
                list[35].sj = (int.Parse(list[15].sj) - int.Parse(list[25].sj)).ToString();
                list[35].yj = (int.Parse(list[15].yj) - int.Parse(list[25].yj)).ToString();
                list[35].ys = (int.Parse(list[15].ys) - int.Parse(list[25].ys)).ToString();

                //序号38
                //18行 减去 28行
                list[36].classify = "干部";
                list[36].yearLine = (int.Parse(list[16].yearLine) - int.Parse(list[26].yearLine)).ToString();
                list[36].goal = (int.Parse(list[16].goal) - int.Parse(list[26].goal)).ToString();
                list[36].month1 = (int.Parse(list[16].month1) - int.Parse(list[26].month1)).ToString();
                list[36].month2 = (int.Parse(list[16].month2) - int.Parse(list[26].month2)).ToString();
                list[36].month3 = (int.Parse(list[16].month3) - int.Parse(list[26].month3)).ToString();
                list[36].month4 = (int.Parse(list[16].month4) - int.Parse(list[26].month4)).ToString();
                list[36].month5 = (int.Parse(list[16].month5) - int.Parse(list[26].month5)).ToString();
                list[36].month6 = (int.Parse(list[16].month6) - int.Parse(list[26].month6)).ToString();
                list[36].month7 = (int.Parse(list[16].month7) - int.Parse(list[26].month7)).ToString();
                list[36].month8 = (int.Parse(list[16].month8) - int.Parse(list[26].month8)).ToString();
                list[36].month9 = (int.Parse(list[16].month9) - int.Parse(list[26].month9)).ToString();
                list[36].month10 = (int.Parse(list[16].month10) - int.Parse(list[26].month10)).ToString();
                list[36].month11 = (int.Parse(list[16].month11) - int.Parse(list[26].month11)).ToString();
                list[36].month12 = (int.Parse(list[16].month12) - int.Parse(list[26].month12)).ToString();
                list[36].sj = (int.Parse(list[16].sj) - int.Parse(list[26].sj)).ToString();
                list[36].yj = (int.Parse(list[16].yj) - int.Parse(list[26].yj)).ToString();
                list[36].ys = (int.Parse(list[16].ys) - int.Parse(list[26].ys)).ToString();

                //序号39
                //19行 减去 29行
                list[37].classify = "OP ";
                list[37].yearLine = (int.Parse(list[17].yearLine) - int.Parse(list[27].yearLine)).ToString();
                list[37].goal = (int.Parse(list[17].goal) - int.Parse(list[27].goal)).ToString();
                list[37].month1 = (int.Parse(list[17].month1) - int.Parse(list[27].month1)).ToString();
                list[37].month2 = (int.Parse(list[17].month2) - int.Parse(list[27].month2)).ToString();
                list[37].month3 = (int.Parse(list[17].month3) - int.Parse(list[27].month3)).ToString();
                list[37].month4 = (int.Parse(list[17].month4) - int.Parse(list[27].month4)).ToString();
                list[37].month5 = (int.Parse(list[17].month5) - int.Parse(list[27].month5)).ToString();
                list[37].month6 = (int.Parse(list[17].month6) - int.Parse(list[27].month6)).ToString();
                list[37].month7 = (int.Parse(list[17].month7) - int.Parse(list[27].month7)).ToString();
                list[37].month8 = (int.Parse(list[17].month8) - int.Parse(list[27].month8)).ToString();
                list[37].month9 = (int.Parse(list[17].month9) - int.Parse(list[27].month9)).ToString();
                list[37].month10 = (int.Parse(list[17].month10) - int.Parse(list[27].month10)).ToString();
                list[37].month11 = (int.Parse(list[17].month11) - int.Parse(list[27].month11)).ToString();
                list[37].month12 = (int.Parse(list[17].month12) - int.Parse(list[27].month12)).ToString();
                list[37].sj = (int.Parse(list[17].sj) - int.Parse(list[27].sj)).ToString();
                list[37].yj = (int.Parse(list[17].yj) - int.Parse(list[27].yj)).ToString();
                list[37].ys = (int.Parse(list[17].ys) - int.Parse(list[27].ys)).ToString();

                //序号40
                //20行 减去 30行
                list[38].classify = "OO";
                list[38].yearLine = (int.Parse(list[18].yearLine) - int.Parse(list[28].yearLine)).ToString();
                list[38].goal = (int.Parse(list[18].goal) - int.Parse(list[28].goal)).ToString();
                list[38].month1 = (int.Parse(list[18].month1) - int.Parse(list[28].month1)).ToString();
                list[38].month2 = (int.Parse(list[18].month2) - int.Parse(list[28].month2)).ToString();
                list[38].month3 = (int.Parse(list[18].month3) - int.Parse(list[28].month3)).ToString();
                list[38].month4 = (int.Parse(list[18].month4) - int.Parse(list[28].month4)).ToString();
                list[38].month5 = (int.Parse(list[18].month5) - int.Parse(list[28].month5)).ToString();
                list[38].month6 = (int.Parse(list[18].month6) - int.Parse(list[28].month6)).ToString();
                list[38].month7 = (int.Parse(list[18].month7) - int.Parse(list[28].month7)).ToString();
                list[38].month8 = (int.Parse(list[18].month8) - int.Parse(list[28].month8)).ToString();
                list[38].month9 = (int.Parse(list[18].month9) - int.Parse(list[28].month9)).ToString();
                list[38].month10 = (int.Parse(list[18].month10) - int.Parse(list[28].month10)).ToString();
                list[38].month11 = (int.Parse(list[18].month11) - int.Parse(list[28].month11)).ToString();
                list[38].month12 = (int.Parse(list[18].month12) - int.Parse(list[28].month12)).ToString();
                list[38].sj = (int.Parse(list[18].sj) - int.Parse(list[28].sj)).ToString();
                list[38].yj = (int.Parse(list[18].yj) - int.Parse(list[28].yj)).ToString();
                list[38].ys = (int.Parse(list[18].ys) - int.Parse(list[28].ys)).ToString();

                //序号41
                //21行 减去 31行
                list[39].classify = "BMI";
                list[39].yearLine = (int.Parse(list[19].yearLine) - int.Parse(list[29].yearLine)).ToString();
                list[39].goal = (int.Parse(list[19].goal) - int.Parse(list[29].goal)).ToString();
                list[39].month1 = (int.Parse(list[19].month1) - int.Parse(list[29].month1)).ToString();
                list[39].month2 = (int.Parse(list[19].month2) - int.Parse(list[29].month2)).ToString();
                list[39].month3 = (int.Parse(list[19].month3) - int.Parse(list[29].month3)).ToString();
                list[39].month4 = (int.Parse(list[19].month4) - int.Parse(list[29].month4)).ToString();
                list[39].month5 = (int.Parse(list[19].month5) - int.Parse(list[29].month5)).ToString();
                list[39].month6 = (int.Parse(list[19].month6) - int.Parse(list[29].month6)).ToString();
                list[39].month7 = (int.Parse(list[19].month7) - int.Parse(list[29].month7)).ToString();
                list[39].month8 = (int.Parse(list[19].month8) - int.Parse(list[29].month8)).ToString();
                list[39].month9 = (int.Parse(list[19].month9) - int.Parse(list[29].month9)).ToString();
                list[39].month10 = (int.Parse(list[19].month10) - int.Parse(list[29].month10)).ToString();
                list[39].month11 = (int.Parse(list[19].month11) - int.Parse(list[29].month11)).ToString();
                list[39].month12 = (int.Parse(list[19].month12) - int.Parse(list[29].month12)).ToString();
                list[39].sj = (int.Parse(list[19].sj) - int.Parse(list[29].sj)).ToString();
                list[39].yj = (int.Parse(list[19].yj) - int.Parse(list[29].yj)).ToString();
                list[39].ys = (int.Parse(list[19].ys) - int.Parse(list[29].ys)).ToString();

                //序号42
                //22行 减去 32行
                list[40].classify = "BPL";
                list[40].yearLine = (int.Parse(list[20].yearLine) - int.Parse(list[30].yearLine)).ToString();
                list[40].goal = (int.Parse(list[20].goal) - int.Parse(list[30].goal)).ToString();
                list[40].month1 = (int.Parse(list[20].month1) - int.Parse(list[30].month1)).ToString();
                list[40].month2 = (int.Parse(list[20].month2) - int.Parse(list[30].month2)).ToString();
                list[40].month3 = (int.Parse(list[20].month3) - int.Parse(list[30].month3)).ToString();
                list[40].month4 = (int.Parse(list[20].month4) - int.Parse(list[30].month4)).ToString();
                list[40].month5 = (int.Parse(list[20].month5) - int.Parse(list[30].month5)).ToString();
                list[40].month6 = (int.Parse(list[20].month6) - int.Parse(list[30].month6)).ToString();
                list[40].month7 = (int.Parse(list[20].month7) - int.Parse(list[30].month7)).ToString();
                list[40].month8 = (int.Parse(list[20].month8) - int.Parse(list[30].month8)).ToString();
                list[40].month9 = (int.Parse(list[20].month9) - int.Parse(list[30].month9)).ToString();
                list[40].month10 = (int.Parse(list[20].month10) - int.Parse(list[30].month10)).ToString();
                list[40].month11 = (int.Parse(list[20].month11) - int.Parse(list[30].month11)).ToString();
                list[40].month12 = (int.Parse(list[20].month12) - int.Parse(list[30].month12)).ToString();
                list[40].sj = (int.Parse(list[20].sj) - int.Parse(list[30].sj)).ToString();
                list[40].yj = (int.Parse(list[20].yj) - int.Parse(list[30].yj)).ToString();
                list[40].ys = (int.Parse(list[20].ys) - int.Parse(list[30].ys)).ToString();

                #endregion
                

                #region//人工成本 数据
                //收入-人工成本
                var list_rgsr = new List<Hr_Midrgsr>();

                var dt_rgsr = dal.GetHr_Midrgsr(start, end, mid);
                if (dt_rgsr.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_rgsr.Rows.Count; i++)
                    {
                        var obj5 = new Hr_Midrgsr();

                        obj5.costType = dt_rgsr.Rows[i]["costType"].ToString();
                        obj5.planBudget = double.Parse(dt_rgsr.Rows[i]["planBudget"].ToString());
                        obj5.adjustBudget = double.Parse(dt_rgsr.Rows[i]["adjustBudget"].ToString());
                        obj5.proBudget = double.Parse(dt_rgsr.Rows[i]["proBudget"].ToString());
                        obj5.quotaLabor = double.Parse(dt_rgsr.Rows[i]["quotaLabor"].ToString());
                        obj5.monthly = int.Parse(dt_rgsr.Rows[i]["monthly"].ToString());

                        list_rgsr.Add(obj5);
                    }
                }

                //支出-人工成本
                var list_rgzc = new List<Hr_Midrgzc>();

                var dt_rgzc = dal.GetHr_Midrgzc(start, end, mid);
                if (dt_rgzc.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_rgzc.Rows.Count; i++)
                    {
                        var obj6 = new Hr_Midrgzc();

                        obj6.costType = dt_rgzc.Rows[i]["costType"].ToString();
                        obj6.planBudget = double.Parse(dt_rgzc.Rows[i]["planBudget"].ToString());
                        obj6.adjustBudget = double.Parse(dt_rgzc.Rows[i]["adjustBudget"].ToString());
                        obj6.proBudget = double.Parse(dt_rgzc.Rows[i]["proBudget"].ToString());
                        obj6.monthly = int.Parse(dt_rgzc.Rows[i]["monthly"].ToString());

                        list_rgzc.Add(obj6);
                    }
                }

                //支出-人工成本（实际）
                var list_rgsj = new List<Hr_Midrgsj>();

                var dt_rgsj = dal.GetHr_Midrgsj(start, end, mid);
                if (dt_rgsj.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_rgsj.Rows.Count; i++)
                    {
                        var obj7 = new Hr_Midrgsj();

                        obj7.costType = dt_rgsj.Rows[i]["costType"].ToString();
                        obj7.quotaLabor = int.Parse(dt_rgsj.Rows[i]["quotaLabor"].ToString());
                        obj7.monthly = int.Parse(dt_rgsj.Rows[i]["monthly"].ToString());

                        list_rgsj.Add(obj7);
                    }
                }

                #endregion

                #region//序号44—49，51—56

                #region//收入
                //序号44
                /*
                 * 年度目标：收入-人工成本表：取公司年度费用类别为‘市场人工’调整预算之和
                 * 月份：     当月为x，月份＜x      收入-人工成本表：取公司当月费用类别为‘市场人工’实际值
                 *                  x≤月份≤（x+2）  收入-人工成本表：取公司当月费用类别为‘市场人工’项目进展值
                 *                  月份＞（x+2）     收入-人工成本表：取公司当月费用类别为‘市场人工’调整预算值
                 */
                var _list_rgsr42 = list_rgsr.Where(p => p.costType == "市场人工").ToList();

                list[42].classify = "其中：市场人工";
                list[42].yearLine = "1";
                list[42].goal = _list_rgsr42.Sum(p => p.adjustBudget).ToString();

                /*
                * 年累实际：显示的月份中有实际值月份的和
                * 年累预计：显示的月份的和
                * 年累预算：显示的月份的‘调整预算值’的和
                */
                double _sum42 = _list_rgsr42.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.quotaLabor);
                list[42].sj = (_sum42 / (_nowMonth - startMonth)).ToString();

                _sum42 += _list_rgsr42.Where(p => p.monthly >= _nowMonth && p.monthly <= (_nowMonth + 2)).Sum(p => p.proBudget);
                _sum42 += _list_rgsr42.Where(p => p.monthly >= (_nowMonth + 2) && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[42].yj = (_sum42 / (endMonth - startMonth + 1)).ToString();
                
                _sum42 = _list_rgsr42.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[42].yj = (_sum42 / (endMonth - startMonth + 1)).ToString();

                //序号45
                /*
                 * 年度目标：收入-人工成本表：取公司年度费用类别为‘管理人工’调整预算之和
                 * 月份：     当月为x，月份＜x      收入-人工成本表：取公司当月费用类别为‘管理人工’实际值
                 *                  x≤月份≤（x+2）  收入-人工成本表：取公司当月费用类别为‘管理人工’项目进展值
                 *                  月份＞（x+2）     收入-人工成本表：取公司当月费用类别为‘管理人工’调整预算值
                 */
                var _list_rgsr43 = list_rgsr.Where(p => p.costType == "管理人工").ToList();

                list[43].classify = "管理人工";
                list[43].yearLine = "1";
                list[43].goal = _list_rgsr43.Sum(p => p.adjustBudget).ToString();

                double _sum43 = _list_rgsr43.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.quotaLabor);
                list[43].sj = (_sum43 / (_nowMonth - startMonth)).ToString();

                _sum43 += _list_rgsr43.Where(p => p.monthly >= _nowMonth && p.monthly <= (_nowMonth + 2)).Sum(p => p.proBudget);
                _sum43 += _list_rgsr43.Where(p => p.monthly >= (_nowMonth + 2) && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[43].yj = (_sum43 / (endMonth - startMonth + 1)).ToString();

                _sum43 = _list_rgsr43.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[43].yj = (_sum43 / (endMonth - startMonth + 1)).ToString();

                //序号46
                /*
                 * 年度目标：收入-人工成本表：取公司年度费用类别为‘制造人工’调整预算之和
                 * 月份：     当月为x，月份＜x      收入-人工成本表：取公司当月费用类别为‘制造人工’实际值
                 *                  x≤月份≤（x+2）  收入-人工成本表：取公司当月费用类别为‘制造人工’项目进展值
                 *                  月份＞（x+2）     收入-人工成本表：取公司当月费用类别为‘制造人工’调整预算值
                 */
                var _list_rgsr44 = list_rgsr.Where(p => p.costType == "制造人工").ToList();

                list[44].classify = "制造人工";
                list[44].yearLine = "1";
                list[44].goal = _list_rgsr44.Sum(p => p.adjustBudget).ToString();

                double _sum44 = _list_rgsr44.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.quotaLabor);
                list[44].sj = (_sum44 / (_nowMonth - startMonth)).ToString();

                _sum44 += _list_rgsr44.Where(p => p.monthly >= _nowMonth && p.monthly <= (_nowMonth + 2)).Sum(p => p.proBudget);
                _sum44 += _list_rgsr44.Where(p => p.monthly >= (_nowMonth + 2) && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[44].yj = (_sum44 / (endMonth - startMonth + 1)).ToString();

                _sum44 = _list_rgsr44.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[44].yj = (_sum44 / (endMonth - startMonth + 1)).ToString();

                //序号47
                /*
                 * 年度目标：收入-人工成本表：取公司年度费用类别为‘制造人工’调整预算之和
                 * 月份：     当月为x，月份＜x      收入-人工成本表：取公司当月费用类别为‘直接人工’实际值
                 *                  x≤月份≤（x+2）  收入-人工成本表：取公司当月费用类别为‘直接人工’项目进展值
                 *                  月份＞（x+2）     收入-人工成本表：取公司当月费用类别为‘直接人工’调整预算值
                 */
                var _list_rgsr45 = list_rgsr.Where(p => p.costType == "直接人工").ToList();

                list[45].classify = "直接人工";
                list[45].yearLine = "1";
                list[45].goal = _list_rgsr45.Sum(p => p.adjustBudget).ToString();

                double _sum45 = _list_rgsr45.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.quotaLabor);
                list[45].sj = (_sum45 / (_nowMonth - startMonth)).ToString();

                _sum45 += _list_rgsr45.Where(p => p.monthly >= _nowMonth && p.monthly <= (_nowMonth + 2)).Sum(p => p.proBudget);
                _sum45 += _list_rgsr45.Where(p => p.monthly >= (_nowMonth + 2) && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[45].yj = (_sum45 / (endMonth - startMonth + 1)).ToString();

                _sum45 = _list_rgsr45.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[45].yj = (_sum45 / (endMonth - startMonth + 1)).ToString();

                //序号48
                /*
                 * 年度目标：收入-人工成本表：取公司年度费用类别为‘制造人工’调整预算之和
                 * 月份：     当月为x，月份＜x      收入-人工成本表：取公司当月费用类别为‘直接人工BMI’实际值
                 *                  x≤月份≤（x+2）  收入-人工成本表：取公司当月费用类别为‘直接人工BMI’项目进展值
                 *                  月份＞（x+2）     收入-人工成本表：取公司当月费用类别为‘直接人工BMI’调整预算值
                 */
                var _list_rgsr46 = list_rgsr.Where(p => p.costType == "直接人工BMI").ToList();

                list[46].classify = "直接人工BMI";
                list[46].yearLine = "1";
                list[46].goal = _list_rgsr46.Sum(p => p.adjustBudget).ToString();

                double _sum46 = _list_rgsr46.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.quotaLabor);
                list[46].sj = (_sum46 / (_nowMonth - startMonth)).ToString();

                _sum46 += _list_rgsr46.Where(p => p.monthly >= _nowMonth && p.monthly <= (_nowMonth + 2)).Sum(p => p.proBudget);
                _sum46 += _list_rgsr46.Where(p => p.monthly >= (_nowMonth + 2) && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[46].yj = (_sum46 / (endMonth - startMonth + 1)).ToString();

                _sum46 = _list_rgsr46.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[46].yj = (_sum46 / (endMonth - startMonth + 1)).ToString();

                //序号49
                /*
                 * 年度目标：收入-人工成本表：取公司年度费用类别为‘直接人工BPL’调整预算之和
                 * 月份：     当月为x，月份＜x      收入-人工成本表：取公司当月费用类别为‘直接人工BPL’实际值
                 *                  x≤月份≤（x+2）  收入-人工成本表：取公司当月费用类别为‘直接人工BPL’项目进展值
                 *                  月份＞（x+2）     收入-人工成本表：取公司当月费用类别为‘直接人工BPL’调整预算值
                 */
                var _list_rgsr47 = list_rgsr.Where(p => p.costType == "直接人工BPL").ToList();

                list[47].classify = "直接人工BPL";
                list[47].yearLine = "1";
                list[47].goal = _list_rgsr47.Sum(p => p.adjustBudget).ToString();

                double _sum47 = _list_rgsr47.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.quotaLabor);
                list[47].sj = (_sum47 / (_nowMonth - startMonth)).ToString();

                _sum47 += _list_rgsr47.Where(p => p.monthly >= _nowMonth && p.monthly <= (_nowMonth + 2)).Sum(p => p.proBudget);
                _sum47 += _list_rgsr47.Where(p => p.monthly >= (_nowMonth + 2) && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[47].yj = (_sum47 / (endMonth - startMonth + 1)).ToString();

                _sum47 = _list_rgsr47.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[47].yj = (_sum47 / (endMonth - startMonth + 1)).ToString();

                #endregion

                #region//支出
                //序号51
                /*
                 * 年度目标：支出-人工成本表：取公司年度费用类别为‘市场人工’调整预算之和
                 * 月份：     当月为x，月份＜x      支出-人工成本表：取公司当月费用类别为‘市场人工’实际值
                 *                  x≤月份≤（x+2）  支出-人工成本表：取公司当月费用类别为‘市场人工’项目进展值
                 *                  月份＞（x+2）     支出-人工成本表：取公司当月费用类别为‘市场人工’调整预算值
                 */
                var _list_rgzc49 = list_rgzc.Where(p => p.costType == "市场人工").ToList();
                var _list_rgsj49 = list_rgsj.Where(p => p.costType == "市场人工").ToList();

                list[49].classify = "其中：市场人工";
                list[49].yearLine = "1";
                list[49].goal = _list_rgzc49.Sum(p => p.adjustBudget).ToString();

                double _sum49 = list_rgsj.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.quotaLabor);
                list[49].sj = (_sum49 / (_nowMonth - startMonth)).ToString();

                _sum49 += _list_rgzc49.Where(p => p.monthly >= _nowMonth && p.monthly <= (_nowMonth + 2)).Sum(p => p.proBudget);
                _sum49 += _list_rgzc49.Where(p => p.monthly >= (_nowMonth + 2) && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[49].yj = (_sum49 / (endMonth - startMonth + 1)).ToString();

                _sum49 = _list_rgzc49.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[49].yj = (_sum49 / (endMonth - startMonth + 1)).ToString();

                //序号52
                /*
                 * 年度目标：支出-人工成本表：取公司年度费用类别为‘管理人工’调整预算之和
                 * 月份：     当月为x，月份＜x      支出-人工成本表：取公司当月费用类别为‘管理人工’实际值
                 *                  x≤月份≤（x+2）  支出-人工成本表：取公司当月费用类别为‘管理人工’项目进展值
                 *                  月份＞（x+2）     支出-人工成本表：取公司当月费用类别为‘管理人工’调整预算值
                 */
                var _list_rgzc50 = list_rgzc.Where(p => p.costType == "管理人工").ToList();
                var _list_rgsj50 = list_rgsj.Where(p => p.costType == "管理人工").ToList();

                list[50].classify = "管理人工";
                list[50].yearLine = "1";
                list[50].goal = _list_rgzc50.Sum(p => p.adjustBudget).ToString();

                double _sum50 = list_rgsj.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.quotaLabor);
                list[50].sj = (_sum50 / (_nowMonth - startMonth)).ToString();

                _sum50 += _list_rgzc50.Where(p => p.monthly >= _nowMonth && p.monthly <= (_nowMonth + 2)).Sum(p => p.proBudget);
                _sum50 += _list_rgzc50.Where(p => p.monthly >= (_nowMonth + 2) && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[50].yj = (_sum50 / (endMonth - startMonth + 1)).ToString();

                _sum50 = _list_rgzc50.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[50].yj = (_sum50 / (endMonth - startMonth + 1)).ToString();

                //序号53
                /*
                 * 年度目标：支出-人工成本表：取公司年度费用类别为‘制造人工’调整预算之和
                 * 月份：     当月为x，月份＜x      支出-人工成本表：取公司当月费用类别为‘制造人工’实际值
                 *                  x≤月份≤（x+2）  支出-人工成本表：取公司当月费用类别为‘制造人工’项目进展值
                 *                  月份＞（x+2）     支出-人工成本表：取公司当月费用类别为‘制造人工’调整预算值
                 */
                var _list_rgzc51 = list_rgzc.Where(p => p.costType == "制造人工").ToList();
                var _list_rgsj51 = list_rgsj.Where(p => p.costType == "制造人工").ToList();

                list[51].classify = "制造人工";
                list[51].yearLine = "1";
                list[51].goal = _list_rgzc51.Sum(p => p.adjustBudget).ToString();

                double _sum51 = list_rgsj.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.quotaLabor);
                list[51].sj = (_sum51 / (_nowMonth - startMonth)).ToString();

                _sum51 += _list_rgzc51.Where(p => p.monthly >= _nowMonth && p.monthly <= (_nowMonth + 2)).Sum(p => p.proBudget);
                _sum51 += _list_rgzc51.Where(p => p.monthly >= (_nowMonth + 2) && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[51].yj = (_sum51 / (endMonth - startMonth + 1)).ToString();

                _sum51 = _list_rgzc51.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[51].yj = (_sum51 / (endMonth - startMonth + 1)).ToString();

                //序号54
                /*
                 * 年度目标：支出-人工成本表：取公司年度费用类别为‘制造人工’调整预算之和
                 * 月份：     当月为x，月份＜x      支出-人工成本表：取公司当月费用类别为‘直接人工’实际值
                 *                  x≤月份≤（x+2）  支出-人工成本表：取公司当月费用类别为‘直接人工’项目进展值
                 *                  月份＞（x+2）     支出-人工成本表：取公司当月费用类别为‘直接人工’调整预算值
                 */
                var _list_rgzc52 = list_rgzc.Where(p => p.costType == "直接人工").ToList();
                var _list_rgsj52 = list_rgsj.Where(p => p.costType == "直接人工").ToList();

                list[52].classify = "直接人工";
                list[52].yearLine = "1";
                list[52].goal = _list_rgzc52.Sum(p => p.adjustBudget).ToString();

                double _sum52 = list_rgsj.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.quotaLabor);
                list[52].sj = (_sum52 / (_nowMonth - startMonth)).ToString();

                _sum52 += _list_rgzc52.Where(p => p.monthly >= _nowMonth && p.monthly <= (_nowMonth + 2)).Sum(p => p.proBudget);
                _sum52 += _list_rgzc52.Where(p => p.monthly >= (_nowMonth + 2) && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[52].yj = (_sum52 / (endMonth - startMonth + 1)).ToString();

                _sum52 = _list_rgzc52.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[52].yj = (_sum52 / (endMonth - startMonth + 1)).ToString();

                //序号55
                /*
                 * 年度目标：支出-人工成本表：取公司年度费用类别为‘制造人工’调整预算之和
                 * 月份：     当月为x，月份＜x      支出-人工成本表：取公司当月费用类别为‘直接人工BMI’实际值
                 *                  x≤月份≤（x+2）  支出-人工成本表：取公司当月费用类别为‘直接人工BMI’项目进展值
                 *                  月份＞（x+2）     支出-人工成本表：取公司当月费用类别为‘直接人工BMI’调整预算值
                 */
                var _list_rgzc53 = list_rgzc.Where(p => p.costType == "直接人工BMI").ToList();
                var _list_rgsj53 = list_rgsj.Where(p => p.costType == "直接人工BMI").ToList();

                list[53].classify = "直接人工BMI";
                list[53].yearLine = "1";
                list[53].goal = _list_rgzc53.Sum(p => p.adjustBudget).ToString();

                double _sum53 = list_rgsj.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.quotaLabor);
                list[53].sj = (_sum53 / (_nowMonth - startMonth)).ToString();

                _sum53 += _list_rgzc53.Where(p => p.monthly >= _nowMonth && p.monthly <= (_nowMonth + 2)).Sum(p => p.proBudget);
                _sum53 += _list_rgzc53.Where(p => p.monthly >= (_nowMonth + 2) && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[53].yj = (_sum53 / (endMonth - startMonth + 1)).ToString();

                _sum53 = _list_rgzc53.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[53].yj = (_sum53 / (endMonth - startMonth + 1)).ToString();

                //序号56
                /*
                 * 年度目标：支出-人工成本表：取公司年度费用类别为‘直接人工BPL’调整预算之和
                 * 月份：     当月为x，月份＜x      支出-人工成本表：取公司当月费用类别为‘直接人工BPL’实际值
                 *                  x≤月份≤（x+2）  支出-人工成本表：取公司当月费用类别为‘直接人工BPL’项目进展值
                 *                  月份＞（x+2）     支出-人工成本表：取公司当月费用类别为‘直接人工BPL’调整预算值
                 */
                var _list_rgzc54 = list_rgzc.Where(p => p.costType == "直接人工BPL").ToList();
                var _list_rgsj54 = list_rgsj.Where(p => p.costType == "直接人工BPL").ToList();

                list[54].classify = "直接人工BPL";
                list[54].yearLine = "1";
                list[54].goal = _list_rgzc54.Sum(p => p.adjustBudget).ToString();

                double _sum54 = list_rgsj.Where(p => p.monthly >= startMonth && p.monthly < _nowMonth).Sum(p => p.quotaLabor);
                list[54].sj = (_sum54 / (_nowMonth - startMonth)).ToString();

                _sum54 += _list_rgzc54.Where(p => p.monthly >= _nowMonth && p.monthly <= (_nowMonth + 2)).Sum(p => p.proBudget);
                _sum54 += _list_rgzc54.Where(p => p.monthly >= (_nowMonth + 2) && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[54].yj = (_sum54 / (endMonth - startMonth + 1)).ToString();

                _sum54 = _list_rgzc54.Where(p => p.monthly >= startMonth && p.monthly <= endMonth).Sum(p => p.adjustBudget);
                list[54].yj = (_sum54 / (endMonth - startMonth + 1)).ToString();

                #endregion

                #region//月份
                switch (_nowMonth)
                {
                    case 1:
                        #region
                        list[42].month1 = _list_rgsr42.Where(p => p.monthly == 1).Sum(p => p.proBudget).ToString();
                        list[42].month2 = _list_rgsr42.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[42].month3 = _list_rgsr42.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[42].month4 = _list_rgsr42.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                        list[42].month5 = _list_rgsr42.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[42].month6 = _list_rgsr42.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[42].month7 = _list_rgsr42.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[42].month8 = _list_rgsr42.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[42].month9 = _list_rgsr42.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[42].month10 = _list_rgsr42.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[42].month11 = _list_rgsr42.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[42].month12 = _list_rgsr42.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[43].month1 = _list_rgsr43.Where(p => p.monthly == 1).Sum(p => p.proBudget).ToString();
                        list[43].month2 = _list_rgsr43.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[43].month3 = _list_rgsr43.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[43].month4 = _list_rgsr43.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                        list[43].month5 = _list_rgsr43.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[43].month6 = _list_rgsr43.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[43].month7 = _list_rgsr43.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[43].month8 = _list_rgsr43.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[43].month9 = _list_rgsr43.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[43].month10 = _list_rgsr43.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[43].month11 = _list_rgsr43.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[43].month12 = _list_rgsr43.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[44].month1 = _list_rgsr44.Where(p => p.monthly == 1).Sum(p => p.proBudget).ToString();
                        list[44].month2 = _list_rgsr44.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[44].month3 = _list_rgsr44.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[44].month4 = _list_rgsr44.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                        list[44].month5 = _list_rgsr44.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[44].month6 = _list_rgsr44.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[44].month7 = _list_rgsr44.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[44].month8 = _list_rgsr44.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[44].month9 = _list_rgsr44.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[44].month10 = _list_rgsr44.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[44].month11 = _list_rgsr44.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[44].month12 = _list_rgsr44.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[45].month1 = _list_rgsr45.Where(p => p.monthly == 1).Sum(p => p.proBudget).ToString();
                        list[45].month2 = _list_rgsr45.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[45].month3 = _list_rgsr45.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[45].month4 = _list_rgsr45.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                        list[45].month5 = _list_rgsr45.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[45].month6 = _list_rgsr45.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[45].month7 = _list_rgsr45.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[45].month8 = _list_rgsr45.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[45].month9 = _list_rgsr45.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[45].month10 = _list_rgsr45.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[45].month11 = _list_rgsr45.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[45].month12 = _list_rgsr45.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[46].month1 = _list_rgsr46.Where(p => p.monthly == 1).Sum(p => p.proBudget).ToString();
                        list[46].month2 = _list_rgsr46.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[46].month3 = _list_rgsr46.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[46].month4 = _list_rgsr46.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                        list[46].month5 = _list_rgsr46.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[46].month6 = _list_rgsr46.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[46].month7 = _list_rgsr46.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[46].month8 = _list_rgsr46.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[46].month9 = _list_rgsr46.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[46].month10 = _list_rgsr46.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[46].month11 = _list_rgsr46.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[46].month12 = _list_rgsr46.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[47].month1 = _list_rgsr47.Where(p => p.monthly == 1).Sum(p => p.proBudget).ToString();
                        list[47].month2 = _list_rgsr47.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[47].month3 = _list_rgsr47.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[47].month4 = _list_rgsr47.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                        list[47].month5 = _list_rgsr47.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[47].month6 = _list_rgsr47.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[47].month7 = _list_rgsr47.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[47].month8 = _list_rgsr47.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[47].month9 = _list_rgsr47.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[47].month10 = _list_rgsr47.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[47].month11 = _list_rgsr47.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[47].month12 = _list_rgsr47.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();


                        list[49].month1 = _list_rgzc49.Where(p => p.monthly == 1).Sum(p => p.proBudget).ToString();
                        list[49].month2 = _list_rgzc49.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[49].month3 = _list_rgzc49.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[49].month4 = _list_rgzc49.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                        list[49].month5 = _list_rgzc49.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[49].month6 = _list_rgzc49.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[49].month7 = _list_rgzc49.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[49].month8 = _list_rgzc49.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[49].month9 = _list_rgzc49.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[49].month10 = _list_rgzc49.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[49].month11 = _list_rgzc49.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[49].month12 = _list_rgzc49.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[50].month1 = _list_rgzc50.Where(p => p.monthly == 1).Sum(p => p.proBudget).ToString();
                        list[50].month2 = _list_rgzc50.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[50].month3 = _list_rgzc50.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[50].month4 = _list_rgzc50.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                        list[50].month5 = _list_rgzc50.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[50].month6 = _list_rgzc50.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[50].month7 = _list_rgzc50.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[50].month8 = _list_rgzc50.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[50].month9 = _list_rgzc50.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[50].month10 = _list_rgzc50.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[50].month11 = _list_rgzc50.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[50].month12 = _list_rgzc50.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[51].month1 = _list_rgzc51.Where(p => p.monthly == 1).Sum(p => p.proBudget).ToString();
                        list[51].month2 = _list_rgzc51.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[51].month3 = _list_rgzc51.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[51].month4 = _list_rgzc51.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                        list[51].month5 = _list_rgzc51.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[51].month6 = _list_rgzc51.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[51].month7 = _list_rgzc51.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[51].month8 = _list_rgzc51.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[51].month9 = _list_rgzc51.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[51].month10 = _list_rgzc51.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[51].month11 = _list_rgzc51.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[51].month12 = _list_rgzc51.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[52].month1 = _list_rgzc52.Where(p => p.monthly == 1).Sum(p => p.proBudget).ToString();
                        list[52].month2 = _list_rgzc52.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[52].month3 = _list_rgzc52.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[52].month4 = _list_rgzc52.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                        list[52].month5 = _list_rgzc52.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[52].month6 = _list_rgzc52.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[52].month7 = _list_rgzc52.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[52].month8 = _list_rgzc52.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[52].month9 = _list_rgzc52.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[52].month10 = _list_rgzc52.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[52].month11 = _list_rgzc52.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[52].month12 = _list_rgzc52.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[53].month1 = _list_rgzc53.Where(p => p.monthly == 1).Sum(p => p.proBudget).ToString();
                        list[53].month2 = _list_rgzc53.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[53].month3 = _list_rgzc53.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[53].month4 = _list_rgzc53.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                        list[53].month5 = _list_rgzc53.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[53].month6 = _list_rgzc53.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[53].month7 = _list_rgzc53.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[53].month8 = _list_rgzc53.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[53].month9 = _list_rgzc53.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[53].month10 = _list_rgzc53.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[53].month11 = _list_rgzc53.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[53].month12 = _list_rgzc53.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[54].month1 = _list_rgzc54.Where(p => p.monthly == 1).Sum(p => p.proBudget).ToString();
                        list[54].month2 = _list_rgzc54.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[54].month3 = _list_rgzc54.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[54].month4 = _list_rgzc54.Where(p => p.monthly == 4).Sum(p => p.adjustBudget).ToString();
                        list[54].month5 = _list_rgzc54.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[54].month6 = _list_rgzc54.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[54].month7 = _list_rgzc54.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[54].month8 = _list_rgzc54.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[54].month9 = _list_rgzc54.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[54].month10 = _list_rgzc54.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[54].month11 = _list_rgzc54.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[54].month12 = _list_rgzc54.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();
                        #endregion
                        break;
                    case 2:
                        #region
                        list[42].month1 = _list_rgsr42.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[42].month2 = _list_rgsr42.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[42].month3 = _list_rgsr42.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[42].month4 = _list_rgsr42.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[42].month5 = _list_rgsr42.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[42].month6 = _list_rgsr42.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[42].month7 = _list_rgsr42.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[42].month8 = _list_rgsr42.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[42].month9 = _list_rgsr42.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[42].month10 = _list_rgsr42.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[42].month11 = _list_rgsr42.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[42].month12 = _list_rgsr42.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[43].month1 = _list_rgsr43.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[43].month2 = _list_rgsr43.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[43].month3 = _list_rgsr43.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[43].month4 = _list_rgsr43.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[43].month5 = _list_rgsr43.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[43].month6 = _list_rgsr43.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[43].month7 = _list_rgsr43.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[43].month8 = _list_rgsr43.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[43].month9 = _list_rgsr43.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[43].month10 = _list_rgsr43.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[43].month11 = _list_rgsr43.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[43].month12 = _list_rgsr43.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[44].month1 = _list_rgsr44.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[44].month2 = _list_rgsr44.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[44].month3 = _list_rgsr44.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[44].month4 = _list_rgsr44.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[44].month5 = _list_rgsr44.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[44].month6 = _list_rgsr44.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[44].month7 = _list_rgsr44.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[44].month8 = _list_rgsr44.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[44].month9 = _list_rgsr44.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[44].month10 = _list_rgsr44.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[44].month11 = _list_rgsr44.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[44].month12 = _list_rgsr44.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[45].month1 = _list_rgsr45.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[45].month2 = _list_rgsr45.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[45].month3 = _list_rgsr45.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[45].month4 = _list_rgsr45.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[45].month5 = _list_rgsr45.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[45].month6 = _list_rgsr45.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[45].month7 = _list_rgsr45.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[45].month8 = _list_rgsr45.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[45].month9 = _list_rgsr45.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[45].month10 = _list_rgsr45.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[45].month11 = _list_rgsr45.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[45].month12 = _list_rgsr45.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[46].month1 = _list_rgsr46.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[46].month2 = _list_rgsr46.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[46].month3 = _list_rgsr46.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[46].month4 = _list_rgsr46.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[46].month5 = _list_rgsr46.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[46].month6 = _list_rgsr46.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[46].month7 = _list_rgsr46.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[46].month8 = _list_rgsr46.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[46].month9 = _list_rgsr46.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[46].month10 = _list_rgsr46.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[46].month11 = _list_rgsr46.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[46].month12 = _list_rgsr46.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[47].month1 = _list_rgsr47.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[47].month2 = _list_rgsr47.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[47].month3 = _list_rgsr47.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[47].month4 = _list_rgsr47.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[47].month5 = _list_rgsr47.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[47].month6 = _list_rgsr47.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[47].month7 = _list_rgsr47.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[47].month8 = _list_rgsr47.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[47].month9 = _list_rgsr47.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[47].month10 = _list_rgsr47.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[47].month11 = _list_rgsr47.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[47].month12 = _list_rgsr47.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();


                        list[49].month1 = _list_rgsj49.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[49].month2 = _list_rgzc49.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[49].month3 = _list_rgzc49.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[49].month4 = _list_rgzc49.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[49].month5 = _list_rgzc49.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[49].month6 = _list_rgzc49.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[49].month7 = _list_rgzc49.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[49].month8 = _list_rgzc49.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[49].month9 = _list_rgzc49.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[49].month10 = _list_rgzc49.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[49].month11 = _list_rgzc49.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[49].month12 = _list_rgzc49.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[50].month1 = _list_rgsj50.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[50].month2 = _list_rgzc50.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[50].month3 = _list_rgzc50.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[50].month4 = _list_rgzc50.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[50].month5 = _list_rgzc50.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[50].month6 = _list_rgzc50.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[50].month7 = _list_rgzc50.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[50].month8 = _list_rgzc50.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[50].month9 = _list_rgzc50.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[50].month10 = _list_rgzc50.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[50].month11 = _list_rgzc50.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[50].month12 = _list_rgzc50.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[51].month1 = _list_rgsj51.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[51].month2 = _list_rgzc51.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[51].month3 = _list_rgzc51.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[51].month4 = _list_rgzc51.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[51].month5 = _list_rgzc51.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[51].month6 = _list_rgzc51.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[51].month7 = _list_rgzc51.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[51].month8 = _list_rgzc51.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[51].month9 = _list_rgzc51.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[51].month10 = _list_rgzc51.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[51].month11 = _list_rgzc51.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[51].month12 = _list_rgzc51.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[52].month1 = _list_rgsj52.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[52].month2 = _list_rgzc52.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[52].month3 = _list_rgzc52.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[52].month4 = _list_rgzc52.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[52].month5 = _list_rgzc52.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[52].month6 = _list_rgzc52.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[52].month7 = _list_rgzc52.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[52].month8 = _list_rgzc52.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[52].month9 = _list_rgzc52.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[52].month10 = _list_rgzc52.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[52].month11 = _list_rgzc52.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[52].month12 = _list_rgzc52.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[53].month1 = _list_rgsj53.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[53].month2 = _list_rgzc53.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[53].month3 = _list_rgzc53.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[53].month4 = _list_rgzc53.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[53].month5 = _list_rgzc53.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[53].month6 = _list_rgzc53.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[53].month7 = _list_rgzc53.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[53].month8 = _list_rgzc53.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[53].month9 = _list_rgzc53.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[53].month10 = _list_rgzc53.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[53].month11 = _list_rgzc53.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[53].month12 = _list_rgzc53.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[54].month1 = _list_rgsj54.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[54].month2 = _list_rgzc54.Where(p => p.monthly == 2).Sum(p => p.proBudget).ToString();
                        list[54].month3 = _list_rgzc54.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[54].month4 = _list_rgzc54.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[54].month5 = _list_rgzc54.Where(p => p.monthly == 5).Sum(p => p.adjustBudget).ToString();
                        list[54].month6 = _list_rgzc54.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[54].month7 = _list_rgzc54.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[54].month8 = _list_rgzc54.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[54].month9 = _list_rgzc54.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[54].month10 = _list_rgzc54.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[54].month11 = _list_rgzc54.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[54].month12 = _list_rgzc54.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();
                        #endregion
                        break;
                    case 3:
                        #region
                        list[42].month1 = _list_rgsr42.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[42].month2 = _list_rgsr42.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[42].month3 = _list_rgsr42.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[42].month4 = _list_rgsr42.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[42].month5 = _list_rgsr42.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[42].month6 = _list_rgsr42.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[42].month7 = _list_rgsr42.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[42].month8 = _list_rgsr42.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[42].month9 = _list_rgsr42.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[42].month10 = _list_rgsr42.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[42].month11 = _list_rgsr42.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[42].month12 = _list_rgsr42.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[43].month1 = _list_rgsr43.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[43].month2 = _list_rgsr43.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[43].month3 = _list_rgsr43.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[43].month4 = _list_rgsr43.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[43].month5 = _list_rgsr43.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[43].month6 = _list_rgsr43.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[43].month7 = _list_rgsr43.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[43].month8 = _list_rgsr43.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[43].month9 = _list_rgsr43.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[43].month10 = _list_rgsr43.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[43].month11 = _list_rgsr43.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[43].month12 = _list_rgsr43.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[44].month1 = _list_rgsr44.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[44].month2 = _list_rgsr44.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[44].month3 = _list_rgsr44.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[44].month4 = _list_rgsr44.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[44].month5 = _list_rgsr44.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[44].month6 = _list_rgsr44.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[44].month7 = _list_rgsr44.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[44].month8 = _list_rgsr44.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[44].month9 = _list_rgsr44.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[44].month10 = _list_rgsr44.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[44].month11 = _list_rgsr44.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[44].month12 = _list_rgsr44.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[45].month1 = _list_rgsr45.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[45].month2 = _list_rgsr45.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[45].month3 = _list_rgsr45.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[45].month4 = _list_rgsr45.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[45].month5 = _list_rgsr45.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[45].month6 = _list_rgsr45.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[45].month7 = _list_rgsr45.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[45].month8 = _list_rgsr45.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[45].month9 = _list_rgsr45.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[45].month10 = _list_rgsr45.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[45].month11 = _list_rgsr45.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[45].month12 = _list_rgsr45.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[46].month1 = _list_rgsr46.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[46].month2 = _list_rgsr46.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[46].month3 = _list_rgsr46.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[46].month4 = _list_rgsr46.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[46].month5 = _list_rgsr46.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[46].month6 = _list_rgsr46.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[46].month7 = _list_rgsr46.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[46].month8 = _list_rgsr46.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[46].month9 = _list_rgsr46.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[46].month10 = _list_rgsr46.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[46].month11 = _list_rgsr46.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[46].month12 = _list_rgsr46.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[47].month1 = _list_rgsr47.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[47].month2 = _list_rgsr47.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[47].month3 = _list_rgsr47.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[47].month4 = _list_rgsr47.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[47].month5 = _list_rgsr47.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[47].month6 = _list_rgsr47.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[47].month7 = _list_rgsr47.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[47].month8 = _list_rgsr47.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[47].month9 = _list_rgsr47.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[47].month10 = _list_rgsr47.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[47].month11 = _list_rgsr47.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[47].month12 = _list_rgsr47.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();


                        list[49].month1 = _list_rgsj49.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[49].month2 = _list_rgsj49.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[49].month3 = _list_rgzc49.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[49].month4 = _list_rgzc49.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[49].month5 = _list_rgzc49.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[49].month6 = _list_rgzc49.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[49].month7 = _list_rgzc49.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[49].month8 = _list_rgzc49.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[49].month9 = _list_rgzc49.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[49].month10 = _list_rgzc49.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[49].month11 = _list_rgzc49.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[49].month12 = _list_rgzc49.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[50].month1 = _list_rgsj50.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[50].month2 = _list_rgsj50.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[50].month3 = _list_rgzc50.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[50].month4 = _list_rgzc50.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[50].month5 = _list_rgzc50.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[50].month6 = _list_rgzc50.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[50].month7 = _list_rgzc50.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[50].month8 = _list_rgzc50.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[50].month9 = _list_rgzc50.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[50].month10 = _list_rgzc50.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[50].month11 = _list_rgzc50.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[50].month12 = _list_rgzc50.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[51].month1 = _list_rgsj51.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[51].month2 = _list_rgsj51.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[51].month3 = _list_rgzc51.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[51].month4 = _list_rgzc51.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[51].month5 = _list_rgzc51.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[51].month6 = _list_rgzc51.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[51].month7 = _list_rgzc51.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[51].month8 = _list_rgzc51.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[51].month9 = _list_rgzc51.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[51].month10 = _list_rgzc51.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[51].month11 = _list_rgzc51.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[51].month12 = _list_rgzc51.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[52].month1 = _list_rgsj52.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[52].month2 = _list_rgsj52.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[52].month3 = _list_rgzc52.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[52].month4 = _list_rgzc52.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[52].month5 = _list_rgzc52.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[52].month6 = _list_rgzc52.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[52].month7 = _list_rgzc52.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[52].month8 = _list_rgzc52.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[52].month9 = _list_rgzc52.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[52].month10 = _list_rgzc52.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[52].month11 = _list_rgzc52.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[52].month12 = _list_rgzc52.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[53].month1 = _list_rgsj53.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[53].month2 = _list_rgsj53.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[53].month3 = _list_rgzc53.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[53].month4 = _list_rgzc53.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[53].month5 = _list_rgzc53.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[53].month6 = _list_rgzc53.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[53].month7 = _list_rgzc53.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[53].month8 = _list_rgzc53.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[53].month9 = _list_rgzc53.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[53].month10 = _list_rgzc53.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[53].month11 = _list_rgzc53.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[53].month12 = _list_rgzc53.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[54].month1 = _list_rgsj54.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[54].month2 = _list_rgsj54.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[54].month3 = _list_rgzc54.Where(p => p.monthly == 3).Sum(p => p.proBudget).ToString();
                        list[54].month4 = _list_rgzc54.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[54].month5 = _list_rgzc54.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[54].month6 = _list_rgzc54.Where(p => p.monthly == 6).Sum(p => p.adjustBudget).ToString();
                        list[54].month7 = _list_rgzc54.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[54].month8 = _list_rgzc54.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[54].month9 = _list_rgzc54.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[54].month10 = _list_rgzc54.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[54].month11 = _list_rgzc54.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[54].month12 = _list_rgzc54.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();
                        #endregion
                        break;
                    case 4:
                        #region
                        list[42].month1 = _list_rgsr42.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[42].month2 = _list_rgsr42.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[42].month3 = _list_rgsr42.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[42].month4 = _list_rgsr42.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[42].month5 = _list_rgsr42.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[42].month6 = _list_rgsr42.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[42].month7 = _list_rgsr42.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[42].month8 = _list_rgsr42.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[42].month9 = _list_rgsr42.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[42].month10 = _list_rgsr42.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[42].month11 = _list_rgsr42.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[42].month12 = _list_rgsr42.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[43].month1 = _list_rgsr43.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[43].month2 = _list_rgsr43.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[43].month3 = _list_rgsr43.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[43].month4 = _list_rgsr43.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[43].month5 = _list_rgsr43.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[43].month6 = _list_rgsr43.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[43].month7 = _list_rgsr43.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[43].month8 = _list_rgsr43.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[43].month9 = _list_rgsr43.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[43].month10 = _list_rgsr43.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[43].month11 = _list_rgsr43.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[43].month12 = _list_rgsr43.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[44].month1 = _list_rgsr44.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[44].month2 = _list_rgsr44.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[44].month3 = _list_rgsr44.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[44].month4 = _list_rgsr44.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[44].month5 = _list_rgsr44.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[44].month6 = _list_rgsr44.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[44].month7 = _list_rgsr44.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[44].month8 = _list_rgsr44.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[44].month9 = _list_rgsr44.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[44].month10 = _list_rgsr44.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[44].month11 = _list_rgsr44.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[44].month12 = _list_rgsr44.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[45].month1 = _list_rgsr45.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[45].month2 = _list_rgsr45.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[45].month3 = _list_rgsr45.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[45].month4 = _list_rgsr45.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[45].month5 = _list_rgsr45.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[45].month6 = _list_rgsr45.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[45].month7 = _list_rgsr45.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[45].month8 = _list_rgsr45.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[45].month9 = _list_rgsr45.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[45].month10 = _list_rgsr45.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[45].month11 = _list_rgsr45.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[45].month12 = _list_rgsr45.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[46].month1 = _list_rgsr46.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[46].month2 = _list_rgsr46.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[46].month3 = _list_rgsr46.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[46].month4 = _list_rgsr46.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[46].month5 = _list_rgsr46.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[46].month6 = _list_rgsr46.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[46].month7 = _list_rgsr46.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[46].month8 = _list_rgsr46.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[46].month9 = _list_rgsr46.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[46].month10 = _list_rgsr46.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[46].month11 = _list_rgsr46.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[46].month12 = _list_rgsr46.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[47].month1 = _list_rgsr47.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[47].month2 = _list_rgsr47.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[47].month3 = _list_rgsr47.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[47].month4 = _list_rgsr47.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[47].month5 = _list_rgsr47.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[47].month6 = _list_rgsr47.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[47].month7 = _list_rgsr47.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[47].month8 = _list_rgsr47.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[47].month9 = _list_rgsr47.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[47].month10 = _list_rgsr47.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[47].month11 = _list_rgsr47.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[47].month12 = _list_rgsr47.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();


                        list[49].month1 = _list_rgsj49.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[49].month2 = _list_rgsj49.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[49].month3 = _list_rgsj49.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[49].month4 = _list_rgzc49.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[49].month5 = _list_rgzc49.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[49].month6 = _list_rgzc49.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[49].month7 = _list_rgzc49.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[49].month8 = _list_rgzc49.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[49].month9 = _list_rgzc49.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[49].month10 = _list_rgzc49.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[49].month11 = _list_rgzc49.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[49].month12 = _list_rgzc49.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[50].month1 = _list_rgsj50.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[50].month2 = _list_rgsj50.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[50].month3 = _list_rgsj50.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[50].month4 = _list_rgzc50.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[50].month5 = _list_rgzc50.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[50].month6 = _list_rgzc50.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[50].month7 = _list_rgzc50.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[50].month8 = _list_rgzc50.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[50].month9 = _list_rgzc50.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[50].month10 = _list_rgzc50.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[50].month11 = _list_rgzc50.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[50].month12 = _list_rgzc50.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[51].month1 = _list_rgsj51.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[51].month2 = _list_rgsj51.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[51].month3 = _list_rgsj51.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[51].month4 = _list_rgzc51.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[51].month5 = _list_rgzc51.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[51].month6 = _list_rgzc51.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[51].month7 = _list_rgzc51.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[51].month8 = _list_rgzc51.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[51].month9 = _list_rgzc51.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[51].month10 = _list_rgzc51.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[51].month11 = _list_rgzc51.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[51].month12 = _list_rgzc51.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[52].month1 = _list_rgsj52.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[52].month2 = _list_rgsj52.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[52].month3 = _list_rgsj52.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[52].month4 = _list_rgzc52.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[52].month5 = _list_rgzc52.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[52].month6 = _list_rgzc52.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[52].month7 = _list_rgzc52.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[52].month8 = _list_rgzc52.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[52].month9 = _list_rgzc52.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[52].month10 = _list_rgzc52.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[52].month11 = _list_rgzc52.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[52].month12 = _list_rgzc52.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[53].month1 = _list_rgsj53.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[53].month2 = _list_rgsj53.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[53].month3 = _list_rgsj53.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[53].month4 = _list_rgzc53.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[53].month5 = _list_rgzc53.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[53].month6 = _list_rgzc53.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[53].month7 = _list_rgzc53.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[53].month8 = _list_rgzc53.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[53].month9 = _list_rgzc53.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[53].month10 = _list_rgzc53.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[53].month11 = _list_rgzc53.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[53].month12 = _list_rgzc53.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[54].month1 = _list_rgsj54.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[54].month2 = _list_rgsj54.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[54].month3 = _list_rgsj54.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[54].month4 = _list_rgzc54.Where(p => p.monthly == 4).Sum(p => p.proBudget).ToString();
                        list[54].month5 = _list_rgzc54.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[54].month6 = _list_rgzc54.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[54].month7 = _list_rgzc54.Where(p => p.monthly == 7).Sum(p => p.adjustBudget).ToString();
                        list[54].month8 = _list_rgzc54.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[54].month9 = _list_rgzc54.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[54].month10 = _list_rgzc54.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[54].month11 = _list_rgzc54.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[54].month12 = _list_rgzc54.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();
                        #endregion
                        break;
                    case 5:
                        #region
                        list[42].month1 = _list_rgsr42.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[42].month2 = _list_rgsr42.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[42].month3 = _list_rgsr42.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[42].month4 = _list_rgsr42.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[42].month5 = _list_rgsr42.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[42].month6 = _list_rgsr42.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[42].month7 = _list_rgsr42.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[42].month8 = _list_rgsr42.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[42].month9 = _list_rgsr42.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[42].month10 = _list_rgsr42.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[42].month11 = _list_rgsr42.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[42].month12 = _list_rgsr42.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[43].month1 = _list_rgsr43.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[43].month2 = _list_rgsr43.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[43].month3 = _list_rgsr43.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[43].month4 = _list_rgsr43.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[43].month5 = _list_rgsr43.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[43].month6 = _list_rgsr43.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[43].month7 = _list_rgsr43.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[43].month8 = _list_rgsr43.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[43].month9 = _list_rgsr43.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[43].month10 = _list_rgsr43.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[43].month11 = _list_rgsr43.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[43].month12 = _list_rgsr43.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[44].month1 = _list_rgsr44.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[44].month2 = _list_rgsr44.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[44].month3 = _list_rgsr44.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[44].month4 = _list_rgsr44.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[44].month5 = _list_rgsr44.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[44].month6 = _list_rgsr44.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[44].month7 = _list_rgsr44.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[44].month8 = _list_rgsr44.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[44].month9 = _list_rgsr44.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[44].month10 = _list_rgsr44.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[44].month11 = _list_rgsr44.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[44].month12 = _list_rgsr44.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[45].month1 = _list_rgsr45.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[45].month2 = _list_rgsr45.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[45].month3 = _list_rgsr45.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[45].month4 = _list_rgsr45.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[45].month5 = _list_rgsr45.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[45].month6 = _list_rgsr45.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[45].month7 = _list_rgsr45.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[45].month8 = _list_rgsr45.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[45].month9 = _list_rgsr45.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[45].month10 = _list_rgsr45.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[45].month11 = _list_rgsr45.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[45].month12 = _list_rgsr45.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[46].month1 = _list_rgsr46.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[46].month2 = _list_rgsr46.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[46].month3 = _list_rgsr46.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[46].month4 = _list_rgsr46.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[46].month5 = _list_rgsr46.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[46].month6 = _list_rgsr46.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[46].month7 = _list_rgsr46.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[46].month8 = _list_rgsr46.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[46].month9 = _list_rgsr46.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[46].month10 = _list_rgsr46.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[46].month11 = _list_rgsr46.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[46].month12 = _list_rgsr46.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[47].month1 = _list_rgsr47.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[47].month2 = _list_rgsr47.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[47].month3 = _list_rgsr47.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[47].month4 = _list_rgsr47.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[47].month5 = _list_rgsr47.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[47].month6 = _list_rgsr47.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[47].month7 = _list_rgsr47.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[47].month8 = _list_rgsr47.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[47].month9 = _list_rgsr47.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[47].month10 = _list_rgsr47.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[47].month11 = _list_rgsr47.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[47].month12 = _list_rgsr47.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();


                        list[49].month1 = _list_rgsj49.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[49].month2 = _list_rgsj49.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[49].month3 = _list_rgsj49.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[49].month4 = _list_rgsj49.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[49].month5 = _list_rgzc49.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[49].month6 = _list_rgzc49.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[49].month7 = _list_rgzc49.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[49].month8 = _list_rgzc49.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[49].month9 = _list_rgzc49.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[49].month10 = _list_rgzc49.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[49].month11 = _list_rgzc49.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[49].month12 = _list_rgzc49.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[50].month1 = _list_rgsj50.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[50].month2 = _list_rgsj50.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[50].month3 = _list_rgsj50.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[50].month4 = _list_rgsj50.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[50].month5 = _list_rgzc50.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[50].month6 = _list_rgzc50.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[50].month7 = _list_rgzc50.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[50].month8 = _list_rgzc50.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[50].month9 = _list_rgzc50.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[50].month10 = _list_rgzc50.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[50].month11 = _list_rgzc50.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[50].month12 = _list_rgzc50.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[51].month1 = _list_rgsj51.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[51].month2 = _list_rgsj51.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[51].month3 = _list_rgsj51.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[51].month4 = _list_rgsj51.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[51].month5 = _list_rgzc51.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[51].month6 = _list_rgzc51.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[51].month7 = _list_rgzc51.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[51].month8 = _list_rgzc51.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[51].month9 = _list_rgzc51.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[51].month10 = _list_rgzc51.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[51].month11 = _list_rgzc51.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[51].month12 = _list_rgzc51.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[52].month1 = _list_rgsj52.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[52].month2 = _list_rgsj52.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[52].month3 = _list_rgsj52.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[52].month4 = _list_rgsj52.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[52].month5 = _list_rgzc52.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[52].month6 = _list_rgzc52.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[52].month7 = _list_rgzc52.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[52].month8 = _list_rgzc52.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[52].month9 = _list_rgzc52.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[52].month10 = _list_rgzc52.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[52].month11 = _list_rgzc52.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[52].month12 = _list_rgzc52.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[53].month1 = _list_rgsj53.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[53].month2 = _list_rgsj53.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[53].month3 = _list_rgsj53.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[53].month4 = _list_rgsj53.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[53].month5 = _list_rgzc53.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[53].month6 = _list_rgzc53.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[53].month7 = _list_rgzc53.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[53].month8 = _list_rgzc53.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[53].month9 = _list_rgzc53.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[53].month10 = _list_rgzc53.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[53].month11 = _list_rgzc53.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[53].month12 = _list_rgzc53.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[54].month1 = _list_rgsj54.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[54].month2 = _list_rgsj54.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[54].month3 = _list_rgsj54.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[54].month4 = _list_rgsj54.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[54].month5 = _list_rgzc54.Where(p => p.monthly == 5).Sum(p => p.proBudget).ToString();
                        list[54].month6 = _list_rgzc54.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[54].month7 = _list_rgzc54.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[54].month8 = _list_rgzc54.Where(p => p.monthly == 8).Sum(p => p.adjustBudget).ToString();
                        list[54].month9 = _list_rgzc54.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[54].month10 = _list_rgzc54.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[54].month11 = _list_rgzc54.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[54].month12 = _list_rgzc54.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();
                        #endregion
                        break;
                    case 6:
                        #region
                        list[42].month1 = _list_rgsr42.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[42].month2 = _list_rgsr42.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[42].month3 = _list_rgsr42.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[42].month4 = _list_rgsr42.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[42].month5 = _list_rgsr42.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[42].month6 = _list_rgsr42.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[42].month7 = _list_rgsr42.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[42].month8 = _list_rgsr42.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[42].month9 = _list_rgsr42.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[42].month10 = _list_rgsr42.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[42].month11 = _list_rgsr42.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[42].month12 = _list_rgsr42.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[43].month1 = _list_rgsr43.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[43].month2 = _list_rgsr43.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[43].month3 = _list_rgsr43.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[43].month4 = _list_rgsr43.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[43].month5 = _list_rgsr43.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[43].month6 = _list_rgsr43.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[43].month7 = _list_rgsr43.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[43].month8 = _list_rgsr43.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[43].month9 = _list_rgsr43.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[43].month10 = _list_rgsr43.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[43].month11 = _list_rgsr43.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[43].month12 = _list_rgsr43.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[44].month1 = _list_rgsr44.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[44].month2 = _list_rgsr44.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[44].month3 = _list_rgsr44.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[44].month4 = _list_rgsr44.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[44].month5 = _list_rgsr44.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[44].month6 = _list_rgsr44.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[44].month7 = _list_rgsr44.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[44].month8 = _list_rgsr44.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[44].month9 = _list_rgsr44.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[44].month10 = _list_rgsr44.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[44].month11 = _list_rgsr44.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[44].month12 = _list_rgsr44.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[45].month1 = _list_rgsr45.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[45].month2 = _list_rgsr45.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[45].month3 = _list_rgsr45.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[45].month4 = _list_rgsr45.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[45].month5 = _list_rgsr45.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[45].month6 = _list_rgsr45.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[45].month7 = _list_rgsr45.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[45].month8 = _list_rgsr45.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[45].month9 = _list_rgsr45.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[45].month10 = _list_rgsr45.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[45].month11 = _list_rgsr45.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[45].month12 = _list_rgsr45.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[46].month1 = _list_rgsr46.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[46].month2 = _list_rgsr46.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[46].month3 = _list_rgsr46.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[46].month4 = _list_rgsr46.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[46].month5 = _list_rgsr46.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[46].month6 = _list_rgsr46.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[46].month7 = _list_rgsr46.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[46].month8 = _list_rgsr46.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[46].month9 = _list_rgsr46.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[46].month10 = _list_rgsr46.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[46].month11 = _list_rgsr46.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[46].month12 = _list_rgsr46.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[47].month1 = _list_rgsr47.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[47].month2 = _list_rgsr47.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[47].month3 = _list_rgsr47.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[47].month4 = _list_rgsr47.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[47].month5 = _list_rgsr47.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[47].month6 = _list_rgsr47.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[47].month7 = _list_rgsr47.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[47].month8 = _list_rgsr47.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[47].month9 = _list_rgsr47.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[47].month10 = _list_rgsr47.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[47].month11 = _list_rgsr47.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[47].month12 = _list_rgsr47.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();


                        list[49].month1 = _list_rgsj49.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[49].month2 = _list_rgsj49.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[49].month3 = _list_rgsj49.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[49].month4 = _list_rgsj49.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[49].month5 = _list_rgsj49.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[49].month6 = _list_rgzc49.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[49].month7 = _list_rgzc49.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[49].month8 = _list_rgzc49.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[49].month9 = _list_rgzc49.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[49].month10 = _list_rgzc49.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[49].month11 = _list_rgzc49.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[49].month12 = _list_rgzc49.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[50].month1 = _list_rgsj50.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[50].month2 = _list_rgsj50.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[50].month3 = _list_rgsj50.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[50].month4 = _list_rgsj50.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[50].month5 = _list_rgsj50.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[50].month6 = _list_rgzc50.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[50].month7 = _list_rgzc50.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[50].month8 = _list_rgzc50.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[50].month9 = _list_rgzc50.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[50].month10 = _list_rgzc50.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[50].month11 = _list_rgzc50.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[50].month12 = _list_rgzc50.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[51].month1 = _list_rgsj51.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[51].month2 = _list_rgsj51.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[51].month3 = _list_rgsj51.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[51].month4 = _list_rgsj51.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[51].month5 = _list_rgsj51.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[51].month6 = _list_rgzc51.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[51].month7 = _list_rgzc51.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[51].month8 = _list_rgzc51.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[51].month9 = _list_rgzc51.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[51].month10 = _list_rgzc51.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[51].month11 = _list_rgzc51.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[51].month12 = _list_rgzc51.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[52].month1 = _list_rgsj52.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[52].month2 = _list_rgsj52.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[52].month3 = _list_rgsj52.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[52].month4 = _list_rgsj52.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[52].month5 = _list_rgsj52.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[52].month6 = _list_rgzc52.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[52].month7 = _list_rgzc52.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[52].month8 = _list_rgzc52.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[52].month9 = _list_rgzc52.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[52].month10 = _list_rgzc52.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[52].month11 = _list_rgzc52.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[52].month12 = _list_rgzc52.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[53].month1 = _list_rgsj53.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[53].month2 = _list_rgsj53.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[53].month3 = _list_rgsj53.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[53].month4 = _list_rgsj53.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[53].month5 = _list_rgsj53.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[53].month6 = _list_rgzc53.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[53].month7 = _list_rgzc53.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[53].month8 = _list_rgzc53.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[53].month9 = _list_rgzc53.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[53].month10 = _list_rgzc53.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[53].month11 = _list_rgzc53.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[53].month12 = _list_rgzc53.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[54].month1 = _list_rgsj54.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[54].month2 = _list_rgsj54.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[54].month3 = _list_rgsj54.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[54].month4 = _list_rgsj54.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[54].month5 = _list_rgsj54.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[54].month6 = _list_rgzc54.Where(p => p.monthly == 6).Sum(p => p.proBudget).ToString();
                        list[54].month7 = _list_rgzc54.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[54].month8 = _list_rgzc54.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[54].month9 = _list_rgzc54.Where(p => p.monthly == 9).Sum(p => p.adjustBudget).ToString();
                        list[54].month10 = _list_rgzc54.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[54].month11 = _list_rgzc54.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[54].month12 = _list_rgzc54.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();
                        #endregion
                        break;
                    case 7:
                        #region
                        list[42].month1 = _list_rgsr42.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[42].month2 = _list_rgsr42.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[42].month3 = _list_rgsr42.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[42].month4 = _list_rgsr42.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[42].month5 = _list_rgsr42.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[42].month6 = _list_rgsr42.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[42].month7 = _list_rgsr42.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[42].month8 = _list_rgsr42.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[42].month9 = _list_rgsr42.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[42].month10 = _list_rgsr42.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[42].month11 = _list_rgsr42.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[42].month12 = _list_rgsr42.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[43].month1 = _list_rgsr43.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[43].month2 = _list_rgsr43.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[43].month3 = _list_rgsr43.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[43].month4 = _list_rgsr43.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[43].month5 = _list_rgsr43.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[43].month6 = _list_rgsr43.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[43].month7 = _list_rgsr43.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[43].month8 = _list_rgsr43.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[43].month9 = _list_rgsr43.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[43].month10 = _list_rgsr43.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[43].month11 = _list_rgsr43.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[43].month12 = _list_rgsr43.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[44].month1 = _list_rgsr44.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[44].month2 = _list_rgsr44.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[44].month3 = _list_rgsr44.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[44].month4 = _list_rgsr44.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[44].month5 = _list_rgsr44.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[44].month6 = _list_rgsr44.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[44].month7 = _list_rgsr44.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[44].month8 = _list_rgsr44.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[44].month9 = _list_rgsr44.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[44].month10 = _list_rgsr44.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[44].month11 = _list_rgsr44.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[44].month12 = _list_rgsr44.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[45].month1 = _list_rgsr45.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[45].month2 = _list_rgsr45.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[45].month3 = _list_rgsr45.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[45].month4 = _list_rgsr45.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[45].month5 = _list_rgsr45.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[45].month6 = _list_rgsr45.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[45].month7 = _list_rgsr45.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[45].month8 = _list_rgsr45.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[45].month9 = _list_rgsr45.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[45].month10 = _list_rgsr45.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[45].month11 = _list_rgsr45.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[45].month12 = _list_rgsr45.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[46].month1 = _list_rgsr46.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[46].month2 = _list_rgsr46.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[46].month3 = _list_rgsr46.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[46].month4 = _list_rgsr46.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[46].month5 = _list_rgsr46.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[46].month6 = _list_rgsr46.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[46].month7 = _list_rgsr46.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[46].month8 = _list_rgsr46.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[46].month9 = _list_rgsr46.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[46].month10 = _list_rgsr46.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[46].month11 = _list_rgsr46.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[46].month12 = _list_rgsr46.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[47].month1 = _list_rgsr47.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[47].month2 = _list_rgsr47.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[47].month3 = _list_rgsr47.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[47].month4 = _list_rgsr47.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[47].month5 = _list_rgsr47.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[47].month6 = _list_rgsr47.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[47].month7 = _list_rgsr47.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[47].month8 = _list_rgsr47.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[47].month9 = _list_rgsr47.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[47].month10 = _list_rgsr47.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[47].month11 = _list_rgsr47.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[47].month12 = _list_rgsr47.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();


                        list[49].month1 = _list_rgsj49.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[49].month2 = _list_rgsj49.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[49].month3 = _list_rgsj49.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[49].month4 = _list_rgsj49.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[49].month5 = _list_rgsj49.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[49].month6 = _list_rgsj49.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[49].month7 = _list_rgzc49.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[49].month8 = _list_rgzc49.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[49].month9 = _list_rgzc49.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[49].month10 = _list_rgzc49.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[49].month11 = _list_rgzc49.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[49].month12 = _list_rgzc49.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[50].month1 = _list_rgsj50.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[50].month2 = _list_rgsj50.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[50].month3 = _list_rgsj50.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[50].month4 = _list_rgsj50.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[50].month5 = _list_rgsj50.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[50].month6 = _list_rgsj50.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[50].month7 = _list_rgzc50.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[50].month8 = _list_rgzc50.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[50].month9 = _list_rgzc50.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[50].month10 = _list_rgzc50.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[50].month11 = _list_rgzc50.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[50].month12 = _list_rgzc50.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[51].month1 = _list_rgsj51.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[51].month2 = _list_rgsj51.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[51].month3 = _list_rgsj51.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[51].month4 = _list_rgsj51.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[51].month5 = _list_rgsj51.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[51].month6 = _list_rgsj51.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[51].month7 = _list_rgzc51.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[51].month8 = _list_rgzc51.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[51].month9 = _list_rgzc51.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[51].month10 = _list_rgzc51.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[51].month11 = _list_rgzc51.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[51].month12 = _list_rgzc51.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[52].month1 = _list_rgsj52.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[52].month2 = _list_rgsj52.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[52].month3 = _list_rgsj52.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[52].month4 = _list_rgsj52.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[52].month5 = _list_rgsj52.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[52].month6 = _list_rgsj52.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[52].month7 = _list_rgzc52.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[52].month8 = _list_rgzc52.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[52].month9 = _list_rgzc52.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[52].month10 = _list_rgzc52.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[52].month11 = _list_rgzc52.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[52].month12 = _list_rgzc52.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[53].month1 = _list_rgsj53.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[53].month2 = _list_rgsj53.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[53].month3 = _list_rgsj53.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[53].month4 = _list_rgsj53.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[53].month5 = _list_rgsj53.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[53].month6 = _list_rgsj53.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[53].month7 = _list_rgzc53.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[53].month8 = _list_rgzc53.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[53].month9 = _list_rgzc53.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[53].month10 = _list_rgzc53.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[53].month11 = _list_rgzc53.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[53].month12 = _list_rgzc53.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[54].month1 = _list_rgsj54.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[54].month2 = _list_rgsj54.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[54].month3 = _list_rgsj54.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[54].month4 = _list_rgsj54.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[54].month5 = _list_rgsj54.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[54].month6 = _list_rgsj54.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[54].month7 = _list_rgzc54.Where(p => p.monthly == 7).Sum(p => p.proBudget).ToString();
                        list[54].month8 = _list_rgzc54.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[54].month9 = _list_rgzc54.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[54].month10 = _list_rgzc54.Where(p => p.monthly == 10).Sum(p => p.adjustBudget).ToString();
                        list[54].month11 = _list_rgzc54.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[54].month12 = _list_rgzc54.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();
                        #endregion
                        break;
                    case 8:
                        #region
                        list[42].month1 = _list_rgsr42.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[42].month2 = _list_rgsr42.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[42].month3 = _list_rgsr42.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[42].month4 = _list_rgsr42.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[42].month5 = _list_rgsr42.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[42].month6 = _list_rgsr42.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[42].month7 = _list_rgsr42.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[42].month8 = _list_rgsr42.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[42].month9 = _list_rgsr42.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[42].month10 = _list_rgsr42.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[42].month11 = _list_rgsr42.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[42].month12 = _list_rgsr42.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[43].month1 = _list_rgsr43.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[43].month2 = _list_rgsr43.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[43].month3 = _list_rgsr43.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[43].month4 = _list_rgsr43.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[43].month5 = _list_rgsr43.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[43].month6 = _list_rgsr43.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[43].month7 = _list_rgsr43.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[43].month8 = _list_rgsr43.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[43].month9 = _list_rgsr43.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[43].month10 = _list_rgsr43.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[43].month11 = _list_rgsr43.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[43].month12 = _list_rgsr43.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[44].month1 = _list_rgsr44.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[44].month2 = _list_rgsr44.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[44].month3 = _list_rgsr44.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[44].month4 = _list_rgsr44.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[44].month5 = _list_rgsr44.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[44].month6 = _list_rgsr44.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[44].month7 = _list_rgsr44.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[44].month8 = _list_rgsr44.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[44].month9 = _list_rgsr44.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[44].month10 = _list_rgsr44.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[44].month11 = _list_rgsr44.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[44].month12 = _list_rgsr44.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[45].month1 = _list_rgsr45.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[45].month2 = _list_rgsr45.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[45].month3 = _list_rgsr45.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[45].month4 = _list_rgsr45.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[45].month5 = _list_rgsr45.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[45].month6 = _list_rgsr45.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[45].month7 = _list_rgsr45.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[45].month8 = _list_rgsr45.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[45].month9 = _list_rgsr45.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[45].month10 = _list_rgsr45.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[45].month11 = _list_rgsr45.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[45].month12 = _list_rgsr45.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[46].month1 = _list_rgsr46.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[46].month2 = _list_rgsr46.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[46].month3 = _list_rgsr46.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[46].month4 = _list_rgsr46.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[46].month5 = _list_rgsr46.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[46].month6 = _list_rgsr46.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[46].month7 = _list_rgsr46.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[46].month8 = _list_rgsr46.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[46].month9 = _list_rgsr46.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[46].month10 = _list_rgsr46.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[46].month11 = _list_rgsr46.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[46].month12 = _list_rgsr46.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[47].month1 = _list_rgsr47.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[47].month2 = _list_rgsr47.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[47].month3 = _list_rgsr47.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[47].month4 = _list_rgsr47.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[47].month5 = _list_rgsr47.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[47].month6 = _list_rgsr47.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[47].month7 = _list_rgsr47.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[47].month8 = _list_rgsr47.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[47].month9 = _list_rgsr47.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[47].month10 = _list_rgsr47.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[47].month11 = _list_rgsr47.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[47].month12 = _list_rgsr47.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();


                        list[49].month1 = _list_rgsj49.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[49].month2 = _list_rgsj49.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[49].month3 = _list_rgsj49.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[49].month4 = _list_rgsj49.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[49].month5 = _list_rgsj49.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[49].month6 = _list_rgsj49.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[49].month7 = _list_rgsj49.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[49].month8 = _list_rgzc49.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[49].month9 = _list_rgzc49.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[49].month10 = _list_rgzc49.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[49].month11 = _list_rgzc49.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[49].month12 = _list_rgzc49.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[50].month1 = _list_rgsj50.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[50].month2 = _list_rgsj50.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[50].month3 = _list_rgsj50.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[50].month4 = _list_rgsj50.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[50].month5 = _list_rgsj50.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[50].month6 = _list_rgsj50.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[50].month7 = _list_rgsj50.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[50].month8 = _list_rgzc50.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[50].month9 = _list_rgzc50.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[50].month10 = _list_rgzc50.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[50].month11 = _list_rgzc50.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[50].month12 = _list_rgzc50.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[51].month1 = _list_rgsj51.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[51].month2 = _list_rgsj51.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[51].month3 = _list_rgsj51.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[51].month4 = _list_rgsj51.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[51].month5 = _list_rgsj51.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[51].month6 = _list_rgsj51.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[51].month7 = _list_rgsj51.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[51].month8 = _list_rgzc51.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[51].month9 = _list_rgzc51.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[51].month10 = _list_rgzc51.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[51].month11 = _list_rgzc51.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[51].month12 = _list_rgzc51.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[52].month1 = _list_rgsj52.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[52].month2 = _list_rgsj52.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[52].month3 = _list_rgsj52.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[52].month4 = _list_rgsj52.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[52].month5 = _list_rgsj52.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[52].month6 = _list_rgsj52.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[52].month7 = _list_rgsj52.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[52].month8 = _list_rgzc52.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[52].month9 = _list_rgzc52.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[52].month10 = _list_rgzc52.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[52].month11 = _list_rgzc52.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[52].month12 = _list_rgzc52.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[53].month1 = _list_rgsj53.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[53].month2 = _list_rgsj53.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[53].month3 = _list_rgsj53.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[53].month4 = _list_rgsj53.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[53].month5 = _list_rgsj53.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[53].month6 = _list_rgsj53.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[53].month7 = _list_rgsj53.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[53].month8 = _list_rgzc53.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[53].month9 = _list_rgzc53.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[53].month10 = _list_rgzc53.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[53].month11 = _list_rgzc53.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[53].month12 = _list_rgzc53.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[54].month1 = _list_rgsj54.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[54].month2 = _list_rgsj54.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[54].month3 = _list_rgsj54.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[54].month4 = _list_rgsj54.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[54].month5 = _list_rgsj54.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[54].month6 = _list_rgsj54.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[54].month7 = _list_rgsj54.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[54].month8 = _list_rgzc54.Where(p => p.monthly == 8).Sum(p => p.proBudget).ToString();
                        list[54].month9 = _list_rgzc54.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[54].month10 = _list_rgzc54.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[54].month11 = _list_rgzc54.Where(p => p.monthly == 11).Sum(p => p.adjustBudget).ToString();
                        list[54].month12 = _list_rgzc54.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();
                        #endregion
                        break;
                    case 9:
                        #region
                        list[42].month1 = _list_rgsr42.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[42].month2 = _list_rgsr42.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[42].month3 = _list_rgsr42.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[42].month4 = _list_rgsr42.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[42].month5 = _list_rgsr42.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[42].month6 = _list_rgsr42.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[42].month7 = _list_rgsr42.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[42].month8 = _list_rgsr42.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[42].month9 = _list_rgsr42.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[42].month10 = _list_rgsr42.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[42].month11 = _list_rgsr42.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[42].month12 = _list_rgsr42.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[43].month1 = _list_rgsr43.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[43].month2 = _list_rgsr43.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[43].month3 = _list_rgsr43.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[43].month4 = _list_rgsr43.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[43].month5 = _list_rgsr43.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[43].month6 = _list_rgsr43.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[43].month7 = _list_rgsr43.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[43].month8 = _list_rgsr43.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[43].month9 = _list_rgsr43.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[43].month10 = _list_rgsr43.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[43].month11 = _list_rgsr43.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[43].month12 = _list_rgsr43.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[44].month1 = _list_rgsr44.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[44].month2 = _list_rgsr44.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[44].month3 = _list_rgsr44.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[44].month4 = _list_rgsr44.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[44].month5 = _list_rgsr44.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[44].month6 = _list_rgsr44.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[44].month7 = _list_rgsr44.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[44].month8 = _list_rgsr44.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[44].month9 = _list_rgsr44.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[44].month10 = _list_rgsr44.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[44].month11 = _list_rgsr44.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[44].month12 = _list_rgsr44.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[45].month1 = _list_rgsr45.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[45].month2 = _list_rgsr45.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[45].month3 = _list_rgsr45.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[45].month4 = _list_rgsr45.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[45].month5 = _list_rgsr45.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[45].month6 = _list_rgsr45.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[45].month7 = _list_rgsr45.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[45].month8 = _list_rgsr45.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[45].month9 = _list_rgsr45.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[45].month10 = _list_rgsr45.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[45].month11 = _list_rgsr45.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[45].month12 = _list_rgsr45.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[46].month1 = _list_rgsr46.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[46].month2 = _list_rgsr46.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[46].month3 = _list_rgsr46.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[46].month4 = _list_rgsr46.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[46].month5 = _list_rgsr46.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[46].month6 = _list_rgsr46.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[46].month7 = _list_rgsr46.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[46].month8 = _list_rgsr46.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[46].month9 = _list_rgsr46.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[46].month10 = _list_rgsr46.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[46].month11 = _list_rgsr46.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[46].month12 = _list_rgsr46.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[47].month1 = _list_rgsr47.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[47].month2 = _list_rgsr47.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[47].month3 = _list_rgsr47.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[47].month4 = _list_rgsr47.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[47].month5 = _list_rgsr47.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[47].month6 = _list_rgsr47.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[47].month7 = _list_rgsr47.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[47].month8 = _list_rgsr47.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[47].month9 = _list_rgsr47.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[47].month10 = _list_rgsr47.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[47].month11 = _list_rgsr47.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[47].month12 = _list_rgsr47.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();


                        list[49].month1 = _list_rgsj49.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[49].month2 = _list_rgsj49.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[49].month3 = _list_rgsj49.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[49].month4 = _list_rgsj49.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[49].month5 = _list_rgsj49.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[49].month6 = _list_rgsj49.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[49].month7 = _list_rgsj49.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[49].month8 = _list_rgsj49.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[49].month9 = _list_rgzc49.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[49].month10 = _list_rgzc49.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[49].month11 = _list_rgzc49.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[49].month12 = _list_rgzc49.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[50].month1 = _list_rgsj50.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[50].month2 = _list_rgsj50.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[50].month3 = _list_rgsj50.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[50].month4 = _list_rgsj50.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[50].month5 = _list_rgsj50.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[50].month6 = _list_rgsj50.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[50].month7 = _list_rgsj50.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[50].month8 = _list_rgsj50.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[50].month9 = _list_rgzc50.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[50].month10 = _list_rgzc50.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[50].month11 = _list_rgzc50.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[50].month12 = _list_rgzc50.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[51].month1 = _list_rgsj51.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[51].month2 = _list_rgsj51.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[51].month3 = _list_rgsj51.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[51].month4 = _list_rgsj51.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[51].month5 = _list_rgsj51.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[51].month6 = _list_rgsj51.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[51].month7 = _list_rgsj51.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[51].month8 = _list_rgsj51.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[51].month9 = _list_rgzc51.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[51].month10 = _list_rgzc51.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[51].month11 = _list_rgzc51.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[51].month12 = _list_rgzc51.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[52].month1 = _list_rgsj52.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[52].month2 = _list_rgsj52.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[52].month3 = _list_rgsj52.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[52].month4 = _list_rgsj52.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[52].month5 = _list_rgsj52.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[52].month6 = _list_rgsj52.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[52].month7 = _list_rgsj52.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[52].month8 = _list_rgsj52.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[52].month9 = _list_rgzc52.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[52].month10 = _list_rgzc52.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[52].month11 = _list_rgzc52.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[52].month12 = _list_rgzc52.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[53].month1 = _list_rgsj53.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[53].month2 = _list_rgsj53.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[53].month3 = _list_rgsj53.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[53].month4 = _list_rgsj53.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[53].month5 = _list_rgsj53.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[53].month6 = _list_rgsj53.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[53].month7 = _list_rgsj53.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[53].month8 = _list_rgsj53.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[53].month9 = _list_rgzc53.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[53].month10 = _list_rgzc53.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[53].month11 = _list_rgzc53.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[53].month12 = _list_rgzc53.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();

                        list[54].month1 = _list_rgsj54.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[54].month2 = _list_rgsj54.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[54].month3 = _list_rgsj54.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[54].month4 = _list_rgsj54.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[54].month5 = _list_rgsj54.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[54].month6 = _list_rgsj54.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[54].month7 = _list_rgsj54.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[54].month8 = _list_rgsj54.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[54].month9 = _list_rgzc54.Where(p => p.monthly == 9).Sum(p => p.proBudget).ToString();
                        list[54].month10 = _list_rgzc54.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[54].month11 = _list_rgzc54.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[54].month12 = _list_rgzc54.Where(p => p.monthly == 12).Sum(p => p.adjustBudget).ToString();
                        #endregion
                        break;
                    case 10:
                        #region
                        list[42].month1 = _list_rgsr42.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[42].month2 = _list_rgsr42.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[42].month3 = _list_rgsr42.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[42].month4 = _list_rgsr42.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[42].month5 = _list_rgsr42.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[42].month6 = _list_rgsr42.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[42].month7 = _list_rgsr42.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[42].month8 = _list_rgsr42.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[42].month9 = _list_rgsr42.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[42].month10 = _list_rgsr42.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[42].month11 = _list_rgsr42.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[42].month12 = _list_rgsr42.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[43].month1 = _list_rgsr43.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[43].month2 = _list_rgsr43.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[43].month3 = _list_rgsr43.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[43].month4 = _list_rgsr43.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[43].month5 = _list_rgsr43.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[43].month6 = _list_rgsr43.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[43].month7 = _list_rgsr43.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[43].month8 = _list_rgsr43.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[43].month9 = _list_rgsr43.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[43].month10 = _list_rgsr43.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[43].month11 = _list_rgsr43.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[43].month12 = _list_rgsr43.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[44].month1 = _list_rgsr44.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[44].month2 = _list_rgsr44.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[44].month3 = _list_rgsr44.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[44].month4 = _list_rgsr44.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[44].month5 = _list_rgsr44.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[44].month6 = _list_rgsr44.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[44].month7 = _list_rgsr44.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[44].month8 = _list_rgsr44.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[44].month9 = _list_rgsr44.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[44].month10 = _list_rgsr44.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[44].month11 = _list_rgsr44.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[44].month12 = _list_rgsr44.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[45].month1 = _list_rgsr45.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[45].month2 = _list_rgsr45.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[45].month3 = _list_rgsr45.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[45].month4 = _list_rgsr45.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[45].month5 = _list_rgsr45.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[45].month6 = _list_rgsr45.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[45].month7 = _list_rgsr45.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[45].month8 = _list_rgsr45.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[45].month9 = _list_rgsr45.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[45].month10 = _list_rgsr45.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[45].month11 = _list_rgsr45.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[45].month12 = _list_rgsr45.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[46].month1 = _list_rgsr46.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[46].month2 = _list_rgsr46.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[46].month3 = _list_rgsr46.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[46].month4 = _list_rgsr46.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[46].month5 = _list_rgsr46.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[46].month6 = _list_rgsr46.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[46].month7 = _list_rgsr46.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[46].month8 = _list_rgsr46.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[46].month9 = _list_rgsr46.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[46].month10 = _list_rgsr46.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[46].month11 = _list_rgsr46.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[46].month12 = _list_rgsr46.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[47].month1 = _list_rgsr47.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[47].month2 = _list_rgsr47.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[47].month3 = _list_rgsr47.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[47].month4 = _list_rgsr47.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[47].month5 = _list_rgsr47.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[47].month6 = _list_rgsr47.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[47].month7 = _list_rgsr47.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[47].month8 = _list_rgsr47.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[47].month9 = _list_rgsr47.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[47].month10 = _list_rgsr47.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[47].month11 = _list_rgsr47.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[47].month12 = _list_rgsr47.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();


                        list[49].month1 = _list_rgsj49.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[49].month2 = _list_rgsj49.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[49].month3 = _list_rgsj49.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[49].month4 = _list_rgsj49.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[49].month5 = _list_rgsj49.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[49].month6 = _list_rgsj49.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[49].month7 = _list_rgsj49.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[49].month8 = _list_rgsj49.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[49].month9 = _list_rgsj49.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[49].month10 = _list_rgzc49.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[49].month11 = _list_rgzc49.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[49].month12 = _list_rgzc49.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[50].month1 = _list_rgsj50.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[50].month2 = _list_rgsj50.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[50].month3 = _list_rgsj50.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[50].month4 = _list_rgsj50.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[50].month5 = _list_rgsj50.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[50].month6 = _list_rgsj50.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[50].month7 = _list_rgsj50.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[50].month8 = _list_rgsj50.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[50].month9 = _list_rgsj50.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[50].month10 = _list_rgzc50.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[50].month11 = _list_rgzc50.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[50].month12 = _list_rgzc50.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[51].month1 = _list_rgsj51.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[51].month2 = _list_rgsj51.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[51].month3 = _list_rgsj51.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[51].month4 = _list_rgsj51.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[51].month5 = _list_rgsj51.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[51].month6 = _list_rgsj51.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[51].month7 = _list_rgsj51.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[51].month8 = _list_rgsj51.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[51].month9 = _list_rgsj51.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[51].month10 = _list_rgzc51.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[51].month11 = _list_rgzc51.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[51].month12 = _list_rgzc51.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[52].month1 = _list_rgsj52.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[52].month2 = _list_rgsj52.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[52].month3 = _list_rgsj52.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[52].month4 = _list_rgsj52.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[52].month5 = _list_rgsj52.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[52].month6 = _list_rgsj52.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[52].month7 = _list_rgsj52.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[52].month8 = _list_rgsj52.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[52].month9 = _list_rgsj52.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[52].month10 = _list_rgzc52.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[52].month11 = _list_rgzc52.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[52].month12 = _list_rgzc52.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[53].month1 = _list_rgsj53.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[53].month2 = _list_rgsj53.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[53].month3 = _list_rgsj53.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[53].month4 = _list_rgsj53.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[53].month5 = _list_rgsj53.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[53].month6 = _list_rgsj53.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[53].month7 = _list_rgsj53.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[53].month8 = _list_rgsj53.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[53].month9 = _list_rgsj53.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[53].month10 = _list_rgzc53.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[53].month11 = _list_rgzc53.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[53].month12 = _list_rgzc53.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[54].month1 = _list_rgsj54.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[54].month2 = _list_rgsj54.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[54].month3 = _list_rgsj54.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[54].month4 = _list_rgsj54.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[54].month5 = _list_rgsj54.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[54].month6 = _list_rgsj54.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[54].month7 = _list_rgsj54.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[54].month8 = _list_rgsj54.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[54].month9 = _list_rgsj54.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[54].month10 = _list_rgzc54.Where(p => p.monthly == 10).Sum(p => p.proBudget).ToString();
                        list[54].month11 = _list_rgzc54.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[54].month12 = _list_rgzc54.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();
                        #endregion
                        break;
                    case 11:
                        #region
                        list[42].month1 = _list_rgsr42.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[42].month2 = _list_rgsr42.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[42].month3 = _list_rgsr42.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[42].month4 = _list_rgsr42.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[42].month5 = _list_rgsr42.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[42].month6 = _list_rgsr42.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[42].month7 = _list_rgsr42.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[42].month8 = _list_rgsr42.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[42].month9 = _list_rgsr42.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[42].month10 = _list_rgsr42.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[42].month11 = _list_rgsr42.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[42].month12 = _list_rgsr42.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[43].month1 = _list_rgsr43.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[43].month2 = _list_rgsr43.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[43].month3 = _list_rgsr43.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[43].month4 = _list_rgsr43.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[43].month5 = _list_rgsr43.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[43].month6 = _list_rgsr43.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[43].month7 = _list_rgsr43.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[43].month8 = _list_rgsr43.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[43].month9 = _list_rgsr43.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[43].month10 = _list_rgsr43.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[43].month11 = _list_rgsr43.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[43].month12 = _list_rgsr43.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[44].month1 = _list_rgsr44.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[44].month2 = _list_rgsr44.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[44].month3 = _list_rgsr44.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[44].month4 = _list_rgsr44.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[44].month5 = _list_rgsr44.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[44].month6 = _list_rgsr44.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[44].month7 = _list_rgsr44.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[44].month8 = _list_rgsr44.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[44].month9 = _list_rgsr44.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[44].month10 = _list_rgsr44.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[44].month11 = _list_rgsr44.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[44].month12 = _list_rgsr44.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[45].month1 = _list_rgsr45.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[45].month2 = _list_rgsr45.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[45].month3 = _list_rgsr45.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[45].month4 = _list_rgsr45.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[45].month5 = _list_rgsr45.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[45].month6 = _list_rgsr45.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[45].month7 = _list_rgsr45.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[45].month8 = _list_rgsr45.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[45].month9 = _list_rgsr45.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[45].month10 = _list_rgsr45.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[45].month11 = _list_rgsr45.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[45].month12 = _list_rgsr45.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[46].month1 = _list_rgsr46.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[46].month2 = _list_rgsr46.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[46].month3 = _list_rgsr46.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[46].month4 = _list_rgsr46.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[46].month5 = _list_rgsr46.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[46].month6 = _list_rgsr46.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[46].month7 = _list_rgsr46.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[46].month8 = _list_rgsr46.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[46].month9 = _list_rgsr46.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[46].month10 = _list_rgsr46.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[46].month11 = _list_rgsr46.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[46].month12 = _list_rgsr46.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[47].month1 = _list_rgsr47.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[47].month2 = _list_rgsr47.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[47].month3 = _list_rgsr47.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[47].month4 = _list_rgsr47.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[47].month5 = _list_rgsr47.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[47].month6 = _list_rgsr47.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[47].month7 = _list_rgsr47.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[47].month8 = _list_rgsr47.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[47].month9 = _list_rgsr47.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[47].month10 = _list_rgsr47.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[47].month11 = _list_rgsr47.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[47].month12 = _list_rgsr47.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();


                        list[49].month1 = _list_rgsj49.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[49].month2 = _list_rgsj49.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[49].month3 = _list_rgsj49.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[49].month4 = _list_rgsj49.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[49].month5 = _list_rgsj49.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[49].month6 = _list_rgsj49.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[49].month7 = _list_rgsj49.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[49].month8 = _list_rgsj49.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[49].month9 = _list_rgsj49.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[49].month10 = _list_rgsj49.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[49].month11 = _list_rgzc49.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[49].month12 = _list_rgzc49.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[50].month1 = _list_rgsj50.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[50].month2 = _list_rgsj50.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[50].month3 = _list_rgsj50.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[50].month4 = _list_rgsj50.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[50].month5 = _list_rgsj50.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[50].month6 = _list_rgsj50.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[50].month7 = _list_rgsj50.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[50].month8 = _list_rgsj50.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[50].month9 = _list_rgsj50.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[50].month10 = _list_rgsj50.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[50].month11 = _list_rgzc50.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[50].month12 = _list_rgzc50.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[51].month1 = _list_rgsj51.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[51].month2 = _list_rgsj51.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[51].month3 = _list_rgsj51.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[51].month4 = _list_rgsj51.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[51].month5 = _list_rgsj51.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[51].month6 = _list_rgsj51.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[51].month7 = _list_rgsj51.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[51].month8 = _list_rgsj51.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[51].month9 = _list_rgsj51.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[51].month10 = _list_rgsj51.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[51].month11 = _list_rgzc51.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[51].month12 = _list_rgzc51.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[52].month1 = _list_rgsj52.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[52].month2 = _list_rgsj52.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[52].month3 = _list_rgsj52.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[52].month4 = _list_rgsj52.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[52].month5 = _list_rgsj52.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[52].month6 = _list_rgsj52.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[52].month7 = _list_rgsj52.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[52].month8 = _list_rgsj52.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[52].month9 = _list_rgsj52.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[52].month10 = _list_rgsj52.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[52].month11 = _list_rgzc52.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[52].month12 = _list_rgzc52.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[53].month1 = _list_rgsj53.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[53].month2 = _list_rgsj53.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[53].month3 = _list_rgsj53.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[53].month4 = _list_rgsj53.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[53].month5 = _list_rgsj53.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[53].month6 = _list_rgsj53.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[53].month7 = _list_rgsj53.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[53].month8 = _list_rgsj53.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[53].month9 = _list_rgsj53.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[53].month10 = _list_rgsj53.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[53].month11 = _list_rgzc53.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[53].month12 = _list_rgzc53.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[54].month1 = _list_rgsj54.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[54].month2 = _list_rgsj54.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[54].month3 = _list_rgsj54.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[54].month4 = _list_rgsj54.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[54].month5 = _list_rgsj54.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[54].month6 = _list_rgsj54.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[54].month7 = _list_rgsj54.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[54].month8 = _list_rgsj54.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[54].month9 = _list_rgsj54.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[54].month10 = _list_rgsj54.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[54].month11 = _list_rgzc54.Where(p => p.monthly == 11).Sum(p => p.proBudget).ToString();
                        list[54].month12 = _list_rgzc54.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();
                        #endregion
                        break;
                    case 12:
                        #region
                        list[42].month1 = _list_rgsr42.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[42].month2 = _list_rgsr42.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[42].month3 = _list_rgsr42.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[42].month4 = _list_rgsr42.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[42].month5 = _list_rgsr42.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[42].month6 = _list_rgsr42.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[42].month7 = _list_rgsr42.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[42].month8 = _list_rgsr42.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[42].month9 = _list_rgsr42.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[42].month10 = _list_rgsr42.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[42].month11 = _list_rgsr42.Where(p => p.monthly == 11).Sum(p => p.quotaLabor).ToString();
                        list[42].month12 = _list_rgsr42.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[43].month1 = _list_rgsr43.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[43].month2 = _list_rgsr43.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[43].month3 = _list_rgsr43.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[43].month4 = _list_rgsr43.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[43].month5 = _list_rgsr43.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[43].month6 = _list_rgsr43.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[43].month7 = _list_rgsr43.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[43].month8 = _list_rgsr43.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[43].month9 = _list_rgsr43.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[43].month10 = _list_rgsr43.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[43].month11 = _list_rgsr43.Where(p => p.monthly == 11).Sum(p => p.quotaLabor).ToString();
                        list[43].month12 = _list_rgsr43.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[44].month1 = _list_rgsr44.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[44].month2 = _list_rgsr44.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[44].month3 = _list_rgsr44.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[44].month4 = _list_rgsr44.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[44].month5 = _list_rgsr44.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[44].month6 = _list_rgsr44.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[44].month7 = _list_rgsr44.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[44].month8 = _list_rgsr44.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[44].month9 = _list_rgsr44.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[44].month10 = _list_rgsr44.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[44].month11 = _list_rgsr44.Where(p => p.monthly == 11).Sum(p => p.quotaLabor).ToString();
                        list[44].month12 = _list_rgsr44.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[45].month1 = _list_rgsr45.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[45].month2 = _list_rgsr45.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[45].month3 = _list_rgsr45.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[45].month4 = _list_rgsr45.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[45].month5 = _list_rgsr45.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[45].month6 = _list_rgsr45.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[45].month7 = _list_rgsr45.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[45].month8 = _list_rgsr45.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[45].month9 = _list_rgsr45.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[45].month10 = _list_rgsr45.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[45].month11 = _list_rgsr45.Where(p => p.monthly == 11).Sum(p => p.quotaLabor).ToString();
                        list[45].month12 = _list_rgsr45.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[46].month1 = _list_rgsr46.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[46].month2 = _list_rgsr46.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[46].month3 = _list_rgsr46.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[46].month4 = _list_rgsr46.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[46].month5 = _list_rgsr46.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[46].month6 = _list_rgsr46.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[46].month7 = _list_rgsr46.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[46].month8 = _list_rgsr46.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[46].month9 = _list_rgsr46.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[46].month10 = _list_rgsr46.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[46].month11 = _list_rgsr46.Where(p => p.monthly == 11).Sum(p => p.quotaLabor).ToString();
                        list[46].month12 = _list_rgsr46.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[47].month1 = _list_rgsr47.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[47].month2 = _list_rgsr47.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[47].month3 = _list_rgsr47.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[47].month4 = _list_rgsr47.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[47].month5 = _list_rgsr47.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[47].month6 = _list_rgsr47.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[47].month7 = _list_rgsr47.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[47].month8 = _list_rgsr47.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[47].month9 = _list_rgsr47.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[47].month10 = _list_rgsr47.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[47].month11 = _list_rgsr47.Where(p => p.monthly == 11).Sum(p => p.quotaLabor).ToString();
                        list[47].month12 = _list_rgsr47.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();


                        list[49].month1 = _list_rgsj49.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[49].month2 = _list_rgsj49.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[49].month3 = _list_rgsj49.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[49].month4 = _list_rgsj49.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[49].month5 = _list_rgsj49.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[49].month6 = _list_rgsj49.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[49].month7 = _list_rgsj49.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[49].month8 = _list_rgsj49.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[49].month9 = _list_rgsj49.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[49].month10 = _list_rgsj49.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[49].month11 = _list_rgsj49.Where(p => p.monthly == 11).Sum(p => p.quotaLabor).ToString();
                        list[49].month12 = _list_rgzc49.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[50].month1 = _list_rgsj50.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[50].month2 = _list_rgsj50.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[50].month3 = _list_rgsj50.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[50].month4 = _list_rgsj50.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[50].month5 = _list_rgsj50.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[50].month6 = _list_rgsj50.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[50].month7 = _list_rgsj50.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[50].month8 = _list_rgsj50.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[50].month9 = _list_rgsj50.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[50].month10 = _list_rgsj50.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[50].month11 = _list_rgsj50.Where(p => p.monthly == 11).Sum(p => p.quotaLabor).ToString();
                        list[50].month12 = _list_rgzc50.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[51].month1 = _list_rgsj51.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[51].month2 = _list_rgsj51.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[51].month3 = _list_rgsj51.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[51].month4 = _list_rgsj51.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[51].month5 = _list_rgsj51.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[51].month6 = _list_rgsj51.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[51].month7 = _list_rgsj51.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[51].month8 = _list_rgsj51.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[51].month9 = _list_rgsj51.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[51].month10 = _list_rgsj51.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[51].month11 = _list_rgsj51.Where(p => p.monthly == 11).Sum(p => p.quotaLabor).ToString();
                        list[51].month12 = _list_rgzc51.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[52].month1 = _list_rgsj52.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[52].month2 = _list_rgsj52.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[52].month3 = _list_rgsj52.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[52].month4 = _list_rgsj52.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[52].month5 = _list_rgsj52.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[52].month6 = _list_rgsj52.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[52].month7 = _list_rgsj52.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[52].month8 = _list_rgsj52.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[52].month9 = _list_rgsj52.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[52].month10 = _list_rgsj52.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[52].month11 = _list_rgsj52.Where(p => p.monthly == 11).Sum(p => p.quotaLabor).ToString();
                        list[52].month12 = _list_rgzc52.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[53].month1 = _list_rgsj53.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[53].month2 = _list_rgsj53.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[53].month3 = _list_rgsj53.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[53].month4 = _list_rgsj53.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[53].month5 = _list_rgsj53.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[53].month6 = _list_rgsj53.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[53].month7 = _list_rgsj53.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[53].month8 = _list_rgsj53.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[53].month9 = _list_rgsj53.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[53].month10 = _list_rgsj53.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[53].month11 = _list_rgsj53.Where(p => p.monthly == 11).Sum(p => p.quotaLabor).ToString();
                        list[53].month12 = _list_rgzc53.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();

                        list[54].month1 = _list_rgsj54.Where(p => p.monthly == 1).Sum(p => p.quotaLabor).ToString();
                        list[54].month2 = _list_rgsj54.Where(p => p.monthly == 2).Sum(p => p.quotaLabor).ToString();
                        list[54].month3 = _list_rgsj54.Where(p => p.monthly == 3).Sum(p => p.quotaLabor).ToString();
                        list[54].month4 = _list_rgsj54.Where(p => p.monthly == 4).Sum(p => p.quotaLabor).ToString();
                        list[54].month5 = _list_rgsj54.Where(p => p.monthly == 5).Sum(p => p.quotaLabor).ToString();
                        list[54].month6 = _list_rgsj54.Where(p => p.monthly == 6).Sum(p => p.quotaLabor).ToString();
                        list[54].month7 = _list_rgsj54.Where(p => p.monthly == 7).Sum(p => p.quotaLabor).ToString();
                        list[54].month8 = _list_rgsj54.Where(p => p.monthly == 8).Sum(p => p.quotaLabor).ToString();
                        list[54].month9 = _list_rgsj54.Where(p => p.monthly == 9).Sum(p => p.quotaLabor).ToString();
                        list[54].month10 = _list_rgsj54.Where(p => p.monthly == 10).Sum(p => p.quotaLabor).ToString();
                        list[54].month11 = _list_rgsj54.Where(p => p.monthly == 11).Sum(p => p.quotaLabor).ToString();
                        list[54].month12 = _list_rgzc54.Where(p => p.monthly == 12).Sum(p => p.proBudget).ToString();
                        #endregion
                        break;                        
                }
                #endregion

                #endregion
                
                #region//序号43, 50
                /*
                  * Cyber-人工成本 (万元)
                  *  = 44 + 45 + 46 + 47
                  */
                list[41].classify = "Cyber-人工成本 (万元)";
                list[41].yearLine = "1";
                list[41].goal = (double.Parse(list[42].goal) + double.Parse(list[43].goal) + double.Parse(list[44].goal) + double.Parse(list[45].goal)).ToString();
                list[41].month1 = (double.Parse(list[42].month1) + double.Parse(list[43].month1) + double.Parse(list[44].month1) + double.Parse(list[45].month1)).ToString();
                list[41].month2 = (double.Parse(list[42].month2) + double.Parse(list[43].month2) + double.Parse(list[44].month2) + double.Parse(list[45].month2)).ToString();
                list[41].month3 = (double.Parse(list[42].month3) + double.Parse(list[43].month3) + double.Parse(list[44].month3) + double.Parse(list[45].month3)).ToString();
                list[41].month4 = (double.Parse(list[42].month4) + double.Parse(list[43].month4) + double.Parse(list[44].month4) + double.Parse(list[45].month4)).ToString();
                list[41].month5 = (double.Parse(list[42].month5) + double.Parse(list[43].month5) + double.Parse(list[44].month5) + double.Parse(list[45].month5)).ToString();
                list[41].month6 = (double.Parse(list[42].month6) + double.Parse(list[43].month6) + double.Parse(list[44].month6) + double.Parse(list[45].month6)).ToString();
                list[41].month7 = (double.Parse(list[42].month7) + double.Parse(list[43].month7) + double.Parse(list[44].month7) + double.Parse(list[45].month7)).ToString();
                list[41].month8 = (double.Parse(list[42].month8) + double.Parse(list[43].month8) + double.Parse(list[44].month8) + double.Parse(list[45].month8)).ToString();
                list[41].month9 = (double.Parse(list[42].month9) + double.Parse(list[43].month9) + double.Parse(list[44].month9) + double.Parse(list[45].month9)).ToString();
                list[41].month10 = (double.Parse(list[42].month10) + double.Parse(list[43].month10) + double.Parse(list[44].month10) + double.Parse(list[45].month10)).ToString();
                list[41].month11 = (double.Parse(list[42].month11) + double.Parse(list[43].month11) + double.Parse(list[44].month11) + double.Parse(list[45].month11)).ToString();
                list[41].month12 = (double.Parse(list[42].month12) + double.Parse(list[43].month12) + double.Parse(list[44].month12) + double.Parse(list[45].month12)).ToString();
                list[41].sj = (double.Parse(list[42].sj) + double.Parse(list[43].sj) + double.Parse(list[44].sj) + double.Parse(list[45].sj)).ToString();
                list[41].yj = (double.Parse(list[42].yj) + double.Parse(list[43].yj) + double.Parse(list[44].yj) + double.Parse(list[45].yj)).ToString();
                list[41].ys = (double.Parse(list[42].ys) + double.Parse(list[43].ys) + double.Parse(list[44].ys) + double.Parse(list[45].ys)).ToString();

                //序号50
                /*
                  * Physical-人工成本 (万元)
                  *  = 51 + 52 + 53 + 54
                  */
                list[48].classify = "Physical-人工成本 (万元)";
                list[48].yearLine = "1";
                list[48].goal = (double.Parse(list[49].goal) + double.Parse(list[50].goal) + double.Parse(list[51].goal) + double.Parse(list[52].goal)).ToString();
                list[48].month1 = (double.Parse(list[49].month1) + double.Parse(list[50].month1) + double.Parse(list[51].month1) + double.Parse(list[52].month1)).ToString();
                list[48].month2 = (double.Parse(list[49].month2) + double.Parse(list[50].month2) + double.Parse(list[51].month2) + double.Parse(list[52].month2)).ToString();
                list[48].month3 = (double.Parse(list[49].month3) + double.Parse(list[50].month3) + double.Parse(list[51].month3) + double.Parse(list[52].month3)).ToString();
                list[48].month4 = (double.Parse(list[49].month4) + double.Parse(list[50].month4) + double.Parse(list[51].month4) + double.Parse(list[52].month4)).ToString();
                list[48].month5 = (double.Parse(list[49].month5) + double.Parse(list[50].month5) + double.Parse(list[51].month5) + double.Parse(list[52].month5)).ToString();
                list[48].month6 = (double.Parse(list[49].month6) + double.Parse(list[50].month6) + double.Parse(list[51].month6) + double.Parse(list[52].month6)).ToString();
                list[48].month7 = (double.Parse(list[49].month7) + double.Parse(list[50].month7) + double.Parse(list[51].month7) + double.Parse(list[52].month7)).ToString();
                list[48].month8 = (double.Parse(list[49].month8) + double.Parse(list[50].month8) + double.Parse(list[51].month8) + double.Parse(list[52].month8)).ToString();
                list[48].month9 = (double.Parse(list[49].month9) + double.Parse(list[50].month9) + double.Parse(list[51].month9) + double.Parse(list[52].month9)).ToString();
                list[48].month10 = (double.Parse(list[49].month10) + double.Parse(list[50].month10) + double.Parse(list[51].month10) + double.Parse(list[52].month10)).ToString();
                list[48].month11 = (double.Parse(list[49].month11) + double.Parse(list[50].month11) + double.Parse(list[51].month11) + double.Parse(list[52].month11)).ToString();
                list[48].month12 = (double.Parse(list[49].month12) + double.Parse(list[50].month12) + double.Parse(list[51].month12) + double.Parse(list[52].month12)).ToString();
                list[48].sj = (double.Parse(list[49].sj) + double.Parse(list[50].sj) + double.Parse(list[51].sj) + double.Parse(list[52].sj)).ToString();
                list[48].yj = (double.Parse(list[49].yj) + double.Parse(list[50].yj) + double.Parse(list[51].yj) + double.Parse(list[52].yj)).ToString();
                list[48].ys = (double.Parse(list[49].ys) + double.Parse(list[50].ys) + double.Parse(list[51].ys) + double.Parse(list[52].ys)).ToString();

                #endregion

                #region//序号57—63
                //序号57
                //43行 减去 50行
                list[55].classify = "盈亏（万元）";
                list[55].yearLine = (double.Parse(list[41].yearLine) - double.Parse(list[48].yearLine)).ToString();
                list[55].goal = (double.Parse(list[41].goal) - double.Parse(list[48].goal)).ToString();
                list[55].month1 = (double.Parse(list[41].month1) - double.Parse(list[48].month1)).ToString();
                list[55].month2 = (double.Parse(list[41].month2) - double.Parse(list[48].month2)).ToString();
                list[55].month3 = (double.Parse(list[41].month3) - double.Parse(list[48].month3)).ToString();
                list[55].month4 = (double.Parse(list[41].month4) - double.Parse(list[48].month4)).ToString();
                list[55].month5 = (double.Parse(list[41].month5) - double.Parse(list[48].month5)).ToString();
                list[55].month6 = (double.Parse(list[41].month6) - double.Parse(list[48].month6)).ToString();
                list[55].month7 = (double.Parse(list[41].month7) - double.Parse(list[48].month7)).ToString();
                list[55].month8 = (double.Parse(list[41].month8) - double.Parse(list[48].month8)).ToString();
                list[55].month9 = (double.Parse(list[41].month9) - double.Parse(list[48].month9)).ToString();
                list[55].month10 = (double.Parse(list[41].month10) - double.Parse(list[48].month10)).ToString();
                list[55].month11 = (double.Parse(list[41].month11) - double.Parse(list[48].month11)).ToString();
                list[55].month12 = (double.Parse(list[41].month12) - double.Parse(list[48].month12)).ToString();
                list[55].sj = (double.Parse(list[41].sj) - double.Parse(list[48].sj)).ToString();
                list[55].yj = (double.Parse(list[41].yj) - double.Parse(list[48].yj)).ToString();
                list[55].ys = (double.Parse(list[41].ys) - double.Parse(list[48].ys)).ToString();

                //序号58
                //44行 减去 51行
                list[56].classify = "其中：市场人工";
                list[56].yearLine = (double.Parse(list[42].yearLine) - double.Parse(list[49].yearLine)).ToString();
                list[56].goal = (double.Parse(list[42].goal) - double.Parse(list[49].goal)).ToString();
                list[56].month1 = (double.Parse(list[42].month1) - double.Parse(list[49].month1)).ToString();
                list[56].month2 = (double.Parse(list[42].month2) - double.Parse(list[49].month2)).ToString();
                list[56].month3 = (double.Parse(list[42].month3) - double.Parse(list[49].month3)).ToString();
                list[56].month4 = (double.Parse(list[42].month4) - double.Parse(list[49].month4)).ToString();
                list[56].month5 = (double.Parse(list[42].month5) - double.Parse(list[49].month5)).ToString();
                list[56].month6 = (double.Parse(list[42].month6) - double.Parse(list[49].month6)).ToString();
                list[56].month7 = (double.Parse(list[42].month7) - double.Parse(list[49].month7)).ToString();
                list[56].month8 = (double.Parse(list[42].month8) - double.Parse(list[49].month8)).ToString();
                list[56].month9 = (double.Parse(list[42].month9) - double.Parse(list[49].month9)).ToString();
                list[56].month10 = (double.Parse(list[42].month10) - double.Parse(list[49].month10)).ToString();
                list[56].month11 = (double.Parse(list[42].month11) - double.Parse(list[49].month11)).ToString();
                list[56].month12 = (double.Parse(list[42].month12) - double.Parse(list[49].month12)).ToString();
                list[56].sj = (double.Parse(list[42].sj) - double.Parse(list[49].sj)).ToString();
                list[56].yj = (double.Parse(list[42].yj) - double.Parse(list[49].yj)).ToString();
                list[56].ys = (double.Parse(list[42].ys) - double.Parse(list[49].ys)).ToString();

                //序号59
                //45行 减去 52行
                list[57].classify = "管理人工";
                list[57].yearLine = (double.Parse(list[43].yearLine) - double.Parse(list[50].yearLine)).ToString();
                list[57].goal = (double.Parse(list[43].goal) - double.Parse(list[50].goal)).ToString();
                list[57].month1 = (double.Parse(list[43].month1) - double.Parse(list[50].month1)).ToString();
                list[57].month2 = (double.Parse(list[43].month2) - double.Parse(list[50].month2)).ToString();
                list[57].month3 = (double.Parse(list[43].month3) - double.Parse(list[50].month3)).ToString();
                list[57].month4 = (double.Parse(list[43].month4) - double.Parse(list[50].month4)).ToString();
                list[57].month5 = (double.Parse(list[43].month5) - double.Parse(list[50].month5)).ToString();
                list[57].month6 = (double.Parse(list[43].month6) - double.Parse(list[50].month6)).ToString();
                list[57].month7 = (double.Parse(list[43].month7) - double.Parse(list[50].month7)).ToString();
                list[57].month8 = (double.Parse(list[43].month8) - double.Parse(list[50].month8)).ToString();
                list[57].month9 = (double.Parse(list[43].month9) - double.Parse(list[50].month9)).ToString();
                list[57].month10 = (double.Parse(list[43].month10) - double.Parse(list[50].month10)).ToString();
                list[57].month11 = (double.Parse(list[43].month11) - double.Parse(list[50].month11)).ToString();
                list[57].month12 = (double.Parse(list[43].month12) - double.Parse(list[50].month12)).ToString();
                list[57].sj = (double.Parse(list[43].sj) - double.Parse(list[50].sj)).ToString();
                list[57].yj = (double.Parse(list[43].yj) - double.Parse(list[50].yj)).ToString();
                list[57].ys = (double.Parse(list[43].ys) - double.Parse(list[50].ys)).ToString();

                //序号60
                //46行 减去 53行
                list[58].classify = "制造人工";
                list[58].yearLine = (double.Parse(list[44].yearLine) - double.Parse(list[51].yearLine)).ToString();
                list[58].goal = (double.Parse(list[44].goal) - double.Parse(list[51].goal)).ToString();
                list[58].month1 = (double.Parse(list[44].month1) - double.Parse(list[51].month1)).ToString();
                list[58].month2 = (double.Parse(list[44].month2) - double.Parse(list[51].month2)).ToString();
                list[58].month3 = (double.Parse(list[44].month3) - double.Parse(list[51].month3)).ToString();
                list[58].month4 = (double.Parse(list[44].month4) - double.Parse(list[51].month4)).ToString();
                list[58].month5 = (double.Parse(list[44].month5) - double.Parse(list[51].month5)).ToString();
                list[58].month6 = (double.Parse(list[44].month6) - double.Parse(list[51].month6)).ToString();
                list[58].month7 = (double.Parse(list[44].month7) - double.Parse(list[51].month7)).ToString();
                list[58].month8 = (double.Parse(list[44].month8) - double.Parse(list[51].month8)).ToString();
                list[58].month9 = (double.Parse(list[44].month9) - double.Parse(list[51].month9)).ToString();
                list[58].month10 = (double.Parse(list[44].month10) - double.Parse(list[51].month10)).ToString();
                list[58].month11 = (double.Parse(list[44].month11) - double.Parse(list[51].month11)).ToString();
                list[58].month12 = (double.Parse(list[44].month12) - double.Parse(list[51].month12)).ToString();
                list[58].sj = (double.Parse(list[44].sj) - double.Parse(list[51].sj)).ToString();
                list[58].yj = (double.Parse(list[44].yj) - double.Parse(list[51].yj)).ToString();
                list[58].ys = (double.Parse(list[44].ys) - double.Parse(list[51].ys)).ToString();

                //序号61
                //47行 减去 54行
                list[59].classify = "直接人工";
                list[59].yearLine = (double.Parse(list[45].yearLine) - double.Parse(list[52].yearLine)).ToString();
                list[59].goal = (double.Parse(list[45].goal) - double.Parse(list[52].goal)).ToString();
                list[59].month1 = (double.Parse(list[45].month1) - double.Parse(list[52].month1)).ToString();
                list[59].month2 = (double.Parse(list[45].month2) - double.Parse(list[52].month2)).ToString();
                list[59].month3 = (double.Parse(list[45].month3) - double.Parse(list[52].month3)).ToString();
                list[59].month4 = (double.Parse(list[45].month4) - double.Parse(list[52].month4)).ToString();
                list[59].month5 = (double.Parse(list[45].month5) - double.Parse(list[52].month5)).ToString();
                list[59].month6 = (double.Parse(list[45].month6) - double.Parse(list[52].month6)).ToString();
                list[59].month7 = (double.Parse(list[45].month7) - double.Parse(list[52].month7)).ToString();
                list[59].month8 = (double.Parse(list[45].month8) - double.Parse(list[52].month8)).ToString();
                list[59].month9 = (double.Parse(list[45].month9) - double.Parse(list[52].month9)).ToString();
                list[59].month10 = (double.Parse(list[45].month10) - double.Parse(list[52].month10)).ToString();
                list[59].month11 = (double.Parse(list[45].month11) - double.Parse(list[52].month11)).ToString();
                list[59].month12 = (double.Parse(list[45].month12) - double.Parse(list[52].month12)).ToString();
                list[59].sj = (double.Parse(list[45].sj) - double.Parse(list[52].sj)).ToString();
                list[59].yj = (double.Parse(list[45].yj) - double.Parse(list[52].yj)).ToString();
                list[59].ys = (double.Parse(list[45].ys) - double.Parse(list[52].ys)).ToString();

                //序号62
                //48行 减去 55行
                list[60].classify = "直接人工BMI";
                list[60].yearLine = (double.Parse(list[46].yearLine) - double.Parse(list[53].yearLine)).ToString();
                list[60].goal = (double.Parse(list[46].goal) - double.Parse(list[53].goal)).ToString();
                list[60].month1 = (double.Parse(list[46].month1) - double.Parse(list[53].month1)).ToString();
                list[60].month2 = (double.Parse(list[46].month2) - double.Parse(list[53].month2)).ToString();
                list[60].month3 = (double.Parse(list[46].month3) - double.Parse(list[53].month3)).ToString();
                list[60].month4 = (double.Parse(list[46].month4) - double.Parse(list[53].month4)).ToString();
                list[60].month5 = (double.Parse(list[46].month5) - double.Parse(list[53].month5)).ToString();
                list[60].month6 = (double.Parse(list[46].month6) - double.Parse(list[53].month6)).ToString();
                list[60].month7 = (double.Parse(list[46].month7) - double.Parse(list[53].month7)).ToString();
                list[60].month8 = (double.Parse(list[46].month8) - double.Parse(list[53].month8)).ToString();
                list[60].month9 = (double.Parse(list[46].month9) - double.Parse(list[53].month9)).ToString();
                list[60].month10 = (double.Parse(list[46].month10) - double.Parse(list[53].month10)).ToString();
                list[60].month11 = (double.Parse(list[46].month11) - double.Parse(list[53].month11)).ToString();
                list[60].month12 = (double.Parse(list[46].month12) - double.Parse(list[53].month12)).ToString();
                list[60].sj = (double.Parse(list[46].sj) - double.Parse(list[53].sj)).ToString();
                list[60].yj = (double.Parse(list[46].yj) - double.Parse(list[53].yj)).ToString();
                list[60].ys = (double.Parse(list[46].ys) - double.Parse(list[53].ys)).ToString();

                //序号63
                //49行 减去 56行
                list[61].classify = "直接人工BPL";
                list[61].yearLine = (double.Parse(list[47].yearLine) - double.Parse(list[54].yearLine)).ToString();
                list[61].goal = (double.Parse(list[47].goal) - double.Parse(list[54].goal)).ToString();
                list[61].month1 = (double.Parse(list[47].month1) - double.Parse(list[54].month1)).ToString();
                list[61].month2 = (double.Parse(list[47].month2) - double.Parse(list[54].month2)).ToString();
                list[61].month3 = (double.Parse(list[47].month3) - double.Parse(list[54].month3)).ToString();
                list[61].month4 = (double.Parse(list[47].month4) - double.Parse(list[54].month4)).ToString();
                list[61].month5 = (double.Parse(list[47].month5) - double.Parse(list[54].month5)).ToString();
                list[61].month6 = (double.Parse(list[47].month6) - double.Parse(list[54].month6)).ToString();
                list[61].month7 = (double.Parse(list[47].month7) - double.Parse(list[54].month7)).ToString();
                list[61].month8 = (double.Parse(list[47].month8) - double.Parse(list[54].month8)).ToString();
                list[61].month9 = (double.Parse(list[47].month9) - double.Parse(list[54].month9)).ToString();
                list[61].month10 = (double.Parse(list[47].month10) - double.Parse(list[54].month10)).ToString();
                list[61].month11 = (double.Parse(list[47].month11) - double.Parse(list[54].month11)).ToString();
                list[61].month12 = (double.Parse(list[47].month12) - double.Parse(list[54].month12)).ToString();
                list[61].sj = (double.Parse(list[47].sj) - double.Parse(list[54].sj)).ToString();
                list[61].yj = (double.Parse(list[47].yj) - double.Parse(list[54].yj)).ToString();
                list[61].ys = (double.Parse(list[47].ys) - double.Parse(list[54].ys)).ToString();

                #endregion

                #region//序号7—12
                //序号7
                //营收 除以 Cyber-总人数 序号3 除以 序号13
                //月份： 累计营收（和） 除以 Cyber-总人数月度人数（平均值）
                list[5].classify = "Cyber/人均产值(万元/人)";
                list[5].yearLine = (double.Parse(list[1].yearLine) / double.Parse(list[11].yearLine)).ToString();
                list[5].goal = (double.Parse(list[1].goal) / double.Parse(list[11].goal)).ToString();
                list[5].sj = (double.Parse(list[1].sj) / double.Parse(list[11].sj)).ToString();
                list[5].yj = (double.Parse(list[1].yj) / double.Parse(list[11].yj)).ToString();
                list[5].ys = (double.Parse(list[1].ys) / double.Parse(list[11].ys)).ToString();

                //序号8
                //营收 除以 Physical-总人数 序号3 除以 序号23
                //月份： 累计营收（和） 除以 Physical-总人数月度人数（平均值）
                list[6].classify = "Physical-人均产值(万元/人)";
                list[6].yearLine = (double.Parse(list[1].yearLine) / double.Parse(list[21].yearLine)).ToString();
                list[6].goal = (double.Parse(list[1].goal) / double.Parse(list[21].goal)).ToString();
                list[6].sj = (double.Parse(list[1].sj) / double.Parse(list[21].sj)).ToString();

                list[6].yj = (double.Parse(list[1].yj) / double.Parse(list[21].yj)).ToString();
                list[6].ys = (double.Parse(list[1].ys) / double.Parse(list[21].ys)).ToString();

                //序号9
                //营收 除以Cyber-人工成本 序号3 除以 序号43
                //月份： 累计营收（和） 除以 Cyber-人工成本月度之和
                list[7].classify = "Cyber-人工成本占营收比";
                list[7].yearLine = string.Format("{0:P}", (double.Parse(list[1].yearLine) / double.Parse(list[41].yearLine)));
                list[7].goal = string.Format("{0:P}", (double.Parse(list[1].goal) / double.Parse(list[41].goal)));
                list[7].sj = string.Format("{0:P}", (double.Parse(list[1].sj) / double.Parse(list[41].sj)));
                list[7].yj = string.Format("{0:P}", (double.Parse(list[1].yj) / double.Parse(list[41].yj)));
                list[7].ys = string.Format("{0:P}", (double.Parse(list[1].ys) / double.Parse(list[41].ys)));


                //序号10
                //营收 除以 Physical-人工成本 序号3 除以 序号50
                //月份： 累计营收（和） 除以 Physical-人工成本月度之和
                list[8].classify = "Physical-人工成本占营收比";
                list[8].yearLine = string.Format("{0:P}", (double.Parse(list[1].yearLine) / double.Parse(list[48].yearLine)));
                list[8].goal = string.Format("{0:P}", (double.Parse(list[1].goal) / double.Parse(list[48].goal)));
                list[8].sj = string.Format("{0:P}", (double.Parse(list[1].sj) / double.Parse(list[48].sj)));
                list[8].yj = string.Format("{0:P}", (double.Parse(list[1].yj) / double.Parse(list[48].yj)));
                list[8].ys = string.Format("{0:P}", (double.Parse(list[1].ys) / double.Parse(list[48].ys)));


                //序号11
                //利润 除以 Cyber-人工成本 序号4 除以 序号43
                //月份： 累计利润（和） 除以 Cyber-人工成本月度之和
                list[9].classify = "Cyber-劳动效率";
                list[9].yearLine = (double.Parse(list[2].yearLine) / double.Parse(list[41].yearLine)).ToString();
                list[9].goal = (double.Parse(list[2].goal) / double.Parse(list[41].goal)).ToString();
                list[9].sj = (double.Parse(list[2].sj) / double.Parse(list[41].sj)).ToString();
                list[9].yj = (double.Parse(list[2].yj) / double.Parse(list[41].yj)).ToString();
                list[9].ys = (double.Parse(list[2].ys) / double.Parse(list[41].ys)).ToString();


                //序号12
                //利润 除以 Physical-人工成本 序号4 除以 序号50
                //月份： 累计利润（和） 除以 Physical-人工成本月度之和
                list[10].classify = "Physical-劳动效率";
                list[10].yearLine = (double.Parse(list[2].yearLine) / double.Parse(list[48].yearLine)).ToString();
                list[10].goal = (double.Parse(list[2].goal) / double.Parse(list[48].goal)).ToString();
                list[10].sj = (double.Parse(list[2].sj) / double.Parse(list[48].sj)).ToString();
                list[10].yj = (double.Parse(list[2].yj) / double.Parse(list[48].yj)).ToString();
                list[10].ys = (double.Parse(list[2].ys) / double.Parse(list[48].ys)).ToString();

                //月份
                double ysSum = double.Parse(list[1].month1);//营收累计
                double lrSum = double.Parse(list[2].month1);//利润累计
                double cyNum = double.Parse(list[11].month1);//Cyber-总人数 累计
                double phNum = double.Parse(list[21].month1);//Physical-总人数 累计
                double cySum = double.Parse(list[41].month1);//Cyber-人工成本
                double phSum = double.Parse(list[48].month1);//Physical-人工成本

                list[5].month1 = string.Format("{0:N2}", (ysSum / cyNum));
                list[6].month1 = string.Format("{0:N2}", (ysSum / phNum));
                list[7].month1 = string.Format("{0:P}", (ysSum / cySum));
                list[8].month1 = string.Format("{0:P}", (ysSum / phSum));
                list[9].month1 = string.Format("{0:N2}", (lrSum / cySum));
                list[10].month1 = string.Format("{0:N2}", (lrSum / phSum));

                ysSum += double.Parse(list[1].month2);//营收累计
                lrSum += double.Parse(list[2].month2);//利润累计
                cyNum += double.Parse(list[11].month2);//Cyber-总人数 累计
                phNum += double.Parse(list[21].month2);//Physical-总人数 累计
                cySum += double.Parse(list[41].month2);//Cyber-人工成本
                phSum += double.Parse(list[48].month2);//Physical-人工成本
                list[5].month2 = string.Format("{0:N2}", (ysSum / (cyNum / 2)));
                list[6].month2 = string.Format("{0:N2}", (ysSum / (phNum / 2)));
                list[7].month2 = string.Format("{0:P}", (ysSum / cySum));
                list[8].month2 = string.Format("{0:P}", (ysSum / phSum));
                list[9].month2 = string.Format("{0:N2}", (lrSum / cySum));
                list[10].month2 = string.Format("{0:N2}", (lrSum / phSum));

                ysSum += double.Parse(list[1].month3);//营收累计
                lrSum += double.Parse(list[2].month3);//利润累计
                cyNum += double.Parse(list[11].month3);//Cyber-总人数 累计
                phNum += double.Parse(list[21].month3);//Physical-总人数 累计
                cySum += double.Parse(list[41].month3);//Cyber-人工成本
                phSum += double.Parse(list[48].month3);//Physical-人工成本
                list[5].month3 = string.Format("{0:N2}", (ysSum / (cyNum / 3)));
                list[6].month3 = string.Format("{0:N2}", (ysSum / (phNum / 3)));
                list[7].month3 = string.Format("{0:P}", (ysSum / cySum));
                list[8].month3 = string.Format("{0:P}", (ysSum / phSum));
                list[9].month3 = string.Format("{0:N2}", (lrSum / cySum));
                list[10].month3 = string.Format("{0:N2}", (lrSum / phSum));

                ysSum += double.Parse(list[1].month4);//营收累计
                lrSum += double.Parse(list[2].month4);//利润累计
                cyNum += double.Parse(list[11].month4);//Cyber-总人数 累计
                phNum += double.Parse(list[21].month4);//Physical-总人数 累计
                cySum += double.Parse(list[41].month4);//Cyber-人工成本
                phSum += double.Parse(list[48].month4);//Physical-人工成本
                list[5].month4 = string.Format("{0:N2}", (ysSum / (cyNum / 4)));
                list[6].month4 = string.Format("{0:N2}", (ysSum / (phNum / 4)));
                list[7].month4 = string.Format("{0:P}", (ysSum / cySum));
                list[8].month4 = string.Format("{0:P}", (ysSum / phSum));
                list[9].month4 = string.Format("{0:N2}", (lrSum / cySum));
                list[10].month4 = string.Format("{0:N2}", (lrSum / phSum));

                ysSum += double.Parse(list[1].month5);//营收累计
                lrSum += double.Parse(list[2].month5);//利润累计
                cyNum += double.Parse(list[11].month5);//Cyber-总人数 累计
                phNum += double.Parse(list[21].month5);//Physical-总人数 累计
                cySum += double.Parse(list[41].month5);//Cyber-人工成本
                phSum += double.Parse(list[48].month5);//Physical-人工成本
                list[5].month5 = string.Format("{0:N2}", (ysSum / (cyNum / 5)));
                list[6].month5 = string.Format("{0:N2}", (ysSum / (phNum / 5)));
                list[7].month5 = string.Format("{0:P}", (ysSum / cySum));
                list[8].month5 = string.Format("{0:P}", (ysSum / phSum));
                list[9].month5 = string.Format("{0:N2}", (lrSum / cySum));
                list[10].month5 = string.Format("{0:N2}", (lrSum / phSum));

                ysSum += double.Parse(list[1].month6);//营收累计
                lrSum += double.Parse(list[2].month6);//利润累计
                cyNum += double.Parse(list[11].month6);//Cyber-总人数 累计
                phNum += double.Parse(list[21].month6);//Physical-总人数 累计
                cySum += double.Parse(list[41].month6);//Cyber-人工成本
                phSum += double.Parse(list[48].month6);//Physical-人工成本
                list[5].month6 = string.Format("{0:N2}", (ysSum / (cyNum / 6)));
                list[6].month6 = string.Format("{0:N2}", (ysSum / (phNum / 6)));
                list[7].month6 = string.Format("{0:P}", (ysSum / cySum));
                list[8].month6 = string.Format("{0:P}", (ysSum / phSum));
                list[9].month6 = string.Format("{0:N2}", (lrSum / cySum));
                list[10].month6 = string.Format("{0:N2}", (lrSum / phSum));

                ysSum += double.Parse(list[1].month7);//营收累计
                lrSum += double.Parse(list[2].month7);//利润累计
                cyNum += double.Parse(list[11].month7);//Cyber-总人数 累计
                phNum += double.Parse(list[21].month7);//Physical-总人数 累计
                cySum += double.Parse(list[41].month7);//Cyber-人工成本
                phSum += double.Parse(list[48].month7);//Physical-人工成本
                list[5].month7 = string.Format("{0:N2}", (ysSum / (cyNum / 7)));
                list[6].month7 = string.Format("{0:N2}", (ysSum / (phNum / 7)));
                list[7].month7 = string.Format("{0:P}", (ysSum / cySum));
                list[8].month7 = string.Format("{0:P}", (ysSum / phSum));
                list[9].month7 = string.Format("{0:N2}", (lrSum / cySum));
                list[10].month7 = string.Format("{0:N2}", (lrSum / phSum));

                ysSum += double.Parse(list[1].month8);//营收累计
                lrSum += double.Parse(list[2].month8);//利润累计
                cyNum += double.Parse(list[11].month8);//Cyber-总人数 累计
                phNum += double.Parse(list[21].month8);//Physical-总人数 累计
                cySum += double.Parse(list[41].month8);//Cyber-人工成本
                phSum += double.Parse(list[48].month8);//Physical-人工成本
                list[5].month8 = string.Format("{0:N2}", (ysSum / (cyNum / 8)));
                list[6].month8 = string.Format("{0:N2}", (ysSum / (phNum / 8)));
                list[7].month8 = string.Format("{0:P}", (ysSum / cySum));
                list[8].month8 = string.Format("{0:P}", (ysSum / phSum));
                list[9].month8 = string.Format("{0:N2}", (lrSum / cySum));
                list[10].month8 = string.Format("{0:N2}", (lrSum / phSum));

                ysSum += double.Parse(list[1].month9);//营收累计
                lrSum += double.Parse(list[2].month9);//利润累计
                cyNum += double.Parse(list[11].month9);//Cyber-总人数 累计
                phNum += double.Parse(list[21].month9);//Physical-总人数 累计
                cySum += double.Parse(list[41].month9);//Cyber-人工成本
                phSum += double.Parse(list[48].month9);//Physical-人工成本
                list[5].month9 = string.Format("{0:N2}", (ysSum / (cyNum / 9)));
                list[6].month9 = string.Format("{0:N2}", (ysSum / (phNum / 9)));
                list[7].month9 = string.Format("{0:P}", (ysSum / cySum));
                list[8].month9 = string.Format("{0:P}", (ysSum / phSum));
                list[9].month9 = string.Format("{0:N2}", (lrSum / cySum));
                list[10].month9 = string.Format("{0:N2}", (lrSum / phSum));

                ysSum += double.Parse(list[1].month10);//营收累计
                lrSum += double.Parse(list[2].month10);//利润累计
                cyNum += double.Parse(list[11].month10);//Cyber-总人数 累计
                phNum += double.Parse(list[21].month10);//Physical-总人数 累计
                cySum += double.Parse(list[41].month10);//Cyber-人工成本
                phSum += double.Parse(list[48].month10);//Physical-人工成本
                list[5].month10 = string.Format("{0:N2}", (ysSum / (cyNum / 10)));
                list[6].month10 = string.Format("{0:N2}", (ysSum / (phNum / 10)));
                list[7].month10 = string.Format("{0:P}", (ysSum / cySum));
                list[8].month10 = string.Format("{0:P}", (ysSum / phSum));
                list[9].month10 = string.Format("{0:N2}", (lrSum / cySum));
                list[10].month10 = string.Format("{0:N2}", (lrSum / phSum));

                ysSum += double.Parse(list[1].month11);//营收累计
                lrSum += double.Parse(list[2].month11);//利润累计
                cyNum += double.Parse(list[11].month11);//Cyber-总人数 累计
                phNum += double.Parse(list[21].month11);//Physical-总人数 累计
                cySum += double.Parse(list[41].month11);//Cyber-人工成本
                phSum += double.Parse(list[48].month11);//Physical-人工成本
                list[5].month11 = string.Format("{0:N2}", (ysSum / (cyNum / 11)));
                list[6].month11 = string.Format("{0:N2}", (ysSum / (phNum / 11)));
                list[7].month11 = string.Format("{0:P}", (ysSum / cySum));
                list[8].month11 = string.Format("{0:P}", (ysSum / phSum));
                list[9].month11 = string.Format("{0:N2}", (lrSum / cySum));
                list[10].month11 = string.Format("{0:N2}", (lrSum / phSum));

                ysSum += double.Parse(list[1].month12);//营收累计
                lrSum += double.Parse(list[2].month12);//利润累计
                cyNum += double.Parse(list[11].month12);//Cyber-总人数 累计
                phNum += double.Parse(list[21].month12);//Physical-总人数 累计
                cySum += double.Parse(list[41].month12);//Cyber-人工成本
                phSum += double.Parse(list[48].month12);//Physical-人工成本
                list[5].month12 = string.Format("{0:N2}", (ysSum / (cyNum / 12)));
                list[6].month12 = string.Format("{0:N2}", (ysSum / (phNum / 12)));
                list[7].month12 = string.Format("{0:P}", (ysSum / cySum));
                list[8].month12 = string.Format("{0:P}", (ysSum / phSum));
                list[9].month12 = string.Format("{0:N2}", (lrSum / cySum));
                list[10].month12 = string.Format("{0:N2}", (lrSum / phSum));

                #endregion


                return list;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }



    }
}
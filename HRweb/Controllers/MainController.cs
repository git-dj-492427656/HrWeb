using BLL.Tools;
using Newtonsoft.Json.Linq;
using System;
using System.Data;
using System.Web;
using System.Web.Mvc;

namespace HRweb.Controllers
{
    public class MainController : HomeController
    {
        //主页
        [MyAuthAttribute]
        public ActionResult Index()
        {
            var cookie = GetCookie();

            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        //退出登录
        public ActionResult Logout()
        {
            var cookie = GetCookie();
            try
            {
                string txtPath = Server.MapPath("~/MidsTxt/");
                Bll.Tools.GisHelper.DelTxt(HttpUtility.UrlDecode(cookie["txtName"], System.Text.Encoding.GetEncoding("GB2312")), txtPath);
            }
            catch (Exception e)
            {
                throw e;
            }

            cookie.Expires = DateTime.Now.AddDays(-1); //时限
            Response.AppendCookie(cookie);

            return RedirectToAction("Login", "Home");
        }


        //修改密码
        [MyAuthAttribute]
        public ActionResult EditPwd() 
        {
            var cookie = GetCookie();

            string name = cookie["name"];
            string comName = cookie["comName"];
            string type = cookie["type"];

            ViewBag.Name = HttpUtility.UrlDecode(name, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.ComName = HttpUtility.UrlDecode(comName, System.Text.Encoding.GetEncoding("GB2312"));
            ViewBag.Type = HttpUtility.UrlDecode(type, System.Text.Encoding.GetEncoding("GB2312"));

            return View();
        }

        public string SavePwd()
        {
            try
            {
                HttpCookie cookie = GetCookie();

                string oldPwd = Request["oldPwd"];
                string newPwd = Request["newPwd"];
                string newPwd2 = Request["newPwd2"];

                if (newPwd.Contains(" "))
                {
                    return "密码不能存在空格字符！";
                }
                if (newPwd == "")
                {
                    return "密码不能为空！";
                }
                if (newPwd != newPwd2)
                {
                    return "两次密码不一致！";
                }

                //调用接口
                webService.Service1 wbs = new webService.Service1();
                DataTable dt = wbs.Getupdatepassdt(cookie["name"], oldPwd, newPwd, newPwd2);

                if (dt.Rows.Count > 0)
                {
                    string result = dt.Rows[0]["passdr"].ToString();
                    if ("密码修改成功！" == result.Trim())
                    {
                        if (System.Web.HttpContext.Current.Request.Cookies["HrLogin123"] != null)
                        {
                            var loginCookie = Request.Cookies["HrLogin123"];
                            loginCookie["pwd"] = Tools.Encrypt(newPwd);
                            Response.AppendCookie(loginCookie);
                        }
                    }

                    return result;
                }

                return "修改失败！";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }


        //切换组织
        [MyAuthAttribute]
        public ActionResult EditCom()
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

        //组织列表
        public string GetComList()
        {
            try
            {
                JArray arry = new JArray();

                HttpCookie cookie = GetCookie();
                string _key = Tools.Tobyte64(cookie["mobileKey"]);
                string _name = cookie["name"];
                string _comName = Request["comName"].Trim();

                //调用接口
                webService.Service1 wbs = new webService.Service1();
                DataTable dt = wbs.Getmidfulllistds(_key); //所有组织列表
                DataTable dt1 = wbs.Getusermidpermds(_name, _key, _comName); //有权限的管理单元

                //比较筛选数据
                if (dt1.Rows.Count > 0 && dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        for (int j = 0; j < dt.Rows.Count; j++)
                        {
                            if (dt.Rows[j]["Mid"].ToString() == dt1.Rows[i]["Mid"].ToString())
                            {
                                JObject json = new JObject();

                                json.Add("id", dt.Rows[j]["Mid"].ToString());
                                json.Add("pId", dt.Rows[j]["Fatherid"].ToString());
                                json.Add("fatherId", dt.Rows[j]["Fatherid"].ToString());
                                json.Add("name", dt.Rows[j]["Mancompany"].ToString());
                                json.Add("number", dt.Rows[j]["ManNumber"].ToString());
                                json.Add("isSub", dt.Rows[j]["IsSub"].ToString());
                                json.Add("status", dt.Rows[j]["Status"].ToString());
                                json.Add("property", dt.Rows[j]["Property"].ToString());
                                json.Add("EASnumber", dt.Rows[j]["EASnumber"].ToString());
                                json.Add("EASkcnumber", dt.Rows[j]["EASkcnumber"].ToString());

                                if ("远大住工" == json["name"].ToString().Trim())
                                {
                                    json.Add("open", true);//第一级默认展开
                                }

                                arry.Add(json);

                                break;
                            }
                        }
                    }
                }

                return arry.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //确认切换组织
        public void ChangeCom()
        {
            try
            {
                HttpCookie cookie = GetCookie();

                //更新组织单元缓存
                string mids = Request["ids"];
                string txtPath = Server.MapPath("~/MidsTxt/");
                Bll.Tools.GisHelper.WriteTxt(HttpUtility.UrlDecode(cookie["txtName"], System.Text.Encoding.GetEncoding("GB2312")), mids, txtPath);

                cookie["comId"] = Request["comId"];
                cookie["comName"] = HttpUtility.UrlEncode(Request["comName"], System.Text.Encoding.GetEncoding("GB2312"));
                cookie["comNumber"] = Request["comNumber"];
                cookie["fatherId"] = Request["fatherId"];
                cookie["isSub"] = Request["isSub"];
                cookie["EASnumber"] = Request["EASnumber"];
                cookie["EASkcnumber"] = Request["EASkcnumber"];

                Response.AppendCookie(cookie);
                
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        //关于
        [MyAuthAttribute]
        public ActionResult About()
        {
            return RedirectToAction("Permission", "Main");
        }

        //无权限
        [MyAuthAttribute]
        public ActionResult Permission()
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


    }
}
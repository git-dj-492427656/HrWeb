using BLL.Tools;
using System;
using System.Data;
using System.Web;
using System.Web.Mvc;
using Newtonsoft.Json.Linq;
using Bll.Tools;

namespace HRweb.Controllers
{
    public class HomeController : Controller
    {

        [AllowAnonymous]
        public ActionResult Login()
        {
            ViewBag.Name = "";
            ViewBag.Pwd = "";
            ViewBag.Check = "";

            if (System.Web.HttpContext.Current.Request.Cookies["HrLogin123"] != null)
            {
                var cookie = System.Web.HttpContext.Current.Request.Cookies.Get("HrLogin123");

                if (cookie["check"] == "1")
                {
                    ViewBag.Name = HttpUtility.UrlDecode(cookie["name"], System.Text.Encoding.GetEncoding("GB2312"));
                    ViewBag.Pwd = Tools.Decrypt(cookie["pwd"]); //解密
                    ViewBag.Check = cookie["check"];
                }
            }

            return View();
        }
        
        [HttpPost]
        public string LoginBtn()
        {
            string username = Request["username"].Trim();
            string password = Request["password"].Trim();
            string status = Request["status"];
            
            var json = new JObject();
            try
            {
                if (string.IsNullOrWhiteSpace(username) && string.IsNullOrWhiteSpace(password))
                {
                    json.Add("result", false);
                    json.Add("msg", "用户或密码为空！");
                }
                else
                {
                    //调用接口登录方法
                    webService.Service1 wbs = new webService.Service1();
                    DataTable logindt = wbs.Getlogindt(username, password);

                    if (logindt.Rows.Count > 0)
                    {
                        int mid = int.Parse(logindt.Rows[0]["mid"].ToString());

                        // 获取有权限管理单元
                        string mids = "1";
                        DataTable dt = wbs.Getmidfulllistdsbyusermid(username, int.Parse(logindt.Rows[0]["mid"].ToString()));
                        if (dt.Rows.Count > 0)
                        {
                            mids = dt.Rows[0]["midful"].ToString();
                        }

                        //设置管理单元集合缓存
                        DateTime _now = DateTime.Now;
                        string txtName = username + "#" + _now.ToString("yyyy-MM-dd") + "#" + _now.Millisecond;
                        string txtPath = Server.MapPath("~/MidsTxt/");

                        GisHelper.CheckTxt(username, txtPath);
                        GisHelper.WriteTxt(txtName, mids, txtPath);

                        // 设置cookie
                        SetLoginCookie(username, password, status);
                        SetCookie(logindt, txtName);

                        json.Add("result", true);
                        json.Add("msg", "登录成功！");
                        json.Add("url", "../Main/Index");
                    }
                    else
                    {
                        json.Add("result", false);
                        json.Add("msg", "用户或密码错误！");
                    }
                }
                
            }
            catch (Exception ex)
            {
                json.Add("result", false);
                json.Add("msg", ex.Message);
            }

            return json.ToString();
        }

        [HttpPost]
        public ActionResult Login(string returnUrl)
        {
            string username = Request["username"].Trim();
            string password = Request["password"].Trim();
            string status = Request["status"];




            //ViewBag.js = "<script>alert('cccc')</script>";

            //return Content("<script>alert('cccc')</script>");

            //return Content("<script>$('#result').html('登录失败！');$('#myModal').modal('show');history.go(-1);</script>");



            return View();


            //return RedirectToAction("Index", "Main");
        }

        // 登录cookie
        private void SetLoginCookie(string name, string pwd, string check)
        {

            if (System.Web.HttpContext.Current.Request.Cookies["HrLogin123"] != null) //判断cookie存在
            {
                var cookie = System.Web.HttpContext.Current.Request.Cookies.Get("HrLogin123"); //获取cookie
                if ("1" == check)  //更新cookie
                {
                    cookie["check"] = check;
                    cookie["pwd"] = Tools.Encrypt(pwd);
                    cookie["name"] = HttpUtility.UrlEncode(name, System.Text.Encoding.GetEncoding("GB2312"));

                    cookie.Expires = DateTime.Now.AddDays(+30);
                }
                else
                {
                    cookie.Expires = DateTime.Now.AddDays(-1); //时限
                }

                Response.AppendCookie(cookie);
            }
            else
            {
                if ("1" == check) //新建cookie
                {
                    var cookie = new HttpCookie("HrLogin123");
                    cookie.Expires = DateTime.Now.AddDays(+30); //时限

                    cookie.Values.Add("check", check);
                    cookie.Values.Add("pwd", Tools.Encrypt(pwd)); //加密
                    cookie.Values.Add("name", HttpUtility.UrlEncode(name, System.Text.Encoding.GetEncoding("GB2312")));

                    System.Web.HttpContext.Current.Response.Cookies.Add(cookie);
                }
            }
        }

        public void SetCookie(DataTable dt, string txtName)
        {
            HttpCookie cookie = System.Web.HttpContext.Current.Request.Cookies.Get("ydhr123");
            if (cookie == null)
            {
                cookie = new HttpCookie("ydhr123");
                cookie.Expires = DateTime.Now.AddDays(+1); //时限

                cookie.Values.Add("name", HttpUtility.UrlEncode(dt.Rows[0]["username"].ToString(), System.Text.Encoding.GetEncoding("GB2312"))); //登录名
                cookie.Values.Add("userId", dt.Rows[0]["mmid"].ToString());  //用户id
                cookie.Values.Add("mName", HttpUtility.UrlEncode(dt.Rows[0]["menname"].ToString(), System.Text.Encoding.GetEncoding("GB2312"))); //用户名
                cookie.Values.Add("type", HttpUtility.UrlEncode(dt.Rows[0]["usertype"].ToString(), System.Text.Encoding.GetEncoding("GB2312"))); //用户类型
                cookie.Values.Add("loginDate", dt.Rows[0]["lastdate"].ToString());  //登录日期
                cookie.Values.Add("comName", HttpUtility.UrlEncode(dt.Rows[0]["ManCompany"].ToString(), System.Text.Encoding.GetEncoding("GB2312"))); //当前公司
                cookie.Values.Add("comNumber", dt.Rows[0]["ManNumber"].ToString());  //公司编码
                cookie.Values.Add("fatherId", dt.Rows[0]["Fatherid"].ToString());   
                cookie.Values.Add("isSub", dt.Rows[0]["IsSub"].ToString());
                cookie.Values.Add("comId", dt.Rows[0]["mid"].ToString());    //管理单元Id
                cookie.Values.Add("mobileKey", dt.Rows[0]["MobileKey"].ToString());

                cookie.Values.Add("EASnumber", dt.Rows[0]["EASnumber"].ToString());  //eas编码
                cookie.Values.Add("EASkcnumber", dt.Rows[0]["EASkcnumber"].ToString()); //eas库存编码

                cookie.Values.Add("txtName", HttpUtility.UrlEncode(txtName, System.Text.Encoding.GetEncoding("GB2312"))); //缓存文件名称

                Response.Cookies.Add(cookie);
            }
            else
            {
                cookie["userId"] = dt.Rows[0]["mmid"].ToString();
                cookie["name"] = HttpUtility.UrlEncode(dt.Rows[0]["username"].ToString(), System.Text.Encoding.GetEncoding("GB2312"));
                cookie["mName"] = HttpUtility.UrlEncode(dt.Rows[0]["menname"].ToString(), System.Text.Encoding.GetEncoding("GB2312"));
                cookie["type"] = HttpUtility.UrlEncode(dt.Rows[0]["usertype"].ToString(), System.Text.Encoding.GetEncoding("GB2312"));
                cookie["loginDate"] = dt.Rows[0]["lastdate"].ToString();
                cookie["comName"] = HttpUtility.UrlEncode(dt.Rows[0]["ManCompany"].ToString(), System.Text.Encoding.GetEncoding("GB2312"));
                cookie["comNumber"] = dt.Rows[0]["ManNumber"].ToString();
                cookie["fatherId"] = dt.Rows[0]["Fatherid"].ToString();
                cookie["isSub"] = dt.Rows[0]["IsSub"].ToString();
                cookie["comId"] = dt.Rows[0]["mid"].ToString();
                cookie["mobileKey"] = dt.Rows[0]["MobileKey"].ToString();

                cookie["EASnumber"] = dt.Rows[0]["EASnumber"].ToString();
                cookie["EASkcnumber"] = dt.Rows[0]["EASkcnumber"].ToString();

                cookie["txtName"] = HttpUtility.UrlEncode(txtName, System.Text.Encoding.GetEncoding("GB2312"));

                cookie.Expires = DateTime.Now.AddDays(+1); //时限

                Response.AppendCookie(cookie);
            }
        }

        public static HttpCookie GetCookie()
        {
            HttpCookie cookie = null;
            if (System.Web.HttpContext.Current.Request.Cookies["ydhr123"] != null)
            {
                cookie = System.Web.HttpContext.Current.Request.Cookies.Get("ydhr123");
            }
            return cookie;
        }
        
        [NonAction]
        public static bool IsLogin()
        {
            if (null == GetCookie())
            {
                return false;
            }

            return true;
        }




        public ActionResult Error()
        {

            ViewBag.Name = "TTT222";
            ViewBag.ComName = "XXXX222";
            ViewBag.Type = "Admin2";

            return View();
        }


        


    }
}
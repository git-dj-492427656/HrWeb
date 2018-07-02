using System.Web;
using System.Web.Mvc;

namespace HRweb.Controllers
{
    public class MyAuthAttribute : AuthorizeAttribute
    {
        // 只需重载此方法，模拟自定义的角色授权机制 
        public string Model { set; get; }
        protected override bool AuthorizeCore(HttpContextBase httpContext)
        {

            if (!HomeController.IsLogin())
                return false;

            return true;
        }

        public override void OnAuthorization(AuthorizationContext filterContext)
        {
            base.OnAuthorization(filterContext);
            if (filterContext.HttpContext.Response.StatusCode == 403)
            {
                filterContext.Result = new RedirectResult("/Shared/Error");
            }
        }
    }
}
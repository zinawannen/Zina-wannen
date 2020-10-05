using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GRC.Controllers
{
    public class LangController : BaseController
    {
        //
        // GET: /Lang/
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult SetCulture(string culture)
        {
            // Validate input
            culture = CultureHelper.GetImplementedCulture(culture);

            RouteData.Values["culture"] = culture;

            // Enregistrer la culture dans les cookies
            HttpCookie cookie = Request.Cookies["_culture"];
            if (cookie != null)
                cookie.Value = culture;   // Màj de la valeur de la cookie
            else
            {

                cookie = new HttpCookie("_culture");
                cookie.Value = culture;
                cookie.Expires = DateTime.Now.AddYears(1);
            }
            Response.Cookies.Add(cookie);

            //return Redirect(Request.UrlReferrer.ToString());
            return RedirectToAction("Index","Home");
            //http://stackoverflow.com/questions/815229/how-do-i-redirect-to-the-previous-action-in-asp-net-mvc
        }  
	}
}
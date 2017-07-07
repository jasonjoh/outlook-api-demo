using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Cookies;
using Microsoft.Owin.Security.OpenIdConnect;
using OutlookAPIDemo.TokenStorage;
using System.Security.Claims;
using System.Web;
using System.Web.Mvc;

namespace OutlookAPIDemo.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            if (Request.IsAuthenticated)
            {
                string userName = ClaimsPrincipal.Current.FindFirst("name").Value;
                string userId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
                if (string.IsNullOrEmpty(userName) || string.IsNullOrEmpty(userId))
                {
                    // Invalid principal, sign out
                    return RedirectToAction("SignOut");
                }

                // Since we cache tokens in the session, if the server restarts
                // but the browser still has a cached cookie, we may be
                // authenticated but not have a valid token cache. Check for this
                // and force signout.
                SessionTokenCache tokenCache = new SessionTokenCache(userId, HttpContext);
                if (!tokenCache.HasData())
                {
                    // Cache is empty, sign out
                    return RedirectToAction("SignOut");
                }

                ViewBag.UserName = userName;
            }

            return View();
        }

        public void SignIn()
        {
            if (!Request.IsAuthenticated)
            {
                // Signal OWIN to send an authorization request to Azure
                HttpContext.GetOwinContext().Authentication.Challenge(
                    new AuthenticationProperties { RedirectUri = "/" },
                    OpenIdConnectAuthenticationDefaults.AuthenticationType);
            }
        }

        public void SignOut()
        {
            if (Request.IsAuthenticated)
            {
                string userId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;

                if (!string.IsNullOrEmpty(userId))
                {
                    // Get the user's token cache and clear it
                    SessionTokenCache tokenCache = new SessionTokenCache(userId, HttpContext);
                    tokenCache.Clear();
                }
            }
            // Send an OpenID Connect sign-out request. 
            HttpContext.GetOwinContext().Authentication.SignOut(
                CookieAuthenticationDefaults.AuthenticationType);
            Response.Redirect("/");
        }

        public ActionResult Error(string message, string debug)
        {
            ViewBag.Message = message;
            ViewBag.Debug = debug;
            return View("Error");
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using OutlookService;

namespace OutlookAPIDemo.Controllers
{
    public class RoomsController : BaseController
    {
        // GET: Rooms
        public async Task<ActionResult> Index()
        {
            string token = await GetAccessToken();
            if (string.IsNullOrEmpty(token))
            {
                // If there's no token in the session, redirect to Home
                return Redirect("/");
            }

            OutlookService.OutlookService service = new OutlookService.OutlookService(Apiversion.Beta);
            var response = await service.MakeApiCall(token, "/Me");

            return Content(string.Format("Token: {0}", token));
        }
    }
}
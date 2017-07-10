using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using OutlookService;
using OutlookService.Entities;
using OutlookAPIDemo.Models;

namespace OutlookAPIDemo.Controllers
{
    public class RoomsController : BaseController
    {
        // GET: Rooms
        [Authorize]
        public async Task<ActionResult> Index()
        {
            string token = await GetAccessToken();
            if (string.IsNullOrEmpty(token))
            {
                // If there's no token in the session, redirect to Home
                return Redirect("/");
            }

            OutlookService.OutlookService service = new OutlookService.OutlookService(Apiversion.Beta);
            var user = await service.GetMe(token);
            service.SetUser(user.EmailAddress);

            // Get room lists
            var roomLists = await service.GetRoomLists(token);

            // Get all rooms
            var allRooms = await service.GetRooms(token);

            // Create the view model
            RoomView view = new RoomView();

            view.AllRooms = allRooms.Items;
            view.RoomLists = new List<KeyValuePair<EmailAddress, List<EmailAddress>>>();

            foreach (EmailAddress roomList in roomLists.Items)
            {
                // Get the rooms in the list
                var roomsInList = await service.GetRooms(token, roomList.Address);

                view.RoomLists.Add(new KeyValuePair<EmailAddress, List<EmailAddress>>(roomList, roomsInList.Items));
            }

            return View(view);
        }
    }
}
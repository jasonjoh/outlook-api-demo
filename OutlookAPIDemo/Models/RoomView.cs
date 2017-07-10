using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using OutlookService.Entities;

namespace OutlookAPIDemo.Models
{
    public class RoomView
    {
        public List<EmailAddress> AllRooms;
        public List<KeyValuePair<EmailAddress, List<EmailAddress>>> RoomLists;
    }
}
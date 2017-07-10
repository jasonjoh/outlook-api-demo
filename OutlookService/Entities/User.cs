using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookService.Entities
{
    public class User
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string EmailAddress { get; set; }
        public string Alias { get; set; }
        public string MailboxGuid { get; set; }
    }
}

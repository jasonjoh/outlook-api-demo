using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace OutlookService.Entities
{
    public class ItemCollection<T>
    {
        [JsonProperty(PropertyName = "@odata.nextLink")]
        public string NextPage { get; set; }
        [JsonProperty(PropertyName = "value")]
        public List<T> Items { get; set; }
    }
}

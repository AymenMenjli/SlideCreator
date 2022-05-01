using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SlideCreator
{
    public class ComponentJson
    {
        public string type { get; set; }
        public string version { get; set; } = "1.0";
        public DataJson data { get; set; }
        public List<double> position { get; set; } = new List<double>() { 0, 0 };
        public EventsJson events { get; set; }
        public string symbol { get; set; }
        public int? size { get; set; }
        public string id { get; set; }
    }
}

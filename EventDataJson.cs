using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SlideCreator
{
    public class EventDataJson
    {
        public string action { get; set; } = "open";
        public int tab { get; set; } = 0;
        public int ecran { get; set; } = 0;
    }
}

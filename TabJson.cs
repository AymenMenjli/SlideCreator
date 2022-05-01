using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SlideCreator
{
    public class TabJson
    {
        public string titre { get; set; }
        public string couleur { get; set; } = "#568974";
        public List<EcranJson> ecrans { get; set; }
    }
}

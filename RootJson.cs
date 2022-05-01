using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SlideCreator
{
    public class RootJson
    {
        public LayoutJson layout { get; set; }
        public List<ComponentJson> components { get; set; }
    }
}

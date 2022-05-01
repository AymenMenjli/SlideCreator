using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SlideCreator
{
    public class ConfigTabData
    {
        public string TabTitle { get; set; }
        public int? ParentEcranID { get; set; } = null;
        public List<ConfigEcranData> EcranList { get; set; } = new List<ConfigEcranData>();
    }
}

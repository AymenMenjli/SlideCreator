using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SlideCreator
{
    public class ClickJson
    {
        public string eventName { get; set; } = "AtsButtonClick";
        public List<string> triggerComponentsId { get; set; } = new List<string>() { "tickers01" };
        public EventDataJson eventData { get; set; }
    }
}

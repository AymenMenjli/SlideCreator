using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SlideCreator
{
    public class ProgressReportModel
    {
        public int PercentageComplete { get; set; } = 0;
        public List<SlideReportDataModel> CreatedSlide { get; set; } = new List<SlideReportDataModel>();

    }
}

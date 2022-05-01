using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SlideCreator
{
    public class DataJson
    {
        public string textWidth { get; set; }
        public string textHeight { get; set; }
        public string textColor { get; set; }
        public string bulletColor { get; set; }
        public string subBulletColor { get; set; }
        public string animation { get; set; }
        public string fontSize { get; set; }
        public string fontName { get; set; }
        public int? animationDelay { get; set; }
        public string textColorSymbole { get; set; }
        public string secondTextColor { get; set; }
        public bool? surround { get; set; }
        public string surrText { get; set; }
        public List<string> content { get; set; }
        public string ScaleWidth { get; set; }
        public string ScaleHeight { get; set; }
        public string FilterColor { get; set; }
        public string FilterOpacity { get; set; }
        public string type { get; set; }
        public string srcFilter { get; set; }
        public string ParalaxBackground { get; set; }
        public string ParalaxForeground { get; set; }
        public string ParalaxMiddleground { get; set; }
        public string level1 { get; set; }
        public string level2 { get; set; }
        public string level3 { get; set; }
        public string number { get; set; }
        public string mouseColor { get; set; }
        public bool? manualsize { get; set; }
        public bool? useAsButton { get; set; }
        public int? rectangleWidth { get; set; }
        public int? rectangleheight { get; set; }
        public int? cornerRadius { get; set; }
        public string textToAdd { get; set; }
        public string textBtnWidth { get; set; }
        public int? textOffsetX { get; set; }
        public int? textOffsetY { get; set; }
        public string rectangleColor { get; set; }
        public bool? addArrow { get; set; }
        public int? ArrowOffsetX { get; set; }
        public int? ArrowOffsetY { get; set; }
        public bool? useMasterMaxSize { get; set; }
        public string wrapperType { get; set; }
        public int? offset { get; set; }
        public List<ComponentJson> components { get; set; }
        public string clipName { get; set; }
        public List<TabJson> tabs { get; set; }
    }
}

using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SlideCreator
{
    
    //public enum SlideType
    //{
    //    ReadingView,
    //    TickersView,
    //    ResumeView
    //}
    public class ConfigSlideData
    {
        public String TitleSlide { get; set; }
        public String SubTitleSlide { get; set; }
        public String ThirdSubTitleSlide { get; set; }
        public String TickerTitleSlide { get; set; }
        public String FirstTitleNumbers { get; set; } = "";
        public String SecondTitleNumbers { get; set; } = "";
        public String ThirdTitleNumbers { get; set; } = "";
        public String AllNumbers { get; set; }
        public List<String> ComponentTypes { get; set; } = new List<string>();
        public List<String> WrapperCompTypes { get; set; } = new List<string>();
        public List<String> ContentTextData { get; set; } = new List<string>();
        public List<String> TickerTitleSlideList { get; set; } = new List<string>();
        public ConfigTickerData SlideTicker { get; set; } = new ConfigTickerData();
        public bool ParentTicker { get; set; } = false;
        public string TickerID { get; set; } = "";
        public int? ChildernEcranID { get; set; } = null;
        public int? ParentEcranID { get; set; } = null;
        public List<ConfigButtonData> ButtonList = new List<ConfigButtonData>();
        public bool IsObjective { get; set; } = false;
        public bool IsResume { get; set; } = false;
        public bool IsQuiz { get; set; } = false;
        public bool ChildEcranWithTabination = false;
        public SlideTypeNameSpace.SlideType SlideTypeView = SlideTypeNameSpace.SlideType.ReadingView;
        public int ObjectivePlanNumber { get; set; }  = 1;

        public void GetTitleSlide(IShape shapesIn)
        {
            //if (shapesIn.PlaceholderFormat == null)
                //return;
            //shapesIn.SlideItemType == SlideItemType.Placeholder && shapesIn.TextBody.Paragraphs.ElementAt(0).TextParts.ElementAt(0).Font.FontSize == 32 || shapesIn.SlideItemType == SlideItemType.Placeholder && shapesIn.TextBody.Paragraphs.ElementAt(0).TextParts.ElementAt(0).Font.FontSize == 24
            bool GroupeCheck = shapesIn.SlideItemType != SlideItemType.GroupShape;
            bool TitleCheck = false;
            bool CenterTitleCheck = false;
            bool SubTitleCheck = false;
            bool ColorRedCheck = false;
            bool FontSizeCheck = false;
            bool ThirdSubTitleCheck = false;
            bool SecondSubTitleCheck = false;
            string TextTocheck = shapesIn.TextBody.Text;
            System.Drawing.Color colorLoc = System.Drawing.ColorTranslator.FromHtml("#D61947");
            if (shapesIn.PlaceholderFormat != null)
            {
                TitleCheck = shapesIn.PlaceholderFormat.Type == PlaceholderType.Title;
                CenterTitleCheck = shapesIn.PlaceholderFormat.Type == PlaceholderType.CenterTitle;
                SubTitleCheck = shapesIn.PlaceholderFormat.Type == PlaceholderType.Subtitle;
            }
            if (shapesIn.TextBody.Paragraphs.Count != 0 && shapesIn.TextBody.Paragraphs.ElementAt(0).TextParts.Count != 0)
            {
                ColorRedCheck = shapesIn.TextBody.Paragraphs.ElementAt(0).Font.Color.SystemColor.Equals(colorLoc);
                FontSizeCheck = shapesIn.TextBody.Paragraphs.ElementAt(0).TextParts.ElementAt(0).Font.FontSize > 29;
                var currColor = shapesIn.TextBody.Paragraphs.ElementAt(0).Font.Color.SystemColor;
            }


            if (GroupeCheck && TitleCheck || GroupeCheck && CenterTitleCheck || GroupeCheck && FontSizeCheck/* || GroupeCheck && ColorRedCheck*/)
            {
                if (!IsQuiz && FirstTitleNumbers == "")
                {
                    //SlideTypeView = SlideType.ReadingView;
                    String RawTitleSlide = shapesIn.TextBody.Text;
                    TitleSlide = new String(RawTitleSlide.SkipWhile(p => !Char.IsLetter(p)).ToArray());
                    TitleSlide = TitleSlide.Replace("\r\n", "").Replace("\r", "").Replace("\n", "");
                    string RawTitleSlideNoSpace = Regex.Replace(RawTitleSlide, @"\s+", "");
                    FirstTitleNumbers = Regex.Match(RawTitleSlideNoSpace, @"\d+(\.\d+)?").Value;
                    if(FirstTitleNumbers == "" && shapesIn.TextBody.Paragraphs[0].ListFormat.Type == ListType.Numbered)
                    {
                        int? BulletedTitleNumber = shapesIn.TextBody.Paragraphs[0].ListFormat.StartValue;
                        if (BulletedTitleNumber != null)
                        {
                            if (BulletedTitleNumber == 0)
                                BulletedTitleNumber = 1;
                            FirstTitleNumbers = BulletedTitleNumber.ToString();
                        }
                            
                    }
                    System.Diagnostics.Debug.WriteLine(TitleSlide, "Title Slide: ");
                }
                else
                {
                    String RawTitleSlide = shapesIn.TextBody.Text;
                    TitleSlide = Regex.Match(RawTitleSlide, @"\d+(\.\d+)?").Value + "-QUIZ";
                }
                
            }
            else
            {
                //string FirstTitleNumberFromText = "";
                //if (TextTocheck.Length > 5)
                //{
                //    FirstTitleNumberFromText = TextTocheck.Substring(0, 5);
                //    char PointCheck = '.';
                //    char EmptyCheck = ' ';
                //    bool TitleToCheck = FirstTitleNumberFromText[1] == PointCheck && FirstTitleNumberFromText[2] == EmptyCheck;
                //    if (TitleToCheck)
                //    {
                //        String RawTitleSlide = shapesIn.TextBody.Text;
                //        TitleSlide = new String(RawTitleSlide.SkipWhile(p => !Char.IsLetter(p)).ToArray());
                //        TitleSlide = TitleSlide.Replace("\r\n", "").Replace("\r", "").Replace("\n", "");
                //        FirstTitleNumbers = Regex.Match(RawTitleSlide, @"\d+(\.\d+)?").Value;
                //    }
                //}
            }
            
            string SecondSubTitleBumber = "";
            if (TextTocheck.Length > 5)
            {
                SecondSubTitleBumber = TextTocheck.Substring(0, 5);
                char PointCheck = '.';
                char EmptyCheck = ' ';
                SecondSubTitleCheck = SecondSubTitleBumber[1] == PointCheck && SecondSubTitleBumber[3] == EmptyCheck && SecondSubTitleBumber[4] != EmptyCheck || SecondSubTitleBumber[1] == PointCheck && SecondSubTitleBumber[3] == PointCheck && SecondSubTitleBumber[4] == EmptyCheck;
            }

            if (GroupeCheck && SubTitleCheck || GroupeCheck && SecondSubTitleCheck)
            {
                //SubTitleSlide = shapesIn.TextBody.Text;
                String RawTitleSlide = shapesIn.TextBody.Text;
                SubTitleSlide = new String(RawTitleSlide.SkipWhile(p => !Char.IsLetter(p)).ToArray());
                SubTitleSlide = SubTitleSlide.Replace("\r\n", "").Replace("\r", "").Replace("\n", "");
                string RawTitleSlideNoSpace = Regex.Replace(RawTitleSlide, @"\s+", "");
                SecondTitleNumbers = Regex.Match(RawTitleSlideNoSpace, @"\d+(\.\d+)?").Value;
                System.Diagnostics.Debug.WriteLine(SubTitleSlide, "Sub Title Slide: ");
            }

            System.Diagnostics.Debug.WriteLine(TextTocheck);
            string ThirdSubTitleBumber = "";
            if (TextTocheck.Length > 5)
            {
                ThirdSubTitleBumber = TextTocheck.Substring(0, 5);
                char PointCheck = '.';
                char EmptyCheck = ' ';
                ThirdSubTitleCheck = ThirdSubTitleBumber[1] == PointCheck && ThirdSubTitleBumber[3] == PointCheck && ThirdSubTitleBumber[4] != EmptyCheck;
            }
            
            if (GroupeCheck && ThirdSubTitleCheck)
            {
                //ThirdSubTitleSlide = shapesIn.TextBody.Text;
                String RawTitleSlide = shapesIn.TextBody.Text;
                ThirdSubTitleSlide = new String(RawTitleSlide.SkipWhile(p => !Char.IsLetter(p)).ToArray());
                string RawTitleSlideNoSpace = Regex.Replace(RawTitleSlide, @"\s+", "");
                ThirdTitleNumbers = Regex.Match(RawTitleSlideNoSpace, @"\d+(\.\d+\.\d+)?").Value;
                System.Diagnostics.Debug.WriteLine(ThirdSubTitleSlide, "Third Sub Title Slide: ");
                System.Diagnostics.Debug.WriteLine(ThirdSubTitleBumber, "Final Number: ");
            }

            if (GroupeCheck && ColorRedCheck)
            {
                String RawTickerTitleSlide = shapesIn.TextBody.Text;
                TickerTitleSlide = new String(RawTickerTitleSlide.SkipWhile(p => !Char.IsLetter(p)).ToArray());
                TickerTitleSlide = TickerTitleSlide.Replace("\r\n", "").Replace("\r", "").Replace("\n", "");
                //TickerTitleSlide = shapesIn.TextBody.Text;
                System.Diagnostics.Debug.WriteLine(TickerTitleSlide, "Ticker Title Slide: ");
            }
            else
            {
                if (shapesIn.Left > 330 && shapesIn.Left < 340 && shapesIn.Top > 62 && shapesIn.Top < 68)
                {
                    String RawTickerTitleSlide = shapesIn.TextBody.Text;
                    TickerTitleSlide = new String(RawTickerTitleSlide.SkipWhile(p => !Char.IsLetter(p)).ToArray());
                    TickerTitleSlide = TickerTitleSlide.Replace("\r\n", "").Replace("\r", "").Replace("\n", "");
                }
            }

            string TickerPaginationCheck = shapesIn.TextBody.Text;
            char SlashCheck = '/';
            char FirstTabinationCheck = '1';
            if (shapesIn.TextBody.Text.Contains("/") && TickerPaginationCheck[1] == SlashCheck)
            {
                if (shapesIn.TextBody.Text[0] != FirstTabinationCheck)
                    ChildEcranWithTabination = true;
                System.Diagnostics.Debug.WriteLine(shapesIn.TextBody.Text);
            }
            System.Diagnostics.Debug.WriteLine(shapesIn.TextBody, "Final Number: ");
            //if(!ComponentTypes.Contains("BackGround"))
            //    ComponentTypes.Add("BackGround");
            //if (!ComponentTypes.Contains("Title"))
            //    ComponentTypes.Add("Title");
            AddComponentTypeToList("BackGround");
            AddComponentTypeToList("Title");
            if(TitleSlide != null)
            {
                if (TitleSlide.Contains("Résumé"))
                {
                    IsResume = true;
                }
                if (TitleSlide.Contains("Objectifs"))
                {
                    IsObjective = true;
                }
            }
            
        }

        public String GetTitleSlideNumberData()
        {
            
            //System.Diagnostics.Debug.WriteLine(FirstTitleNumbers, "Number Collection: ");
            //System.Diagnostics.Debug.WriteLine(SecondTitleNumbers, "Number Collection: ");
            //System.Diagnostics.Debug.WriteLine(ThirdSubTitleSlide, "Number Collection: ");

            if (ThirdTitleNumbers != "")
            {
                AllNumbers = ThirdTitleNumbers;
                System.Diagnostics.Debug.WriteLine(AllNumbers, "Final Number From Third: ");
                return AllNumbers;
            }
            else
            {
                if (SecondTitleNumbers != "")
                {
                    AllNumbers = SecondTitleNumbers;
                    System.Diagnostics.Debug.WriteLine(AllNumbers, "Final Number From Second: ");
                    return AllNumbers;
                }
                else
                {
                    if (FirstTitleNumbers != "")
                    {
                        AllNumbers = FirstTitleNumbers;
                        System.Diagnostics.Debug.WriteLine(AllNumbers, "Final Number From First: ");
                        return AllNumbers;
                    }
                    else
                    {
                        return "";
                    }
                }
            }
        }

        public void GetSlideTextData(IShape shapesIn)
        {
            System.Diagnostics.Debug.WriteLine(shapesIn.ShapeName);
            System.Diagnostics.Debug.WriteLine(shapesIn.TextBody.Text);
            if (shapesIn.ShapeName.Contains("Group") || shapesIn.ShapeName.Contains("Graphic") || shapesIn.ShapeName.Contains("Picture") || shapesIn.ShapeName.Contains("Image") || shapesIn.ShapeName.Contains("Table") || shapesIn.ShapeName.Contains("Straight Connector") || shapesIn.ShapeName.Contains("Connecteur droit") || shapesIn.ShapeName.Contains("Diagramme"))
                return;
            //if (!ComponentTypes.Contains("Text"))
            //    ComponentTypes.Add("Text");
            try
            {
                AddComponentTypeToList("Text");
                bool TitleCheck = false;
                string TextTocheck = shapesIn.TextBody.Text;
                string TitleNumber = "";
                if (TextTocheck.Length > 5)
                {
                    TitleNumber = shapesIn.TextBody.Text.Substring(0, 5);
                    char PointCheck = '.';
                    char EmptyCheck = ' ';
                    TitleCheck = TitleNumber[1] != PointCheck && TitleNumber[2] != EmptyCheck;
                }
                if (shapesIn.TextBody.Text.Length > 3/* && TitleCheck*/ && !IsQuiz)
                {
                    string QuizString = shapesIn.TextBody.Text.Substring(0, 4);
                    IsQuiz = QuizString.ToLower().Equals("quiz");
                    if (IsQuiz)
                    {
                        TitleSlide = FirstTitleNumbers + "-QUIZ";
                        FirstTitleNumbers = "";
                        SecondTitleNumbers = "";
                        ThirdTitleNumbers = "";
                        AllNumbers = null;
                    }
                    System.Diagnostics.Debug.WriteLine(QuizString);
                }

                //shapesIn.SlideItemType == SlideItemType.Placeholder && shapesIn.TextBody.Paragraphs.ElementAt(0).TextParts.ElementAt(0).Font.FontSize == 20 && shapesIn.TextBody.Paragraphs.ElementAt(0).TextParts.ElementAt(0).Font.Color.ToArgb() != 13426152
                System.Drawing.Color colorLocRed = System.Drawing.ColorTranslator.FromHtml("#D61947");
                System.Drawing.Color colorLoc = System.Drawing.ColorTranslator.FromHtml("#CCDDE8");
                bool PlaceHolderBodyCheck = /*shapesIn.SlideItemType != SlideItemType.GroupShape && shapesIn.PlaceholderFormat.Type == PlaceholderType.Body || */shapesIn.ShapeName.Contains("TextBox") || shapesIn.ShapeName.Contains("Text Placeholder") || shapesIn.ShapeName.Contains("Espace réservé du texte");
                bool TickerPaginationTxtCheck = !shapesIn.TextBody.Paragraphs.ElementAt(0).Font.Color.SystemColor.Equals(colorLoc);
                bool TickerTitleTextCheck = !shapesIn.TextBody.Paragraphs.ElementAt(0).Font.Color.SystemColor.Equals(colorLocRed);
                bool TickerNavigationTextCheck = !shapesIn.TextBody.Paragraphs.ElementAt(0).Font.FontName.Equals("Roboto");
                System.Diagnostics.Debug.WriteLine(shapesIn.TextBody);
                if (PlaceHolderBodyCheck && TickerPaginationTxtCheck && TickerTitleTextCheck && TickerNavigationTextCheck)
                {

                    foreach (IParagraph ParagraphS in shapesIn.TextBody.Paragraphs)
                    {
                        System.Diagnostics.Debug.WriteLine(shapesIn.TextBody.Paragraphs.Count);
                        string TextToAdd = BoldCheck(ParagraphS);
                        if (TextToAdd != "")
                        {
                            ContentTextData.Add(TextToAdd);
                        }
                        
                    }
                    //System.Diagnostics.Debug.WriteLine(ContentTextData);

                    //System.Diagnostics.Debug.WriteLine(shapesIn.TextBody.Paragraphs.ElementAt(0).Font.Color.SystemColor, "COLOR: ");
                    //System.Diagnostics.Debug.WriteLine(colorLoc, "COLOR: ");
                }
            }
            catch (Exception)
            {

                //throw;
            }
            
        }

        public string BoldCheck(IParagraph ParagraphIn)
        {
            string OutText = "";
            foreach (ITextPart TextPartS in ParagraphIn.TextParts)
            {
                string LocalText = TextPartS.Text;
                if (TextPartS.Font.Bold == true)
                {
                    LocalText = "##" + TextPartS.Text + "##";
                }
                OutText += LocalText;
            }
            
            if (ParagraphIn.ListFormat.Type == ListType.Bulleted && ParagraphIn.Text != "" && ParagraphIn.ListFormat.BulletCharacter != '–')
            {
                OutText = "_" + OutText;
            }
            if(OutText.Length != 0)
            {
                if (ParagraphIn.ListFormat.BulletCharacter == '–' || OutText.Substring(0, 1) == "-")
                {
                    var regex = new Regex(Regex.Escape("-"));
                    OutText = regex.Replace(OutText, "", 1);
                    OutText = "°" + OutText;
                }
            }

            if (IsObjective)
            {
                if (ParagraphIn.Text.Contains("Plan") && ParagraphIn.Text.Contains(":") || ParagraphIn.Text.Contains("Plan") && ParagraphIn.Text.Length == 4)
                {
                    OutText = "\r\n" + "\r\n" + "\r\n";
                }
                if(ParagraphIn.ListFormat.Type == ListType.Numbered)
                {
                    OutText = "##" + ObjectivePlanNumber + "##" + "- " + OutText;
                    ObjectivePlanNumber++;
                }
                else
                {
                    if(OutText.Length > 5)
                    {
                        if (OutText.Substring(1, 1) == "-" || OutText.Substring(2, 1) == "-")
                        {
                            
                            int IndexOfFirstCharacter = OutText.IndexOf("-", 0) + 1;
                            System.Diagnostics.Debug.WriteLine(IndexOfFirstCharacter);
                            string OutTextWithoutNumber = OutText.Substring(IndexOfFirstCharacter, OutText.Length- IndexOfFirstCharacter);
                            OutText = "##" + OutText.Substring(0, IndexOfFirstCharacter) + "##" + OutTextWithoutNumber;
                        }
                    }
                }
            }
            //if(ParagraphIn.Text == "")
            //{
            //    OutText = " ";
            //}
            //System.Diagnostics.Debug.WriteLine(OutText, "Text Slide: ");
            System.Diagnostics.Debug.WriteLine(ParagraphIn.ListFormat);
            return OutText;
        }

        public void AddComponentTypeToList(string CompT)
        {
            if (!ComponentTypes.Contains(CompT))
                ComponentTypes.Add(CompT);
        }
    }
}

namespace SlideTypeNameSpace
{
    public enum SlideType
    {
        ReadingView,
        TickersView,
        ResumeView,
        ExampleView
    }

    public enum ComponentType
    {
        BackGround,
        Title,
        Text,
        Warpper,
        Button,
        LibraryPick,
        Ticker
    }
}
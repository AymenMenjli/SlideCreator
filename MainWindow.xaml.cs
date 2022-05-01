using System;
using System.IO.Compression;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Newtonsoft.Json;
using Syncfusion.Presentation;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Text.RegularExpressions;
using System.ComponentModel;
using System.Threading;
using System.Diagnostics;


namespace SlideCreator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        CancellationTokenSource cts = new CancellationTokenSource();

        private String FolderPathForPPTX = "";
        private bool CanStart = false;
        public List<ComponentJson> componentsToAdd = new List<ComponentJson>();
        private int CurrentSlideNumber = 0;
        private List<ConfigSlideData> PPTFileSlide = new List<ConfigSlideData>();
        private List<ConfigSlideData> FinalSlideList = new List<ConfigSlideData>();
        private List<ConfigSlideData> TickerList = new List<ConfigSlideData>();
        private string SubPresentationDirectory = "";
        private int DuplicatedFolderNum = 0;
        private BackgroundWorker worker = new BackgroundWorker();
        private int CurrSlideIndex = 0;
        private int CurrSlideCount = 1;
        private List<SlideReportDataModel> SlideModeloutput = new List<SlideReportDataModel>();
        private string CurrentPresentationName = "";
        //enum SlideType
        //{
        //    ReadingView,
        //    TickersView,
        //    ResumeView
        //}

        //private SlideType SlideTypeView = SlideType.ReadingView;
        public MainWindow()
        {
            InitializeComponent();
            SetupUnhandledExceptionHandling();
            //Progress Bar in Diffrent thread
            //worker.WorkerSupportsCancellation = true;
            //worker.WorkerReportsProgress = true;
            //worker.DoWork += Worker_DoWork;
            //worker.ProgressChanged += Worker_ProgressChanged;
        }

        private async void MainClickButton(object sender, RoutedEventArgs e)
        {
            //Progress<int> progress = new Progress<int>(value => {
            //    SlideProgressBar.Value = value;
            //});
            Progress<ProgressReportModel> progress = new Progress<ProgressReportModel>();
            progress.ProgressChanged += ReportProgress;

            var watch = System.Diagnostics.Stopwatch.StartNew();

            FolderPathForPPTX = FolderPathPPTX.Text;

            //var results = await GetPPTXFiles(progress, cts.Token);
            try
            {
                await Task.Run(() => GetPPTXFiles(progress, cts.Token));
            }
            catch (OperationCanceledException)
            {

                resultsWindow.Text += $"The async download was cancelled. { Environment.NewLine }";
            }
            //PrintResults(results);

            watch.Stop();
            var elapsedMs = watch.ElapsedMilliseconds;

            resultsWindow.Text += $"Total execution time: { elapsedMs }";
            //worker.RunWorkerAsync();

            //GetPPTXFiles();

            //System.Diagnostics.Debug.WriteLine(FinalSlideList);

            System.Windows.MessageBox.Show("Done");
        }

        private void ReportProgress(object sender, ProgressReportModel e)
        {
            SlideProgressBar.Value = e.PercentageComplete;
            try
            {
                PrintResults(e.CreatedSlide);
            }
            catch (Exception)
            {

                //throw;
            }
        }

        private void PrintResults(List<SlideReportDataModel> results)
        {
            resultsWindow.Text = "";
            try
            {
                foreach (var item in results)
                {
                    resultsWindow.Text += $"Creating Competence: { item.SlideParent }  Slide Name: { item.SlideName } .{ Environment.NewLine }";
                }
            }
            catch (Exception)
            {

                //throw;
            }
        }

        private async Task GetPPTXFiles(IProgress<ProgressReportModel> progress, CancellationToken cancellationToken)
        {

            //List<SlideReportDataModel> output = new List<SlideReportDataModel>();
            //ProgressReportModel report = new ProgressReportModel();

            //FolderPathForPPTX = FolderPathPPTX.Text;
            //FolderPathForPPTX = "C:/Users/aymen.menjli/Documents/Aymen menjli/TestBed/TestAppPPTX";
            var TempFilesPathPPTX = Directory.GetFiles(FolderPathForPPTX, "*.pptx", SearchOption.AllDirectories);
            

            for (int pi = 0; pi < TempFilesPathPPTX.Length; pi++)
            {
                IPresentation pptxDoc = Presentation.Open(TempFilesPathPPTX[pi]);
                string PresentationFileName = System.IO.Path.GetFileNameWithoutExtension(TempFilesPathPPTX[pi]);
                string PresentationDirectory = @"" + FolderPathForPPTX + "/" + PresentationFileName;
                CurrentPresentationName = PresentationFileName;
                if (!Directory.Exists(PresentationDirectory))
                {
                    Directory.CreateDirectory(PresentationDirectory);
                }

                CurrentSlideNumber = 0;
                PPTFileSlide.Clear();
                
                foreach (ISlide slide in pptxDoc.Slides)
                {
                    
                    SubPresentationDirectory = PresentationDirectory;
                    
                    ConfigSlideData SlideData = new ConfigSlideData();
                    if (slide.LayoutSlide.Name == "1_Title Slide")
                    {
                        SlideData.SlideTypeView = SlideTypeNameSpace.SlideType.ReadingView;
                    }
                    if (slide.LayoutSlide.Name == "2_Title Slide")
                    {
                        SlideData.SlideTypeView = SlideTypeNameSpace.SlideType.TickersView;
                    }

                    if (slide.LayoutSlide.Name == "Title Slide")
                    {
                        SlideData.SlideTypeView = SlideTypeNameSpace.SlideType.ExampleView;
                    }

                    System.Diagnostics.Debug.WriteLine(slide);
                    bool SlideNameCheck = slide.LayoutSlide.Name == "Title Slide";
                    //"1_Title Slide" "2_Title Slide"
                    foreach (IShape shapes in slide.Shapes)
                    {
                        shapes.TextBody.Text.Equals("Introduction");
                        System.Diagnostics.Debug.WriteLine(shapes.TextBody.Text);
                        string TextToCheck = Regex.Replace(shapes.TextBody.Text, @"\s+", "");
                        if (TextToCheck.Equals("Introduction") && !SlideNameCheck || TextToCheck.Equals("Objectifs") && !SlideNameCheck || TextToCheck.Equals("Objectif") && !SlideNameCheck)
                        {
                            CanStart = true;
                        }
                        string TextToCheckEval = Regex.Replace(shapes.TextBody.Text, @"\s+", "");
                        if (/*SlideNameCheck || */TextToCheckEval.Equals("Evaluation") || TextToCheckEval.Equals("Plan") || TextToCheckEval.Equals("Evaluation:"))
                        {
                            CanStart = false;
                        }
                    }

                    foreach (IShape shapes in slide.Shapes)
                    {
                        if (CanStart)
                        {
                            SlideData.GetTitleSlide(shapes);
                            SlideData.GetSlideTextData(shapes);
                            //GetTextDataSlide(shapes);
                            //System.Diagnostics.Debug.WriteLine(SlideData.ContentTextData);
                        }
                    }

                    if (CanStart)
                    {

                        //CurrentSlideNumber++;
                        //SubPresentationDirectory = SubPresentationDirectory + "/" + CurrentSlideNumber;
                        //System.Diagnostics.Debug.WriteLine(SubPresentationDirectory);
                        if (!SlideData.IsQuiz)
                        {
                            string SlideNumber = SlideData.GetTitleSlideNumberData();
                        }
                        //****create componenet and generate json*****//
                        //CreateComponentByType(SlideData, SubPresentationDirectory);
                        PPTFileSlide.Add(SlideData);
                    }
                    //CurrSlideCount = pptxDoc.Slides.Count;
                    //CurrSlideIndex++;
                    //worker.ReportProgress((int)(CurrSlideIndex * 100 / CurrSlideCount));
                    //ReportFinishSlide(CurrSlideIndex, pptxDoc.Slides.Count);
                    //worker.ReportProgress((int)(CurrSlideIndex * 100/ pptxDoc.Slides.Count));
                    //System.Diagnostics.Debug.WriteLine(CurrSlideIndex * 100 / pptxDoc.Slides.Count);

                }
                SetTickerParent(PPTFileSlide);
                GenerateParentTickerData(PPTFileSlide, TickerList);
                GenerateCleanSlidesList(PPTFileSlide);
                System.Diagnostics.Debug.WriteLine(TickerList);
                //FindDuplicateTitleName(FinalSlideList);
                CreateSlideButtonData(FinalSlideList);
                CurrSlideCount = FinalSlideList.Count;
                System.Diagnostics.Debug.WriteLine(FinalSlideList);
                //SlideReportDataModel results = await CreateComponent(FinalSlideList, PresentationDirectory);
                //CreateComponent(FinalSlideList, PresentationDirectory, progress, cancellationToken);
                await Task.Run(() => CreateComponent(FinalSlideList, PresentationDirectory, progress, cancellationToken));
                //output.Add(results);
                //report.CreatedSlide = output;
                //report.PercentageComplete = (output.Count * 100) / CurrSlideCount;
                //var PercentageComplete = (CurrSlideIndex * 100) / CurrSlideCount;
                //progress.Report(PercentageComplete);

                //System.Diagnostics.Debug.WriteLine(FinalSlideList);
                PPTFileSlide.Clear();
                TickerList.Clear();
                FinalSlideList.Clear();
                CurrSlideIndex = 0;
                DuplicatedFolderNum = 0;
                SlideModeloutput.Clear();
                cancellationToken.ThrowIfCancellationRequested();
            }
            //return output;
        } 

        private void ReportFinishSlide(/*int IndexIn, int SlideCountIn*/)
        {
            
            //while (CurrSlideCount != CurrSlideIndex)
            //{
            //    //worker.ReportProgress((int)(CurrSlideIndex * 100 / CurrSlideCount));
            //    //System.Diagnostics.Debug.WriteLine(FinalSlideList);
            //}
            
        }
        void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            //ReportFinishSlide();
            //GetPPTXFiles();
        }

        void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            SlideProgressBar.Value = e.ProgressPercentage;
            if(SlideProgressBar.Value == 100)
            {
                SlideProgressBar.Value = 0;
                CurrSlideIndex = 0;
                CurrSlideCount = 1;
            }
        }

        private async Task CreateComponent(List<ConfigSlideData> FSldIn, string PDir, IProgress<ProgressReportModel> progress, CancellationToken cancellationToken)
        {
            

            foreach (ConfigSlideData FSldData in FSldIn)
            {
                CurrentSlideNumber++;
                string CurrentSlideNumberName = "";
                if (FSldData.FirstTitleNumbers != "")
                    CurrentSlideNumberName = FSldData.FirstTitleNumbers;
                if (FSldData.SecondTitleNumbers != "")
                    CurrentSlideNumberName = FSldData.SecondTitleNumbers;
                if (FSldData.ThirdTitleNumbers != "")
                    CurrentSlideNumberName = FSldData.ThirdTitleNumbers;
                if (FSldData.AllNumbers == null && FSldData.TitleSlide != null)
                {
                    string TitleSlideNoSpace = Regex.Replace(FSldData.TitleSlide, @"\s+", "");
                    CurrentSlideNumberName = TitleSlideNoSpace;
                }
                    
                if (FSldData.SubTitleSlide != null && FSldData.SubTitleSlide.Contains("(suite)"))
                    CurrentSlideNumberName += "-suite";
                if (FSldData.ThirdSubTitleSlide != null && FSldData.ThirdSubTitleSlide.Contains("(suite)"))
                    CurrentSlideNumberName += "-suite";
                if (FSldData.TitleSlide != null && FSldData.TitleSlide.Contains("(suite)") && !CurrentSlideNumberName.Contains("(suite)"))
                    CurrentSlideNumberName += "-suite";

                if (CurrentSlideNumberName != null && CurrentSlideNumberName.Contains("Résumé"))
                    CurrentSlideNumberName = CurrentSlideNumberName.Replace("(suite)", "");

                CurrentSlideNumberName = CurrentSlideNumberName.Replace(".", "-");
                SubPresentationDirectory = SubPresentationDirectory + "/" + CurrentSlideNumberName;
                System.Diagnostics.Debug.WriteLine(SubPresentationDirectory);
                //****create componenet and generate json*****//
                SlideReportDataModel output = new SlideReportDataModel();
                output.SlideName = CurrentSlideNumberName;
                output.SlideParent = CurrentPresentationName;
                //output.SlideAllNumber = await CreateComponentByType(FSldData, SubPresentationDirectory);
                //CreateComponentByType(FSldData, SubPresentationDirectory);
                await Task.Run(() => CreateComponentByType(FSldData, SubPresentationDirectory, progress, cancellationToken, output));
                
                SubPresentationDirectory = PDir;
            }
            //return output;
        }

        private void SetTickerParent(List<ConfigSlideData> DataIn)
        {
            bool StartCountTicker = false;
            int TickerIndex = 0;
            for (int i = 0; i < DataIn.Count; i++)
            {
                if(DataIn[i].SlideTypeView == SlideTypeNameSpace.SlideType.TickersView && !DataIn[i-1].ParentTicker && DataIn[i-1].SlideTypeView == SlideTypeNameSpace.SlideType.ReadingView)
                {
                    TickerIndex++;
                    DataIn[i - 1].ParentTicker = true;
                    DataIn[i - 1].AddComponentTypeToList("Ticker");
                    DataIn[i - 1].TickerID = "tickers" + TickerIndex;
                    StartCountTicker = true;
                }
                if(StartCountTicker && DataIn[i].SlideTypeView == SlideTypeNameSpace.SlideType.ReadingView)
                {
                    StartCountTicker = false;
                }
                if(DataIn[i].SlideTypeView == SlideTypeNameSpace.SlideType.TickersView && StartCountTicker)
                {
                    DataIn[i].TickerID = "tickers" + TickerIndex;
                    TickerList.Add(DataIn[i]);
                }
            }
        }

        private void GenerateParentTickerData(List<ConfigSlideData> DataIn, List<ConfigSlideData> TickerIn)
        {
            for (int k = 0; k < DataIn.Count; k++)
            {
                if (DataIn[k].ParentTicker)
                {
                    //foreach (ConfigSlideData STData in TickerIn)
                    GenerateTabEcran(DataIn[k], TickerIn);
                    GenerateChildrenEcran(DataIn[k], TickerIn);
                }
            }
        }

        private void GenerateTabEcran(ConfigSlideData DataIn, List<ConfigSlideData> TickerIn)
        {
            int ParentEcranIndex = 1;
            int GEcranIndex = 1;
            for (int Tk = 0; Tk < TickerIn.Count; Tk++)
            {
                if (TickerIn[Tk].TickerID == DataIn.TickerID)
                {
                    if (DataIn.SlideTicker.TabList.Count != 0)
                    {
                        //possible fix: make function that return config tab data to add and then create a list of tabs and ecran to add later to the slide tab
                        //TODO: find solution for the ticker with same titles and are not the same tab, the title is the same but have a diffrent ending
                        if (TickerIn[Tk].TickerTitleSlide != null && !TickerIn[Tk].TickerTitleSlide.Contains(TickerIn[Tk - ParentEcranIndex].TickerTitleSlide)/* || !TickerIn[Tk].TickerTitleSlide.Contains(TickerIn[ParentEcranIndex].TickerTitleSlide)*/)
                        {
                            TickerIn[Tk].ParentEcranID = GEcranIndex;
                            ParentEcranIndex = 1;

                            ConfigTabData ConfigTabDToAdd = new ConfigTabData
                            {
                                TabTitle = TickerIn[Tk].TickerTitleSlide,
                                ParentEcranID = GEcranIndex
                            };

                            ConfigEcranData ConfigEcranToAdd = new ConfigEcranData
                            {
                                EContentTextData = TickerIn[Tk].ContentTextData
                            };

                            ConfigTabDToAdd.EcranList.Add(ConfigEcranToAdd);

                            DataIn.SlideTicker.TabList.Add(ConfigTabDToAdd);
                            GEcranIndex++;
                        }
                        else
                        {
                            TickerIn[Tk].ChildernEcranID = GEcranIndex - 1;
                            DataIn.ChildernEcranID = GEcranIndex - 1;
                            ParentEcranIndex++;
                            
                        }

                    }
                    else
                    {
                        TickerIn[Tk].ParentEcranID = GEcranIndex;
                        ConfigTabData ConfigTabDToAdd1 = new ConfigTabData
                        {
                            TabTitle = TickerIn[Tk].TickerTitleSlide,
                            ParentEcranID = GEcranIndex
                        };

                        ConfigEcranData ConfigEcranToAdd1 = new ConfigEcranData
                        {
                            EContentTextData = TickerIn[Tk].ContentTextData
                        };

                        ConfigTabDToAdd1.EcranList.Add(ConfigEcranToAdd1);

                        DataIn.SlideTicker.TabList.Add(ConfigTabDToAdd1);
                        GEcranIndex++;
                    }
                    System.Diagnostics.Debug.WriteLine(TickerIn[Tk].TickerTitleSlide);
                }
                
            }
        }

        private void GenerateChildrenEcran(ConfigSlideData DataIn, List<ConfigSlideData> TickerIn)
        {
            for (int ce = 0; ce < TickerIn.Count; ce++)
            {
                System.Diagnostics.Debug.WriteLine(TickerIn[ce].TickerTitleSlide);
                if (TickerIn[ce].TickerID == DataIn.TickerID)
                {
                    System.Diagnostics.Debug.WriteLine(TickerIn[ce].ParentEcranID);
                    List<ConfigTabData> LocalTabData = DataIn.SlideTicker.TabList;
                    foreach (ConfigTabData ParentTabData in LocalTabData)
                    {
                        //bool FirstCheck = TickerIn[ce].TickerTitleSlide != null && TickerIn[ce].TickerTitleSlide.Contains(ParentTabData.TabTitle) && ParentTabData.TabTitle != TickerIn[ce].TickerTitleSlide && TickerIn[ce].ChildernEcranID == ParentTabData.ParentEcranID;
                        bool FirstCheck = TickerIn[ce].TickerTitleSlide != null && TickerIn[ce].TickerTitleSlide.Contains(ParentTabData.TabTitle) && TickerIn[ce].TickerTitleSlide.Contains("(suite)") && TickerIn[ce].ChildernEcranID == ParentTabData.ParentEcranID;
                        bool AlternativeFirstCheck = TickerIn[ce].TickerTitleSlide != null && TickerIn[ce].TickerTitleSlide.Contains(ParentTabData.TabTitle) && ParentTabData.TabTitle == TickerIn[ce].TickerTitleSlide && TickerIn[ce].ChildEcranWithTabination && TickerIn[ce].ChildernEcranID == ParentTabData.ParentEcranID;
                        if (FirstCheck || AlternativeFirstCheck)
                        {
                            ConfigEcranData ConfigEcranToAdd2 = new ConfigEcranData();
                            ConfigEcranToAdd2.EContentTextData = TickerIn[ce].ContentTextData;

                            ParentTabData.EcranList.Add(ConfigEcranToAdd2);
                            System.Diagnostics.Debug.WriteLine(TickerIn[ce].TickerTitleSlide);
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine(TickerIn[ce].TickerTitleSlide);
                        }
                        
                    }
                    
                }
                
            }
        }

        private void GenerateCleanSlidesList(List<ConfigSlideData> SlideIn)
        {
            foreach (ConfigSlideData SldData in SlideIn)
            {
                if (SldData.SlideTypeView == SlideTypeNameSpace.SlideType.ReadingView)
                {
                    FinalSlideList.Add(SldData);
                }
            }
        }

        private void FindDuplicateTitleName(List<ConfigSlideData> SlideIn)
        {
            //List<ConfigSlideData> SlideInCopy = SlideIn;
            for (int i = 0; i < SlideIn.Count; i++)
            {
                int DuplicateNum = 1;
                //int keyIndex = SlideIn.FindIndex(w => w.TitleSlide.Equals(SlideIn[i].TitleSlide) && w.ContentTextData != SlideIn[i].ContentTextData /*&& w.FirstTitleNumbers == SlideIn[i].FirstTitleNumbers && w.SecondTitleNumbers == SlideIn[i].SecondTitleNumbers && w.ThirdTitleNumbers == SlideIn[i].ThirdTitleNumbers*/);
                //int keyIndex = SlideIn.IndexOf(SlideIn[i - 1].TitleSlide, i);
                //System.Diagnostics.Debug.WriteLine(SlideIn[keyIndex].TitleSlide);
                foreach (ConfigSlideData SLD in SlideIn)
                {
                    
                    if (SlideIn[i].TitleSlide == SLD.TitleSlide && SlideIn[i].ContentTextData != SLD.ContentTextData)
                    {
                        SlideIn[i].TitleSlide = SlideIn[i].TitleSlide + DuplicateNum;
                        DuplicateNum++;
                        System.Diagnostics.Debug.WriteLine(SlideIn[i].TitleSlide);
                    }
                }
            }
            
        }
        private void CreateSlideButtonData(List<ConfigSlideData> SlideIn)
        {
            foreach (ConfigSlideData FinalSlide in SlideIn)
            {
                List<ConfigTabData> CurrTabList = FinalSlide.SlideTicker.TabList;

                for (int tl = 0; tl < CurrTabList.Count; tl++)
                {
                    ConfigButtonData ButtonToAdd = new ConfigButtonData();
                    ButtonToAdd.ButtonTitle = CurrTabList[tl].TabTitle;
                    ButtonToAdd.ButtonID = FinalSlide.TickerID;
                    ButtonToAdd.ButtonTab = tl;
                    FinalSlide.ButtonList.Add(ButtonToAdd);
                    
                }
                if(FinalSlide.ButtonList.Count != 0)
                {
                    FinalSlide.WrapperCompTypes.Add("Button");
                    FinalSlide.AddComponentTypeToList("Warpper");
                }
            }
        }
        private void GetTextDataSlide(IShape shapesIn)
        {
            if (shapesIn.PlaceholderFormat == null)
                return;
            //shapesIn.SlideItemType == SlideItemType.Placeholder && shapesIn.TextBody.Paragraphs.ElementAt(0).TextParts.ElementAt(0).Font.FontSize == 20 && shapesIn.TextBody.Paragraphs.ElementAt(0).TextParts.ElementAt(0).Font.Color.ToArgb() != 13426152
            System.Drawing.Color colorLocRed = System.Drawing.ColorTranslator.FromHtml("#D61947");
            System.Drawing.Color colorLoc = System.Drawing.ColorTranslator.FromHtml("#CCDDE8");
            if (shapesIn.SlideItemType != SlideItemType.GroupShape && shapesIn.PlaceholderFormat.Type == PlaceholderType.Body  && !shapesIn.TextBody.Paragraphs.ElementAt(0).Font.Color.SystemColor.Equals(colorLoc) && !shapesIn.TextBody.Paragraphs.ElementAt(0).Font.Color.SystemColor.Equals(colorLocRed))
            {
                string testStringValue2 = shapesIn.TextBody.Text;
                System.Diagnostics.Debug.WriteLine(testStringValue2, "Text Slide: ");
                System.Diagnostics.Debug.WriteLine(shapesIn.TextBody.Paragraphs.ElementAt(0).Font.Color.SystemColor, "COLOR: ");
                System.Diagnostics.Debug.WriteLine(colorLoc, "COLOR: ");
            }
        }

        private async Task CreateComponentByType(ConfigSlideData SlideIn, String CurrSubDirIn, IProgress<ProgressReportModel> progress, CancellationToken cancellationToken, SlideReportDataModel output)
        {
            string AllNumOut = "";
            componentsToAdd.Clear();
            if(!SlideIn.IsResume && !SlideIn.IsObjective)
            {
                foreach (String CompT in SlideIn.ComponentTypes)
                {
                    switch (CompT)
                    {
                        case "BackGround":
                            GenerateBackGroundComp();
                            break;
                        case "Title":
                            GenerateTitleComp(SlideIn.TitleSlide, SlideIn.SubTitleSlide, SlideIn.ThirdSubTitleSlide, SlideIn.AllNumbers);
                            break;
                        case "Text":
                            GenerateTextComp(SlideIn.ContentTextData, "600", 100, 150, 800, componentsToAdd);
                            break;
                        case "Warpper":
                            GenerateWrapperComp(SlideIn);
                            break;
                        case "Button":
                            break;
                        case "LibraryPick":
                            //TODO: create library pick function 
                            break;
                        case "Ticker":
                            GenerateTickersComp(SlideIn);
                            break;
                        default:
                            break;
                    }
                }
            }
            else
            {
                if (SlideIn.IsObjective)
                    GenerateOBJRESStructure(SlideIn, "Objectifs", false);
                if(SlideIn.IsResume)
                    GenerateOBJRESStructure(SlideIn, "Résumé", true);
            }
            
            RootJson SlideInstanceJson = new RootJson();
            SlideInstanceJson.layout = new LayoutJson();
            
            SlideInstanceJson.components = componentsToAdd;

            if (SlideIn.AllNumbers != null)
                AllNumOut = SlideIn.AllNumbers;
            output.SlideAllNumber = AllNumOut;
            System.Diagnostics.Debug.WriteLine(componentsToAdd);
            await Task.Run( ()=> GenerateJsonFile(SlideInstanceJson, CurrSubDirIn, progress, cancellationToken, output));
            //GenerateJsonFile(SlideInstanceJson, CurrSubDirIn);

            //return AllNumOut;
        }

        private void GenerateOBJRESStructure(ConfigSlideData SI, string TitleIn, bool IsResumeIn)
        {
            GenerateBackGroundComp();
            GenerateLibraryPickComp();
            List<string> ResumeString = new List<string>();
            ResumeString.Add(TitleIn);
            GenerateTextComp(ResumeString, "600", 100, 30, 0, componentsToAdd, "#FFFFFF", "cuprum", "38");
            if (!IsResumeIn)
            {
                List<string> PlanString = new List<string>();
                PlanString.Add("Plan :");
                GenerateTextComp(PlanString, "600", 100, 200, 800, componentsToAdd, "#BE0000", "raleway", "30");
            }
            GenerateTextComp(SI.ContentTextData, "600", 100, 150, 800, componentsToAdd, "#585757", "raleway", "14", "#BE0000", "#BE0000");
        }
        private void GenerateBackGroundComp()
        {
            ComponentJson Newcomponent = new ComponentJson();
            Newcomponent.type = "ComponentBackgroundImage";
            Newcomponent.data = new DataJson();
            Newcomponent.data.type = "3D Paralax Image";
            Newcomponent.data.FilterOpacity = "0";
            Newcomponent.data.FilterColor = "#000000";
            Newcomponent.data.srcFilter = "1.png";
            Newcomponent.data.ScaleWidth = "0.5";
            Newcomponent.data.ScaleHeight = "0.5";
            Newcomponent.data.ParalaxBackground = "images/Image_0.jpg";
            Newcomponent.data.ParalaxForeground = "images/Image_1.png";
            Newcomponent.data.ParalaxMiddleground = "images/Image_2.png";
            componentsToAdd.Add(Newcomponent);
        }

        private void GenerateTitleComp(String L1, String L2, String L3, String Nm)
        {
            ComponentJson Newcomponent = new ComponentJson();
            Newcomponent.type = "ComponentTitle";
            Newcomponent.data = new DataJson();
            Newcomponent.data.level1 = L1;
            Newcomponent.data.level2 = L2;
            Newcomponent.data.level3 = L3;
            Newcomponent.data.number = Nm;
            componentsToAdd.Add(Newcomponent);
        }

        private ComponentJson GenerateTextComp(List<String> SldTD,string TxtW, double XPos, double YPos, int AnimTD, List<ComponentJson> listIn = null, string TextC = "#585757", string TextFontN = "raleway", string TextFS = "14", string SecondTextC = "#006999", string BulletColor = "#000000")
        {
            ComponentJson Newcomponent = new ComponentJson();
            Newcomponent.type = "ComponentTextPuceTextjs";
            List<double> PositionData = new List<double>();
            PositionData.Add(XPos);
            PositionData.Add(YPos);
            Newcomponent.position = PositionData;
            Newcomponent.data = new DataJson();
            Newcomponent.data.textWidth = TxtW;
            Newcomponent.data.textHeight = "10";
            Newcomponent.data.textColor = TextC;
            if(AnimTD == 0)
            {
                Newcomponent.data.animation = "No Animation";
            }
            else
            {
                Newcomponent.data.animation = "Animation";
            }
            
            Newcomponent.data.fontSize = TextFS;
            Newcomponent.data.fontName = TextFontN;
            Newcomponent.data.animationDelay = AnimTD;
            Newcomponent.data.textColorSymbole = "##";
            Newcomponent.data.bulletColor = BulletColor;
            Newcomponent.data.secondTextColor = SecondTextC;
            Newcomponent.data.content = SldTD;
            if(listIn != null)
                listIn.Add(Newcomponent);

            return Newcomponent;
        }

        private void GenerateLibraryPickComp()
        {
            ComponentJson Newcomponent = new ComponentJson();
            Newcomponent.type = "LibraryPick";
            List<double> PositionData = new List<double>();
            PositionData.Add(0);
            PositionData.Add(0);
            Newcomponent.position = PositionData;
            Newcomponent.symbol = "BarResume";
            Newcomponent.size = 1;

            componentsToAdd.Add(Newcomponent);
        }
        private void GenerateTickersComp(ConfigSlideData Sld)
        {
            ComponentJson Newcomponent = new ComponentJson();
            Newcomponent.type = "ComponentTickers";
            List<double> PositionData = new List<double>();
            PositionData.Add(26.5);
            PositionData.Add(0);
            Newcomponent.position = PositionData;
            Newcomponent.id = Sld.TickerID;
            Newcomponent.data = new DataJson();
            Newcomponent.data.clipName = " ";
            Newcomponent.data.textHeight = "10";
            Newcomponent.data.animation = "Animation";
            List<TabJson> TabsToAdd = GenerateTabs(Sld);
            Newcomponent.data.tabs = TabsToAdd;

            componentsToAdd.Add(Newcomponent);
        }

        private List<TabJson> GenerateTabs(ConfigSlideData Sld)
        {
            List<TabJson> TabsToAdd = new List<TabJson>();
            foreach (ConfigTabData tabS in Sld.SlideTicker.TabList)
            {
                TabJson SingleTab = new TabJson();
                SingleTab.titre = tabS.TabTitle;
                SingleTab.couleur = "#568974";
                SingleTab.ecrans = GenerateEcrans(tabS.EcranList);
                TabsToAdd.Add(SingleTab);
            }
            return TabsToAdd;
        }

        private List<EcranJson> GenerateEcrans(List<ConfigEcranData> SEd)
        {
            List<EcranJson> EcransToAdd = new List<EcranJson>();
            
            foreach (ConfigEcranData EcranS in SEd)
            {
                List<ComponentJson> componentsEcranToAdd = new List<ComponentJson>();
                EcranJson SingleEcran = new EcranJson();
                ComponentJson Comp = GenerateTextComp(EcranS.EContentTextData, "500", -200, -50, 100);
                componentsEcranToAdd.Add(Comp);
                SingleEcran.components = componentsEcranToAdd;
                EcransToAdd.Add(SingleEcran);
            }
            return EcransToAdd;
        }

        private void GenerateWrapperComp(ConfigSlideData Sld)
        {
            ComponentJson Newcomponent = new ComponentJson();
            Newcomponent.type = "ComponentWrapper";
            List<double> PositionData = new List<double>();
            PositionData.Add(50);
            PositionData.Add(400);
            Newcomponent.position = PositionData;
            Newcomponent.data = new DataJson();
            Newcomponent.data.offset = 60;
            Newcomponent.data.wrapperType = "horizontale";
            List<ComponentJson> CompToAdd = GenerateWrapperChildernComp(Sld);
            Newcomponent.data.components = CompToAdd;

            componentsToAdd.Add(Newcomponent);
        }

        //create class button contain button title button id button tab and ecran id
        //add that class to ConfigSlidedata as a list
        //add wrapperCompList variable to ConfigSlidedata
        private List<ComponentJson> GenerateWrapperChildernComp(ConfigSlideData Sld)
        {
            List<ComponentJson> componentsChildrenToAdd = new List<ComponentJson>();
            foreach (String CompT in Sld.WrapperCompTypes)
            {
                switch (CompT)
                {
                    case "Text":
                        GenerateTextComp(Sld.ContentTextData, "600", 100, 150, 800, componentsChildrenToAdd);
                        break;
                    case "Button":
                        GenerateButtonComp(Sld, componentsChildrenToAdd);
                        break;
                    case "LibraryPick":
                        break;
                    default:
                        break;
                }
            }
            return componentsChildrenToAdd;
        }

        private void GenerateButtonComp(ConfigSlideData Sld, List<ComponentJson> listIn = null)
        {
            bool MasterSizeSet = false;
            foreach (ConfigButtonData buttonInst in Sld.ButtonList)
            {
                ComponentJson Newcomponent = new ComponentJson();
                Newcomponent.type = "ComponentButton";
                List<double> PositionData = new List<double>();
                PositionData.Add(0);
                PositionData.Add(0);
                Newcomponent.position = PositionData;
                Newcomponent.data = new DataJson();
                Newcomponent.data.animation = "Animation";
                Newcomponent.data.animationDelay = 1000;
                Newcomponent.data.manualsize = false;
                Newcomponent.data.useAsButton = true;
                Newcomponent.data.rectangleWidth = 40;
                Newcomponent.data.rectangleheight = 45;
                Newcomponent.data.cornerRadius = 3;
                Newcomponent.data.textBtnWidth = "150";
                Newcomponent.data.textColor = "white";
                Newcomponent.data.rectangleColor = "#228e9bc1";
                Newcomponent.data.addArrow = false;
                Newcomponent.data.useMasterMaxSize = !MasterSizeSet;
                MasterSizeSet = true;
                Newcomponent.data.textToAdd = buttonInst.ButtonTitle;
                EventsJson NewEvent = new EventsJson();
                ClickJson NewClick = new ClickJson();
                NewClick.eventName = "AtsButtonClick";
                List<string> IdTickerData = new List<string>();
                IdTickerData.Add(buttonInst.ButtonID.ToString());
                NewClick.triggerComponentsId = IdTickerData;
                EventDataJson NewEventData = new EventDataJson();
                NewEventData.action = "open";
                NewEventData.tab = buttonInst.ButtonTab;
                NewEventData.ecran = buttonInst.ButtonEcran;
                NewClick.eventData = NewEventData;
                NewEvent.click = NewClick;
                Newcomponent.events = NewEvent;
                if (listIn != null)
                    listIn.Add(Newcomponent);
            }
        }
        private void GenerateJsonFile(RootJson JsonIn, String JsonPath, IProgress<ProgressReportModel> progress, CancellationToken cancellationToken, SlideReportDataModel output)
        {
            //List<SlideReportDataModel> output = new List<SlideReportDataModel>();
            ProgressReportModel report = new ProgressReportModel();
            String JsonPathToAdd = JsonPath + "/" + "Components.json";
            
            if (!Directory.Exists(JsonPath))
            {
                DuplicatedFolderNum = 0;
                Directory.CreateDirectory(JsonPath);
                Dispatcher.BeginInvoke(new ThreadStart(() => ZipFile.ExtractToDirectory(ZipFilePath.Text, JsonPath)));
                //await Task.Run(() => ZipFile.ExtractToDirectory(ZipFilePath.Text, JsonPath));
                //ZipFile.ExtractToDirectory(ZipFilePath.Text, JsonPath);
                //await Task.Factory.StartNew(() => ZipFile.ExtractToDirectory(ZipFilePath.Text, JsonPath));
                //DuplicatedFolderNum = 0;
            }
            else
            {
                DuplicatedFolderNum++;
                JsonPath = JsonPath + "-" + DuplicatedFolderNum;
                JsonPathToAdd = JsonPath + "/" + "Components.json";
                Directory.CreateDirectory(JsonPath);
                Dispatcher.BeginInvoke(new ThreadStart(() => ZipFile.ExtractToDirectory(ZipFilePath.Text, JsonPath)));
                //ZipFile.ExtractToDirectory(ZipFilePath.Text, JsonPath);
                //await Task.Run(() => ZipFile.ExtractToDirectory(ZipFilePath.Text, JsonPath));
                //await Task.Factory.StartNew(() => ZipFile.ExtractToDirectory(ZipFilePath.Text, JsonPath));
            }
            //TODO: Space in path will brake the code
            using (StreamWriter file = File.CreateText(JsonPathToAdd))
            {
                JsonSerializer serializer = new JsonSerializer
                {
                    NullValueHandling = NullValueHandling.Ignore
                };
                //serialize object directly into file stream
                serializer.Serialize(file, JsonIn);
            }
            SlideModeloutput.Add(output);
            CurrSlideIndex++;
            report.PercentageComplete = (CurrSlideIndex * 100) / CurrSlideCount;
            report.CreatedSlide = SlideModeloutput;
            //var PercentageComplete = (CurrSlideIndex * 100) / CurrSlideCount;
            //progress.Report(PercentageComplete);
            progress.Report(report);
            cancellationToken.ThrowIfCancellationRequested();
            Thread.Sleep(300);
            //System.Diagnostics.Debug.WriteLine(PercentageComplete);
        }

        private void SelectZipButton(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();



            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".zip";
            dlg.Filter = "Zip Files (*.zip)|*.zip|Rar Files (*.rar)|*.rar";


            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();


            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                string filename = dlg.FileName;
                ZipFilePath.Text = filename;
            }
        }

        private void SelectFolderButton(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = "C:/Users/aymen.menjli/Documents/Aymen menjli/TestBed/TestAppPPTX";
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                FolderPathPPTX.Text = dialog.FileName;
                //MessageBox.Show("You selected: " + dialog.FileName);
            }
        }

        private void cancelOperation_Click(object sender, RoutedEventArgs e)
        {
            cts.Cancel();
        }

        private void SetupUnhandledExceptionHandling()
        {
            // Catch exceptions from all threads in the AppDomain.
            AppDomain.CurrentDomain.UnhandledException += (sender, args) =>
                ShowUnhandledException(args.ExceptionObject as Exception, "AppDomain.CurrentDomain.UnhandledException", false);

            // Catch exceptions from each AppDomain that uses a task scheduler for async operations.
            TaskScheduler.UnobservedTaskException += (sender, args) =>
                ShowUnhandledException(args.Exception, "TaskScheduler.UnobservedTaskException", false);

            // Catch exceptions from a single specific UI dispatcher thread.
            Dispatcher.UnhandledException += (sender, args) =>
            {
                // If we are debugging, let Visual Studio handle the exception and take us to the code that threw it.
                if (!Debugger.IsAttached)
                {
                    args.Handled = true;
                    ShowUnhandledException(args.Exception, "Dispatcher.UnhandledException", true);
                }
            };

            // Catch exceptions from the main UI dispatcher thread.
            // Typically we only need to catch this OR the Dispatcher.UnhandledException.
            // Handling both can result in the exception getting handled twice.
            //Application.Current.DispatcherUnhandledException += (sender, args) =>
            //{
            //	// If we are debugging, let Visual Studio handle the exception and take us to the code that threw it.
            //	if (!Debugger.IsAttached)
            //	{
            //		args.Handled = true;
            //		ShowUnhandledException(args.Exception, "Application.Current.DispatcherUnhandledException", true);
            //	}
            //};
        }

        void ShowUnhandledException(Exception e, string unhandledExceptionType, bool promptUserForShutdown)
        {
            var messageBoxTitle = $"Unexpected Error Occurred: {unhandledExceptionType}";
            var messageBoxMessage = $"The following exception occurred:\n\n{e}";
            var messageBoxButtons = MessageBoxButton.OK;

            if (promptUserForShutdown)
            {
                messageBoxMessage += "\n\nNormally the app would die now. Should we let it die?";
                messageBoxButtons = MessageBoxButton.YesNo;
            }

            // Let the user decide if the app should die or not (if applicable).
            if (System.Windows.MessageBox.Show(messageBoxMessage, messageBoxTitle, messageBoxButtons) == MessageBoxResult.Yes)
            {
                //System.Net.Mime.MediaTypeNames.Application.Current.Shutdown();
                Environment.Exit(0);
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.PowerPoint;
using System.Data;
using DataTable = System.Data.DataTable;
using System.Data.Common;
using System.Diagnostics;
using Microsoft.Office.Core;
using tutor_testing_v4._1.TutorDataSetTableAdapters;
using MsoTriState = Microsoft.Office.Core.MsoTriState;
using tutor_testing_v4._1;
using System.Threading;

namespace ImagePathWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {
        public static Microsoft.Office.Interop.PowerPoint.Application application = new Microsoft.Office.Interop.PowerPoint.Application();
        public static Presentations ppPresens = application.Presentations;
        public static Presentation objPres = ppPresens.Open(AppDomain.CurrentDomain.BaseDirectory + "\\better powerpoint test v2.pptm", MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);
        public static Slides objSlides = objPres.Slides;
        public static SlideShowSettings objSSS = objPres.SlideShowSettings;
        public static SlideShowWindow objSSW;
        public static TutorDataSet.AllTutorsDataTable tutorTable = new TutorDataSet.AllTutorsDataTable();
        public static TutorDataSet.ScheduleDataTable scheduleTable = new TutorDataSet.ScheduleDataTable();
        public static TutorDataSet.SubjectDataTable subjectTable = new TutorDataSet.SubjectDataTable();
        public static int tutorsSlide = 1;
        public static bool canDelete = true;
        public static Image imgFromOrigin1;
        public MainWindow() {
            InitializeComponent();
            imgFromOrigin1 = imgFromOrigin;
            DisplayImage("2.jpg");
            Thread.Sleep(10000);
            //DisplayImage("1.jpg");
        }
        static void Init() {
            TutorDataSet db = new TutorDataSet();
            db.Clear();
            application.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            objSSS.Run();
            AllTutorsTableAdapter tutorTableAdapt = new AllTutorsTableAdapter();
            tutorTableAdapt.Fill(tutorTable);
            ScheduleTableAdapter scheduleAdapt = new ScheduleTableAdapter();
            scheduleAdapt.Fill(scheduleTable);
            SubjectTableAdapter subjectAdapt = new SubjectTableAdapter();
            subjectAdapt.Fill(subjectTable);
            objSSW = objPres.SlideShowWindow;


        }
        static void MainLoop() {
            //for(int i = 0; i < 100; i++) {
            //if (objSSW.View.Slide.SlideIndex == 4 && canDelete) {


            //canDelete = false;  // so to run only once
            //CreateSlide(tutorsSlide);
            //DeleteSlides();
            //Thread.Sleep(1000);
            //CreateSlide(tutorsSlide);
            //Thread.Sleep(10000);
            //DeleteSlides();
            //CreateSlide(tutorsSlide);
            //DeleteSlides();
            //canDelete = false;  // so to run only once

            //Thread.Sleep(1000);
            //} else if (objSSW.View.Slide.SlideIndex == 5) {
            canDelete = true;   // turn on 
                                // DisplayTutors();
                                //}
                                //}
        }
        internal static dynamic CurrentSlide {
            get {
                if (application.Active == MsoTriState.msoTrue &&
                    application.ActiveWindow.Panes[2].Active == MsoTriState.msoTrue) {
                    return application.ActiveWindow.View.Slide.SlideIndex;
                }
                return null;
            }
        }
        static void DisplayTutors() {
            DateTime currentDayTime = DateTime.Now;
            var query =
                from tutor in tutorTable.AsEnumerable()
                join schedule in scheduleTable
                on tutor.Field<int>("ID") equals schedule.Field<int>("ID")
                where schedule.Field<int>("Day") == (int)currentDayTime.DayOfWeek + 1 &&
                schedule.Field<DateTime>("Start").TimeOfDay <= currentDayTime.TimeOfDay &&
                schedule.Field<DateTime>("End").TimeOfDay >= currentDayTime.TimeOfDay
                select new {
                    TutorID = tutor.Field<int>("ID"),
                    Name = tutor.Field<string>("FirstName") + " " + tutor.Field<string>("LastName")
                };
            int i = 0;
            foreach (var q in query) {
                SlideRange slide = CreateSlide(tutorsSlide);
                WriteToTextbox(slide, "TutorName", q.Name + i);
                i++;
            }
        }
        static SlideRange CreateSlide(int copyOfIndex) {
            SlideRange newSlide = objSlides[copyOfIndex].Duplicate();
            newSlide.SlideShowTransition.Hidden = MsoTriState.msoFalse;
            newSlide.Tags.Add("isCreated", "true");
            //newSlide.MoveTo(objSlides.Count);
            return newSlide;
        }
        static string WriteToTextbox(SlideRange slide, string textboxName, string inputString) {
            slide.Shapes[textboxName].TextFrame.TextRange.Text = inputString;
            return inputString;
        }
        static int DeleteSlides() {
            int numberDeleted = 0;
            while (objSlides[objSlides.Count].Tags["isCreated"] == "true") {
                numberDeleted++;
                objSlides[objSlides.Count].Delete();
            }
            return numberDeleted;
        }
        static void DisplayImage(string fileName) {
            string path = @"\Images\" + fileName;
            imgFromOrigin1.Source = new BitmapImage(new Uri(path, UriKind.Relative));

        }
    }
};


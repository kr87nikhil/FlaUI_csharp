namespace FlaUiDemo {
    using System;
    using System.Drawing;
    using System.Diagnostics;
    using System.Threading;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using FlaUI.Core;
    using FlaUI.UIA3;
    using FlaUI.Core.Input;
    using FlaUI.Core.WindowsAPI;
    using FlaUI.Core.Capturing;
    using FlaUI.Core.AutomationElements;


    [TestClass]
    public class DoWordProcessing {
        private Application _application;
        private Window _mainWindow;
        
        [TestInitialize]
        public void Init()
        {
            //var application = Application.LaunchStoreApp("Microsoft.Office.WINWORD.EXE.15");
            ProcessStartInfo processStartInfo = new ProcessStartInfo(@"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE");
            _application = Application.AttachOrLaunch(processStartInfo);
            Thread.Sleep(3000);
            _mainWindow = _application.GetMainWindow(new UIA3Automation(), TimeSpan.FromSeconds(5));
        }

        [TestMethod]
        public void OpenBlankDocument()
        {
            var backstageView = _mainWindow.FindFirstDescendant(cf => cf.ByClassName("NetUIFullpageUIWindow"))
                ?.FindFirstChild(cf => cf.ByAutomationId("BackstageView"));
            var blankDocument = backstageView.FindFirstByXPath("//Pane[@Name='Home']/Group[@Name='New']//ListItem[@Name='Blank document']");
            blankDocument.Click();
            Thread.Sleep(2000);
        }

        [TestMethod]
        public void InsertTextAndCreateTable()
        {
            _mainWindow.FindFirstChild(cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.TitleBar)).
                FindFirstChild(cf => cf.ByAutomationId("Maximize-Restore"))?.AsButton().Click();
            //TODO: Element selection to be refined
            EnterText(_mainWindow);
            Thread.Sleep(2000);
            InsertTableAndApplyStyle(_mainWindow);
            Thread.Sleep(2000);
            EnterHeadingAndBodyInTable();
            Thread.Sleep(5000);
            TakeScreenShot();
        }

        [TestCleanup]
        public void CleanUp()
        {
            _application.Close();
        }

        /// <summary>
        /// Insert text into word document
        /// </summary>
        /// <param name="mainWindow"></param>
        public void EnterText(Window mainWindow) {
            mainWindow.FindFirstDescendant(cf => cf.ByAutomationId("AIOStartDocument"))?.AsListBoxItem().Click();
            Thread.Sleep(2000);
            Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_B);
            Thread.Sleep(2000);
            Keyboard.Type("Windows Automation");
            Thread.Sleep(1000);
            Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_E);
            Keyboard.Press(VirtualKeyShort.ENTER);
            Keyboard.Press(VirtualKeyShort.ENTER);
        }

        /// <summary>
        /// Insert table and apply style to it
        /// </summary>
        /// <param name="mainWindow"></param>
        public void InsertTableAndApplyStyle(Window mainWindow) {
            mainWindow.FindFirstDescendant(cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.TabItem)
                                                .And(cf.ByName("Insert")))?.AsTabItem().Click();
            mainWindow.FindFirstDescendant(cf => cf.ByName("Table"))?.AsMenuItem().Click();
            Point point = new Point(137, 180);
            Mouse.Position = point;
            point = new Point(175, 218);
            Mouse.MoveTo(point);
            Thread.Sleep(500);
            Mouse.LeftClick();
            mainWindow.FindFirstDescendant(cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Button)
                                                .And(cf.ByName("Table Styles")))?.AsButton().Click();
            Thread.Sleep(1000);
            point = new Point(710, 506);
            Mouse.Click(point);
        }

        /// <summary>
        /// Enter data into table
        /// </summary>
        public void EnterHeadingAndBodyInTable() {
            InsertTextAtCoordinates(500, 313, "Items");
            InsertTextAtCoordinates(500, 332, "MacBook");
            InsertTextAtCoordinates(500, 350, "Electric Bike");
            InsertTextAtCoordinates(700, 313, "Quantity");
            InsertTextAtCoordinates(900, 313, "Month");

            Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_B);
            Thread.Sleep(2000);
            InsertTextAtCoordinates(700, 332, "4");
            InsertTextAtCoordinates(900, 332, "June");
            InsertTextAtCoordinates(700, 350, "2");
            InsertTextAtCoordinates(900, 350, "March");
        }

        /// <summary>
        /// Insert text at specified coordinates
        /// </summary>
        /// <param name="xPosition"></param>
        /// <param name="yPosition"></param>
        /// <param name="text"></param>
        public void InsertTextAtCoordinates(int xPosition, int yPosition, string text) {
            var point = new Point(xPosition, yPosition);
            Mouse.Click(point);
            Keyboard.Type(text);
        }

        /// <summary>
        /// Capture screenshot and capture in file
        /// </summary>
        public void TakeScreenShot() {
            //Full screen
            var fullscreenImg = Capture.Screen();
            fullscreenImg.ToFile(@"D:\imgs\Full Screen.png");
            //var loginImg = Capture.Element(loginBtn);
            //var rectangleImg = Capture.Rectangle(new Rectangle(500, 500, 100, 150));
        }
    }
}

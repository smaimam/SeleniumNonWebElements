using AutoIt;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using SikuliModule;
using System;
using System.Diagnostics;
using System.Threading;
using WindowsInput;
using WindowsInput.Native;

namespace UnitTestProject8
{
    [TestClass]
    public class NonHtmlControl
    {
        public TestContext instance;
        public TestContext TestContext
        {
            set { instance = value; }
            get { return instance; }
        }

        #region Input Simulator


        [TestMethod]
        public void InputSimulatorKeyPressAndTextEntry()
        {
            IWebDriver driver = new ChromeDriver();
            driver.Url = "https://adactinhotelapp.com/";

            InputSimulator sim = new InputSimulator();

            Thread.Sleep(1000);
            sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
            Thread.Sleep(1000);
            sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
            Thread.Sleep(1000);
            sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
            Thread.Sleep(1000);

            sim.Keyboard.TextEntry("AmirTester");
            Thread.Sleep(1000);

            sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
            Thread.Sleep(1000);

            sim.Keyboard.TextEntry("AmirTester");
            Thread.Sleep(1000);

            sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            Thread.Sleep(5000);
            driver.Close();
        }


        [TestMethod]
        public void InputSimulator()
        {
            IWebDriver driver = new ChromeDriver();
            driver.Url = "https://adactinhotelapp.com/";
            Thread.Sleep(1000);
            InputSimulator sim = new InputSimulator();

            Thread.Sleep(3000);
            //sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
            //Thread.Sleep(1000);
            //sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
            //Thread.Sleep(1000);
            //sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
            //Thread.Sleep(1000);

            // sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_T);
            Thread.Sleep(3000);
            // enter username: QAUser01 
            sim.Keyboard.TextEntry("AmirTester");
            Thread.Sleep(1000);
            // press Tab key
            sim.Keyboard.KeyPress(VirtualKeyCode.TAB);
            Thread.Sleep(1000);
            // Enter Password 
            sim.Keyboard.TextEntry("AmirTester");
            Thread.Sleep(1000);
            // submit enter 
            sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
            Thread.Sleep(5000);
            driver.Close();
        }

        /// <summary>
        /// Simulate typing of any text as you do when you write.
        /// </summary>
        /// <param name="Text">Text that will be written automatically by simulation.</param>
        /// <param name="typingDelay">Time in ms to wait after 1 character is written.</param>
        /// <param name="startDelay"></param>
        private void simulateTypingText(string Text, int typingDelay = 100, int startDelay = 0)
        {
            InputSimulator sim = new InputSimulator();

            // Wait the start delay time
            sim.Keyboard.Sleep(startDelay);

            // Split the text in lines in case it has
            string[] lines = Text.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);

            // Some flags to calculate the percentage
            int maximum = lines.Length; int current = 1;

            foreach (string line in lines)
            {
                // Split line into characters
                char[] words = line.ToCharArray();

                // Simulate typing of the char i.e: a, e , i ,o ,u etc  // Apply immediately the typing delay
                foreach (char word in words)
                {
                    sim.Keyboard.TextEntry(word).Sleep(typingDelay);
                }

                float percentage = ((float)current / (float)maximum) * 100;

                current++;

                // Add a new line by pressing ENTER
                // Return to start of the line in your editor with HOME
                sim.Keyboard.KeyPress(VirtualKeyCode.RETURN);
                sim.Keyboard.KeyPress(VirtualKeyCode.HOME);

                // Show the percentage in the console
                Console.WriteLine("Percent : {0}", percentage.ToString());
            }
        }

        [TestMethod]
        public void InputSimulatorTextTypingByCharaters()
        {
            string demoText = "This is Demo Text";

            // Open a Notepad 
            Process.Start("notepad.exe");
            Thread.Sleep(1000);

            // Simulate typing text of a textbox multiline
            simulateTypingText(demoText);

            // Simulate typing slowly by waiting half second after typing every character
            simulateTypingText(demoText, 500);

            // Simulate typing slowly by waiting half second after typing every character
            // and wait 5 seconds before starting
            simulateTypingText(demoText, 500, 5000);
        }

        #endregion

        #region Auto IT
        [TestMethod]
        public void AutoITWithSelenium()
        {
            IWebDriver driver = new ChromeDriver();
            driver.Url = "https://adactinhotelapp.com/";
            Thread.Sleep(1000);

            AutoItX.WinWait("https://adactinhotelapp.com/", "", 2);

            AutoItX.WinActivate("https://adactinhotelapp.com/");
            AutoItX.Send("{TAB 3}", 0); Thread.Sleep(1000);
            AutoItX.Send("Amir");
            Thread.Sleep(1000);
            AutoItX.Send("{TAB}", 0);
            AutoItX.Send("Imam");
            Thread.Sleep(1000);
            AutoItX.Send("{Enter}", 0);

            driver.Close();
        }

        [TestMethod]
        public void AutoIT_NotepadSave()
        {
            AutoItX.Run("notepad.exe", "");
            AutoItX.WinWaitActive("Untitled - Notepad");
            AutoItX.Send("I'm in notepad");
            AutoItX.WinWaitActive("*Untitled - Notepad");

            AutoItX.Send("{ALTDOWN}+{F4}", 0);
            AutoItX.Send("{ALTUP}", 0);
            AutoItX.Send("{ENTER}", 0);
            AutoItX.Send(@"C:\test2.txt");
            AutoItX.Send("{ENTER}", 0);
        }

        [TestMethod]
        public void AutoIT_NotepadClose()
        {
            AutoItX.Run("notepad.exe", "");
            AutoItX.WinWaitActive("Untitled - Notepad");
            AutoItX.Send("I'm in notepad");
            AutoItX.WinWaitActive("*Untitled - Notepad");
            IntPtr winHandle = AutoItX.WinGetHandle("*Untitled - Notepad");
            AutoItX.WinKill(winHandle);
        }

        [TestMethod]
        public void AutoITClick()
        {
            AutoItX.Run("notepad.exe", "");
            AutoItX.WinWaitActive("Untitled - Notepad");
            AutoItX.Send("I'm in notepad");
            AutoItX.MouseClick();
        }

        #endregion

        #region Sikuli Integrator

        [TestMethod]
        public void SikuliIntegratorWithSelenium()
        {
            IWebDriver driver = new ChromeDriver();

            driver.Url = "https://adactinhotelapp.com/";
            Thread.Sleep(1000);

            SikuliAction.Click(@"\img\loginBtn.png");
            Thread.Sleep(1000);

            driver.Close();

        }

        [TestMethod]
        public void SikuliIntegratorInputSimulator()
        {
            IWebDriver driver = new ChromeDriver();             // Initialize Driver
            InputSimulator sim = new InputSimulator();          // Initialize InputSimulator

            driver.Url = "https://adactinhotelapp.com/";        // Open Url

            SikuliAction.Click(@"\img\username.png");   // Click on Username Field          
            sim.Keyboard.TextEntry("AmirTester");       // Set value on Username Field
            SikuliAction.Click(@"\img\password.png");   // Click on Password Field
            sim.Keyboard.TextEntry("AmirTester");       // Set value on Password Field
            SikuliAction.Click(@"\img\loginBtn.png");   // Click on Login Button

            driver.Close();                             // Close Driver
        }

        [TestMethod]
        public void SikuliIntegratorAutoIt()
        {
            IWebDriver driver = new ChromeDriver();                 // Initialize Driver                        

            driver.Url = "https://adactinhotelapp.com/";            // Open Url

            SikuliAction.Click(@"\img\username.png");      // Click on Username Field        
            AutoItX.Send("AmirTester");                             // Set value on Username Field
            SikuliAction.Click(@"\img\Sikuli\password.png");       // Click on Password Field
            AutoItX.Send("AmirTester");                             // Set value on Password Field
            SikuliAction.Click(@"\img\Sikuli\loginBtn.png");       // Click on Login Button

            driver.Close();                                         // Close Driver
        }

        [TestMethod]
        public void SikuliIntegratorMethods()
        {
            string MyPicture = @"\img\loginBtn.png";
            string MyPicture2 = @"\img\username.png";
            SikuliAction.Click(MyPicture);
            SikuliAction.Click(MyPicture);
            SikuliAction.DoubleClick(MyPicture);
            SikuliAction.RightClick(MyPicture);
            SikuliAction.DragAndDrop(MyPicture, MyPicture2);
            SikuliAction.Hover(MyPicture);
            SikuliAction.Equals(MyPicture, MyPicture2);
            SikuliAction.Exists(MyPicture);
            SikuliAction.Wait(MyPicture, 5000);
            SikuliAction.WaitVanish(MyPicture, 5000);
        }

        [TestMethod]
        public void SikuliIntegratorOverloadMethods()
        {
            string MyPicture = @"\img\loginBtn.png";
            string MyPicture2 = @"\img\username.png";
            SikuliAction.Click(MyPicture);
            SikuliAction.Click(MyPicture, 90, 5000);

            SikuliAction.DoubleClick(MyPicture);
            SikuliAction.DoubleClick(MyPicture, 90, 5000);

            SikuliAction.RightClick(MyPicture);
            SikuliAction.RightClick(MyPicture, 90, 5000);

            SikuliAction.DragAndDrop(MyPicture, MyPicture2);
            SikuliAction.DragAndDrop(MyPicture, MyPicture2, 90, 5000);

            SikuliAction.Hover(MyPicture);
            SikuliAction.Hover(MyPicture, 90, 5000);

            SikuliAction.Equals(MyPicture, MyPicture2);

            SikuliAction.Exists(MyPicture);
            SikuliAction.Exists(MyPicture, 90, 5000);

            SikuliAction.Wait(MyPicture, 5000);
            SikuliAction.Wait(MyPicture, 90, 5000);

            SikuliAction.WaitVanish(MyPicture, 5000);
            SikuliAction.WaitVanish(MyPicture, 90, 5000);
        }

        #endregion

        #region IJavaScriptExecutor       

        [TestMethod]
        public void ExecuteJavaScriptCode(IWebDriver driver, string JSCode)
        {
            IJavaScriptExecutor executor = (IJavaScriptExecutor)driver;
            executor.ExecuteScript(JSCode);
        }

        [TestMethod]
        public void ScrollToElement()
        {
            IWebDriver driver = new ChromeDriver();
            driver.Url = "https://adactinhotelapp.com/";
            var element = driver.FindElement(By.Id("username"));

            //Scroll To Element
            ((IJavaScriptExecutor)driver)
                .ExecuteScript("arguments[0].scrollIntoView(true);", element);
        }

        [TestMethod]
        public void ScrollToTop()
        {
            IWebDriver driver = new ChromeDriver();
            driver.Url = "https://adactinhotelapp.com/";
            var element = driver.FindElement(By.Id("username"));

            // Scroll Top or Bottom
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("window.scrollTo(0, document." + "head" + ".scrollHeight);");
        }

        [TestMethod]
        public void ScrollToBottom()
        {
            IWebDriver driver = new ChromeDriver();
            driver.Url = "https://adactinhotelapp.com/";
            var element = driver.FindElement(By.Id("username"));

            // Scroll Top or Bottom
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("window.scrollTo(0, document." + "Body" + ".scrollHeight);");
        }

        [TestMethod]
        public void IJavaScriptExecutor_Click()
        {
            IWebDriver driver = new ChromeDriver();

            driver.Url = "https://adactinhotelapp.com/";
            var element = driver.FindElement(By.Id("login"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", element);

            driver.Close();
        }

        [TestMethod]
        public void IJavaScriptExecutor_SendKeys()
        {
            IWebDriver driver = new ChromeDriver();
            driver.Url = "https://adactinhotelapp.com/";
            var element = driver.FindElement(By.Id("username"));

            //Enter Text Using Javascript
            IJavaScriptExecutor executor = (IJavaScriptExecutor)driver;
            executor.ExecuteScript("arguments[0].value ='" + "EnteredText" + "';", element);
        }

        [TestMethod]
        public void IJavaScriptExecutor_GetTitle()
        {
            IWebDriver driver = new ChromeDriver();
            driver.Url = "https://adactinhotelapp.com/";

            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            string title = (string)js.ExecuteScript("return document.title");
            Console.WriteLine(title);

            driver.Close();
        }

        [TestMethod]
        public void IJavaScriptExecutor_IsPageReady()
        {
            IWebDriver driver = new ChromeDriver();
            driver.Url = "https://adactinhotelapp.com/";

            bool pageState = ((IJavaScriptExecutor)driver)
                .ExecuteScript("return document.readyState")
                .Equals("complete");

            Console.WriteLine(pageState.ToString());

            driver.Close();
        }

        [TestMethod]
        public void IJavaScriptExecutor_GetElementState()
        {
            IWebDriver driver = new ChromeDriver();
            driver.Url = "https://adactinhotelapp.com/";

            bool elementState = (bool)((IJavaScriptExecutor)driver)
                .ExecuteScript("return document.getElementById('username').disabled");

            Console.WriteLine(elementState.ToString());

            driver.Close();
        }

        [TestMethod]
        public void IJavaScriptExecutor_PerformLoginUsingJQuery()
        {
            IWebDriver driver = new ChromeDriver();
            driver.Url = "https://adactinhotelapp.com/";

            #region load jQuery
            string loadjQuery = "(function() {" +
                " const script = document.createElement('script');" +
                "script.src = 'https://ajax.googleapis.com/ajax/libs/jquery/3.6.0/jquery.min.js';" +
                "script.type = 'text/javascript';" +
                "script.addEventListener('load', () => {" +
                " });document.head.appendChild(script);})();";
            #endregion

            ((IJavaScriptExecutor)driver).ExecuteScript(loadjQuery);

            Thread.Sleep(5000);
            ((IJavaScriptExecutor)driver).ExecuteScript("$('#username').val('AmirTester')");
            ((IJavaScriptExecutor)driver).ExecuteScript("$('#password').val('AmirTester')");
            ((IJavaScriptExecutor)driver).ExecuteScript("$('#login').click()");

            driver.Close();
        }

        [TestMethod]
        public void IJavaScriptExecutor_JQuery()
        {
            IWebDriver driver = new ChromeDriver();
            driver.Url = "https://demoqa.com/text-box";

            ((IJavaScriptExecutor)driver).ExecuteScript("$('#userName').va('Test User')");
            ((IJavaScriptExecutor)driver).ExecuteScript("$('#userEmail').val('Test@User.com')");
            ((IJavaScriptExecutor)driver).ExecuteScript("$('#currentAddress').val('Test/1, Demo Address')");
            ((IJavaScriptExecutor)driver).ExecuteScript("$('#permanentAddress').val('Test/2 Demo Address2')");
            ((IJavaScriptExecutor)driver).ExecuteScript("$('#submit').click()");

            driver.Close();
        }


        #endregion
    }
}

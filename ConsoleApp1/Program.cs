using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace SeleniumChromeSample
{
    class Program
    {
        /// <summary>
        /// 送られてきたURLをChomeで開く
        /// </summary>
        /// <param name="args"></param>
        /// 

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr GetForegroundWindow();


        [DllImport("user32.dll", EntryPoint = "GetWindowText", CharSet = CharSet.Auto)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);


        // powershellで実行するコマンドを作成
        const string ps_command = "Get-ChildItem . -include hogerihogeta.py -Recurse";

        static void Main(string[] args)
        {

            String url = "https://www.google.co.jp/";
            Program program = new Program();
            // program.opengeturlforchome(url);
            // program.getcurrenturl();
            // string currenturl =  program.getactivetaburl();

            //for(int k = 0; k < 100; k++)
            // {
            //    program.getactivefilepath();
            //    thread.sleep(100);
            //}
            // program.getactivefilepath();

            // program.movedirectoryoffile();
            program.execPowershell();

        }

        private void openGetUrlForChome(String url)
        {
            IWebDriver driver = new ChromeDriver();
            IWebElement textbox;
            IWebElement findbuttom;

            //Webページを開く
            driver.Navigate().GoToUrl(url);


            //検索ボックス
            textbox = driver.FindElement(By.Name("q"));
            //検索ボックスに検索ワードを入力
            textbox.SendKeys("Selenium");
            textbox.Submit();
            
            String currentUrl = driver.Url;
            Console.WriteLine(currentUrl);

        }

        private String GetCurrentUrl()
        {
            IWebDriver driver = new ChromeDriver();
            String currentUrl = driver.Url;
            Console.WriteLine(currentUrl);
            return currentUrl;
        }

        public String GetActiveTabUrl()
        {
            AutomationElement.RootElement
              .FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.ClassNameProperty, "Chrome_WidgetWin_1"))
              .SetFocus();
            SendKeys.SendWait("^l");
            var elmUrlBar = AutomationElement.FocusedElement;
            var valuePattern = (ValuePattern)elmUrlBar.GetCurrentPattern(ValuePattern.Pattern);
           
            Console.WriteLine(valuePattern.Current.Value);

            return valuePattern.Current.Value;
        }

        public void getActiveFilePath()
        {
            StringBuilder sb = new StringBuilder(65535);//65535に特に意味はない
            GetWindowText(GetForegroundWindow(), sb, 65535);
            Console.WriteLine(sb);
        }

        public void moveDirectoryOfFile(){

            string sourceFile = @"C:\Users\g1723035\Documents\R_D\sample_C#\ConsoleApp1\ConsoleApp1\test.py";
            string destinationFile = @"C:\Users\g1723035\Documents\test.py";

            // To move a file or folder to a new location:
            System.IO.File.Move(sourceFile, destinationFile);
        }

        public void getFile()
        {
            Console.WriteLine( System.IO.Directory.GetFiles("test", ".py", System.IO.SearchOption.AllDirectories));

        }
        //PowerShellの実行メソッド（引数:PowerShellコマンド)
        static void OpenWithArguments(string options)
        {
            Process cmd = new Process();
            cmd.StartInfo.FileName = "PowerShell.exe";
            //PowerShellのWindowを立ち上げずに実行。
            cmd.StartInfo.WindowStyle = ProcessWindowStyle.Hidden; 
            // 引数optionsをShellのコマンドとして渡す。
            cmd.StartInfo.Arguments = options;

            Console.Write("=======================================================================================");
            cmd.Start();
        }

        public void execPowershell(){

        string option = ps_command;
        Program.OpenWithArguments(option);
            
        }
    }

}
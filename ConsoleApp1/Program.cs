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
using System.IO.Ports;

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
            
             // いったんダミーファイルの場所をべた書きしておく
            string sourceFile = @"C:\Users\g1723035\Documents\R_D\sample_C#\ConsoleApp1\ConsoleApp1\test.py";
            // ファイルの移動先を示す
            string destinationFile = @"C:\Users\g1723035\Documents\test.py";
            
            String url = "https://www.google.co.jp/";
            
            Program program = new Program();

            var request =  Console.ReadLine();

            // 取得したURLをchromeで開く
            if(request == "open url") {
                program.openGetUrlForChome(url);
            }

            // いらないかな？
            if(request == "i") {
                program.GetCurrentUrl();
            }

            // 現在のchormeのアクティブなURLを取得し返却する
            if(request == "active url"){
                string currenturl =  program.GetActiveTabUrl();
            }
            
            // 現在のアクティブなアプリケーションの名前を取得する
            if (request == "active application and file name") {
                for(int k = 0; k < 100; k++)
                {
                    program.getActiveApplicationAndFileName();
                    Thread.Sleep(100);
                }
            }
            
            // ファイルを所定の場所に移動させる
            if(request == "move file"){
                program.moveDirectoryOfFile(sourceFile, destinationFile);
            }

            // powershellを起動できるらしいが多分動いていない
            if(request == "exec powershell") {
                program.execPowershell();
            }

        }

        /**
         * 取得したURLをもとにChomeで開く
         * @param String url
         */
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

        /**
         * 多分いらないので後で消す 
         */

        private String GetCurrentUrl()
        {
            IWebDriver driver = new ChromeDriver();
            String currentUrl = driver.Url;
            Console.WriteLine(currentUrl);
            return currentUrl;
        }

        /**
         * Chomeで開いているアクティブなタブを取得する
         * @param 
         * @return chome active tab url
         */
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

        /**
         * 現在アクティブになっているアプリケーションの名前とファイルの名前を取得する
         * @param
         * @return active application name and file name
         */
        public void getActiveApplicationAndFileName()
        {
            Console.WriteLine("==========================");
            StringBuilder sb = new StringBuilder(65535);//65535に特に意味はない
            GetWindowText(GetForegroundWindow(), sb, 65535);
            Console.WriteLine(sb);
        }

        /**
         * ファイルを所定の場所に移動する
         * @param sourceFile
         * @param destinationFile
         * 
         * @return
         */
        public void moveDirectoryOfFile(String sourceFile, String destinationFile){

            // To move a file or folder to a new location:
            System.IO.File.Move(sourceFile, destinationFile);
        }

        /**
         * 今動いてない
         * ファイル名からファイルパスを取得するためにpowershellを使用する予定だったが研究に不要なので削除するかもしれない
         */
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
using System;

namespace ReceiveData {
    public class ReceiveData
    {

        /**
            * 取得したURLをもとにChomeで開く
            * @param String url
            */
        public void openGetUrlForChome(String url)
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
       * ファイルを所定の場所に移動する
       * @param sourceFile
       * @param destinationFile
       * 
       * @return
       */
        public void moveDirectoryOfFile(String sourceFile, String destinationFile)
        {

            // To move a file or folder to a new location:
            System.IO.File.Move(sourceFile, destinationFile);
        }
    }

}

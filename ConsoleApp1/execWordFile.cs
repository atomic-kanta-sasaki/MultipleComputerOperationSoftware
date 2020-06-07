using System;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace ConsoleApp1
{
    class execWordFile
    {

        public void exec()
        {
            try
            {

                // Word アプリケーションオブジェクトを作成
                Word.Application word = new Word.Application();
                // Word の GUI を起動しないようにする
                word.Visible = false;

                // 新規文書を作成
                Document document = word.Documents.Add();

                // ヘッダーを編集
                editHeaderSample(ref document, 10, WdColorIndex.wdPink, "Header Area");

                // フッターを編集
                editFooterSample(ref document, 10, WdColorIndex.wdBlue, "Footer Area");

                // 見出しを追加
                addHeadingSample(ref document, "見出し");

                // パラグラフを追加
                document.Content.Paragraphs.Add();

                // テキストを追加
                addTextSample(ref document, WdColorIndex.wdGreen, "Hello, ");
                addTextSample(ref document, WdColorIndex.wdRed, "World");

                // 名前を付けて保存
                object filename = System.IO.Directory.GetCurrentDirectory() + @"\out.docx";
                document.SaveAs2(ref filename);

                // 文書を閉じる
                document.Close();
                document = null;
                word.Quit();
                word = null;

                Console.WriteLine("Document created successfully !");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        /// <summary>
        /// 文書のヘッダーを編集する.
        /// </summary>
        private static void editHeaderSample(ref Document document, int fontSize, WdColorIndex color, string text)
        {
            foreach (Section section in document.Sections)
            {
                //Get the header range and add the header details.
                Range headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                headerRange.Font.ColorIndex = color;
                headerRange.Font.Size = fontSize;
                headerRange.Text = text;
            }
        }

        /// <summary>
        /// 文書のフッターを編集する.
        /// </summary>
        private static void editFooterSample(ref Document document, int fontSize, WdColorIndex color, string text)
        {
            foreach (Section wordSection in document.Sections)
            {
                //Get the footer range and add the footer details.
                Range footerRange = wordSection.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Font.ColorIndex = color;
                footerRange.Font.Size = fontSize;
                footerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                footerRange.Text = text;
            }
        }

        /// <summary>
        /// 文書に見出しを追加する.
        /// </summary>
        private static void addHeadingSample(ref Document document, string text)
        {
            Paragraph para = document.Content.Paragraphs.Add(System.Reflection.Missing.Value);
            object styleHeading1 = "見出し 1";
            para.Range.set_Style(ref styleHeading1);
            para.Range.Text = text;
            para.Range.InsertParagraphAfter();
        }

        /// <summary>
        /// 文書の末尾位置を取得する.
        /// </summary>
        /// <returns></returns>
        private static int getLastPosition(ref Document document)
        {
            return document.Content.End - 1;
        }

        /// <summary>
        /// 文書の末尾にテキストを追加する.
        /// </summary>
        private static void addTextSample(ref Document document, WdColorIndex color, string text)
        {
            int before = getLastPosition(ref document);
            Range rng = document.Range(document.Content.End - 1, document.Content.End - 1);
            rng.Text += text;
            int after = getLastPosition(ref document);

            document.Range(before, after).Font.ColorIndex = color;
        }
    }
}

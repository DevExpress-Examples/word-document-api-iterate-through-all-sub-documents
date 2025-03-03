using System;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit;
using System.Drawing;
using System.Diagnostics;

namespace SubDocumentIterator
{
    class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            using (RichEditDocumentServer wordProcessor = new RichEditDocumentServer())
            {
                wordProcessor.LoadDocument("template.docx");
                Document document = wordProcessor.Document;
                AskAction(document);
                wordProcessor.SaveDocument("Modified.docx", DocumentFormat.OpenXml);
            }

            var p = new Process();
            p.StartInfo = new ProcessStartInfo(@"Modified.docx")
            {
                UseShellExecute = true
            };
            p.Start();

        }

        static private void AskAction(Document document)
        {
            Console.WriteLine("Enter a command's index.\r\nAvailable commands:\r\n1. Update Fields\r\n2. Remove Bookmarks\r\n3. Replace Text\r\n4. Highlight Text");
            var answer = Console.ReadLine();
            switch (answer)
            {
                case "1": { document.ForEachSubDocument((subdoc => { subdoc.Fields.Update(); })); break; }
                case "2":
                    {
                        document.ForEachSubDocument((subdoc =>
                        {
                            for (int i = subdoc.Bookmarks.Count - 1; i >= 0; i--)
                                subdoc.Bookmarks.Remove(subdoc.Bookmarks[i]);
                        }));
                        break;
                    }
                case "3": { document.ForEachSubDocument((subdoc => { subdoc.ReplaceAll("test text", "Hello!!!", SearchOptions.None); })); break; }
                case "4":
                    {
                        document.ForEachSubDocument((subdoc =>
                    {
                        DocumentRange[] ranges = subdoc.FindAll("time", SearchOptions.None);
                        foreach (DocumentRange range in ranges)
                        {
                            CharacterProperties cp = subdoc.BeginUpdateCharacters(range);
                            cp.ForeColor = Color.Red;
                            cp.BackColor = Color.Lavender;
                            subdoc.EndUpdateCharacters(cp);
                        }
                    }));
                        break;
                    }
            }
            Console.WriteLine("Do you want to perform another action? If no, the document is saved and opened in the Word. Y/N");
            var answerReply = Console.ReadLine()?.ToLower();
            if (answerReply == "y")
            {
                AskAction(document);
            }
        }
    }
}

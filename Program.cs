using System;
using Word = Microsoft.Office.Interop.Word;

namespace Testing
{
    internal class Program
    {
        static void Main(string[] args)
        { 
            string templateFile = @"D:\Project via\Demo\Sample.pdf";
            var wordApp = new Word.Application();
            var wordDoc = wordApp.Documents.Add(templateFile);
            

            Console.WriteLine("name:");
            string Name = Console.ReadLine();

            Console.WriteLine("Age:");
            string Age = Console.ReadLine();

            Console.WriteLine("Profesion:");
            string Prof = Console.ReadLine();

            Console.WriteLine("Where you study:");
            string Stud = Console.ReadLine();

            Console.WriteLine("Info by you:");
            string Info = Console.ReadLine();

            Console.WriteLine($"Name - {Name}, Age - {Age}, Profetsion - {Prof}, Study - {Stud}, Info by you - {Info},");

            Console.WriteLine("What format to save");
            string SaveFormat = Console.ReadLine();
            
            void ReplaceStub(string stubToReplace, string text, Word.Document worldDocument)
            {
                var range = worldDocument.Content;
                range.Find.ClearFormatting();
                object wdReplaceAll = Word.WdReplace.wdReplaceAll;
                range.Find.Execute(FindText: stubToReplace, ReplaceWith: text, Replace: wdReplaceAll);
            }

            try
            {
                ReplaceStub("{fio}", Name, wordDoc);
                ReplaceStub("{age}", Age, wordDoc);
                ReplaceStub("{prof}", Prof, wordDoc);
                ReplaceStub("{study}", Stud, wordDoc);
                ReplaceStub("{info}", Info, wordDoc);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

            switch (SaveFormat)
            {
                case "dox":
                    wordDoc.SaveAs(@"D:\Project via\Demo\Резюме" + "_" + Name);
                    wordDoc.Close();
                    break;
                case "pdf":
                    wordDoc.SaveAs(@"D:\Project via\Demo\Резюме" + "_" + Name, Word.WdSaveFormat.wdFormatPDF);
                    wordDoc.Close();
                    break;
                case "txt":
                    wordDoc.SaveAs(@"D:\Project via\Demo\Резюме" + "_" + Name, Word.WdSaveFormat.wdFormatText);
                    wordDoc.Close();
                    break;
                default:
                    Console.WriteLine("this format is not available");
                    break;
            }

        }

    }
}

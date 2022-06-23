using System;
using Word = Microsoft.Office.Interop.Word;

namespace Testing
{
    internal class Program
    {
        static void Main(string[] args)
        { 
            string templateFile = @"D:\Project via\Testing\Шаблон.docx";

            Console.WriteLine("name:");
            string Name = Console.ReadLine();
            Console.WriteLine($"Имя - {Name}");

            var wordApp = new Word.Application();// создаем новый экземпляр ворда
            //Функция для замены наших меток
            void ReplaceStub(string stubToReplace, string text, Word.Document worldDocument)
            {
                var range = worldDocument.Content;
                range.Find.ClearFormatting();
                object wdReplaceAll = Word.WdReplace.wdReplaceAll;
                range.Find.Execute(FindText: stubToReplace, ReplaceWith: text, Replace: wdReplaceAll);
            }

            try
            {
                var wordDoc = wordApp.Documents.Add(templateFile);

                ReplaceStub("{fio}", Name, wordDoc);//Заменяем метку на данные из формы
                ///Может быть много таких меток
                wordDoc.SaveAs(@"D:\Project via\Testing\Резюме" + "_" + Name);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
    }
}

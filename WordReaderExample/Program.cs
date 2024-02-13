using System.Text;
using WordReaderExample.Models;
using WordReaderExample.Services;

namespace WordReaderExample
{
    internal class Program
    {
        
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            string filePath = AppDomain.CurrentDomain.BaseDirectory + "test.docx";
            WordService wordService = new WordService();
             List<DomainObject> readedItems = wordService.ReadWordDocument(filePath);

            foreach(var item in readedItems)
            {
                Console.WriteLine(item.Name);
            }
        }
    }
}

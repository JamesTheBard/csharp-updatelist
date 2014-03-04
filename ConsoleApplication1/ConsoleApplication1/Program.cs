using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using WUApiLib;

namespace ConsoleApplication1
{
    class Program
    {
        static int tableWidth = 99;
        static int widthTwo = 70;
        static void Main(string[] args)
        {
            GetInstalledUpdates();
        }

        static void GetInstalledUpdates()
        {
            UpdateSessionClass uSession = new UpdateSessionClass();
            IUpdateSearcher uSearcher = uSession.CreateUpdateSearcher();
            ISearchResult uResult = uSearcher.Search("IsInstalled=1 and Type='Software'");
            List<IUpdate> results = new List<IUpdate>{ };
            foreach (IUpdate update in uResult.Updates)
            {
                results.Add(update);
            }
            results = results.OrderBy(x => Convert.ToInt32(x.KBArticleIDs[0]))
                             .ToList();
            PrintRow("KB Article", "Description", "Severity");
            PrintLine();

            foreach (IUpdate update in results)
            {
                PrintRow(update.KBArticleIDs[0], update.Title, update.MsrcSeverity);
            }
            PrintLine();
        }

        static void PrintLine()
        {
            Console.WriteLine(new string('-', tableWidth));
        }

        static void PrintRow(params string[] columns)
        {
            columns[0] = columns[0].PadLeft(10, ' ');
            columns[1] = columns[1].Length > widthTwo ? 
                columns[1].Substring(0, widthTwo - 3) + "..." : 
                columns[1].PadRight(widthTwo, ' ');
            string rowFormat = String.Format("{0, 10}  {1, -70}  {2, -10}", columns);
            Console.WriteLine(rowFormat);
        }

        static string AlignCenter(string text, int width)
        {
            text = text.Length > width ? text.Substring(0, width - 3) + "..." : text;
            if (string.IsNullOrEmpty(text))
            {
                return new string(' ', width);
            }
            else
            {
                return text.PadRight(width - (width - text.Length) / 2).PadLeft(width);
            }
        }
    }
}

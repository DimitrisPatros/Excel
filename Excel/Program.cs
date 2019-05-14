using System;

//npoi for excel files to interact with visual studio

namespace Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelInputOutput excel = new ExcelInputOutput();

            excel.FillWithDummyData();
            
            Console.WriteLine("give the name of the new file");
            string filename = Console.ReadLine();
            excel.Load(filename);
            // excel.Save(filename);
            excel.PrintList();
        }
    }
}

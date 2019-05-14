using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Excel
{
    public class ExcelInputOutput : IExcel
    {
        public List<Student> students = new List<Student>();

        public void FillWithDummyData()
        {
            students.Add(new Student("DImitris", "BCs Cs", 1));
            students.Add(new Student("Antonis", "MCs Cs", 2));
            students.Add(new Student("eugenia", "BCs ML", 3));
        }

        public bool Load(string ExcelFileName)
        {
            students.Clear();
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream
                ($@"C:\Users\patro\source\repos\Excel\Excel\bin\Debug\netcoreapp2.2\{ExcelFileName}.xlsx", FileMode.Open,FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }
            ISheet sheet = hssfwb.GetSheet("Sheet");
            for (int row = 1; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) != null)
                {
                   string name = sheet.GetRow(row).GetCell(0).StringCellValue;
                   string course = sheet.GetRow(row).GetCell(1).StringCellValue;
                   double reg = sheet.GetRow(row).GetCell(2).NumericCellValue;
                   int register2 = Convert.ToInt32(reg);
                   students.Add(new Student(name,course,register2));
                }
            }
            return true;
        }

        public bool Save(string ExcelFileName)
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet");
            //ISheet sheet2 = wb.CreateSheet("Sheet 02");
            int x = 0;
            var r1 =sheet.CreateRow(0);
            r1.CreateCell(0).SetCellValue("Name");
            r1.CreateCell(1).SetCellValue("Course");
            r1.CreateCell(2).SetCellValue("RegisterId");

            foreach (Student s in students)
            { x++;
                var r = sheet.CreateRow(x);
                r.CreateCell(0).SetCellValue(students[x - 1].Name);
                r.CreateCell(1).SetCellValue(students[x - 1].Course);
                r.CreateCell(2).SetCellValue(students[x - 1].RegisterId);
            }

            using (var fs = new FileStream($"{ExcelFileName}.xlsx", FileMode.Create,
            FileAccess.Write))
            {
                wb.Write(fs);
            }
            return true;
        }

        public void PrintList()
        {
            foreach (Student s in students)
            {
                Console.WriteLine(s.ToString());
                Console.WriteLine("---------------------");
            }
        }

    }

}

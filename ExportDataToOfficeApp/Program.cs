using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExportDataToOfficeApp
{
    public class User
    {
        public string Name { get; set; }
        public double WorkedDays { get; set; }
        public double DaysInOffice { get; set; }
    }
    class Program
    {
        static void ExportToExcel(List<User> usersInOffice)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Workbooks.Add();
            Excel._Worksheet workSheet = (Excel._Worksheet)excelApp.ActiveSheet;
            workSheet.Cells[1, "A"] = "Name";
            workSheet.Cells[1, "B"] = "Worked in office";
            workSheet.Cells[1, "C"] = "Days in office";

            int row = 1;
            foreach (User c in usersInOffice)
            {
                row++;
                workSheet.Cells[row, "A"] = c.Name;
                workSheet.Cells[row, "B"] = c.WorkedDays;
                workSheet.Cells[row, "C"] = c.DaysInOffice;

            }

            //workSheet.SaveAs($@"{Environment.CurrentDirectory}\Report.xlsx");
            //excelApp.Quit();
            Console.WriteLine("Report file was saved in your app folder");
            Console.ReadLine();

        }
        static void Main(string[] args)
        {
            List<User> users = new List<User>
            {
               new User { Name="Jonas", WorkedDays=10, DaysInOffice=180},
               new User { Name="Gediminas", WorkedDays=15, DaysInOffice=180},
               new User { Name="Monika", WorkedDays=12.5, DaysInOffice=180},
               new User { Name="Gytis", WorkedDays=14, DaysInOffice=180}
            };

            ExportToExcel(users);
        }
    }
}

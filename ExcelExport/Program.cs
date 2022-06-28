using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelExport
{
    class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            var directory = @"C:\Demos\";
            var fileName = "ExcelDemo.xlsx";

            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var file = new FileInfo(directory + fileName);

            var people = GetSetupData();

            await SaveExcelFile(people, file);
        }

        private static async Task SaveExcelFile(List<PersonModel> people, FileInfo file)
        {
            DeleteIfExists(file);

            using var package = new ExcelPackage(file);

            var ws = package.Workbook.Worksheets.Add("MainReport");

            var range = ws.Cells["A2"].LoadFromCollection(people, true);
            range.AutoFitColumns();


            ws.Cells["A1"].Value = "Our Cool Report";
            ws.Cells["A1:Z1"].Merge = true;
            ws.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Row(1).Style.Font.Size = 24;
            ws.Row(1).Style.Font.Color.SetColor(Color.Blue);

            ws.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Row(2).Style.Font.Bold = true;

            await package.SaveAsync();

        }

        private static void DeleteIfExists(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }

        private static List<PersonModel> GetSetupData()
        {
            List<PersonModel> output = new List<PersonModel>()
            {
                new PersonModel(){Id = 1, Firstname = "Tim", LastName = "Corey"},
                new PersonModel(){Id = 2, Firstname = "Sue", LastName = "Storm"},
                new PersonModel(){Id = 3, Firstname = "Jane", LastName = "Smith"}
            };

            return output;
        }
    }
}

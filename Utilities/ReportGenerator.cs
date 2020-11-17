using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using GenerateExcel.Entities;
using System.Collections.Generic;

namespace GenerateExcel.Services
{
    public class ReportGenerator
    {
        public byte[] GenerateEntryReport(List<User> userList)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var memoryStream = new MemoryStream();

            using (var package = new ExcelPackage(memoryStream))
            {
                var workSheet = package.Workbook.Worksheets.Add("Relatório");

                //carrega o objeto todo
                //workSheet.Cells.LoadFromCollection(entryList, true);

                //configurações primeira linha da tabela
                workSheet.Row(1).Height = 20;
                workSheet.Row(1).Style.Font.Bold = true;
                workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                //configurações cabeçalho primeira linha
                workSheet.Cells[1, 1].Value = "Id";
                workSheet.Cells[1, 2].Value = "Nome do Usuário";
                workSheet.Cells[1, 3].Value = "Idade";
                workSheet.Cells[1, 4].Value = "Email";
                


                //popula o excel
                int entryCount = 2;
                foreach (var user in userList)
                {
                    workSheet.Cells[entryCount, 1].Value = user.Id;
                    workSheet.Cells[entryCount, 2].Value = user.Name;
                    workSheet.Cells[entryCount, 3].Value = user.Age;
                    workSheet.Cells[entryCount, 4].Value = user.Email;

                    entryCount++;
                }

                return package.GetAsByteArray();
            }
        }
    }
}

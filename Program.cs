using GenerateExcel.Entities;
using GenerateExcel.Services;
using GenerateExcel.Utilities;
using System;
using System.Collections.Generic;

namespace GenerateExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Gerando relatório...");

            var userList = new List<User>();

            userList.Add(new User("Lucas", "lucas@eu.com", 19));
            userList.Add(new User("Gabriel", "gabriel@eu.com", 20));
            userList.Add(new User("Pedro", "pedro@eu.com", 21));
            userList.Add(new User("Marcos", "marcos@eu.com", 22));
            userList.Add(new User("Matheus", "matheus@eu.com", 23));
            userList.Add(new User("Rafael", "rafael@eu.com", 24));
            userList.Add(new User("Julio", "julio@eu.com", 25));
            userList.Add(new User("Arnaldo", "arnaldo@eu.com", 26));
            userList.Add(new User("Luiz", "luiz@eu.com", 27));
            userList.Add(new User("José", "jose@eu.com", 28));

            var reportBytes = new ReportGenerator().GenerateEntryReport(userList);

            var path = @"C:\Users\Lucas\Desktop\Github\GenerateExcel\Reports\";
            var filename = $"{Guid.NewGuid().ToString().Substring(0, 8)}_report.xls";
            var fullPath = path + filename;

            FilerSave.Save(fullPath, reportBytes);

            Console.WriteLine("Arquivo salvo com sucesso!");
            
            Console.ReadKey();
        }
    }
}

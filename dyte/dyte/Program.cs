using System;

using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

using System.IO;
using IronXL;
using System.Collections.Generic;

namespace dyte
{
    class Program
    {
        private static string jsonString;

        static void Main(string[] args)
        {

            //Creating a Workbook and adding column Names
            WorkBook xlsWorkbook = WorkBook.Create(ExcelFileFormat.XLS);
            xlsWorkbook.Metadata.Author = "IronXL";
            WorkSheet xlsSheet = xlsWorkbook.CreateWorkSheet("new_sheet");
            dynamic result = Newtonsoft.Json.JsonConvert.DeserializeObject(File.ReadAllText(@"C:\Users\sandh\dyte\package.json"));

            var value = new List<string>();
            var key= new List<string>();

            foreach (var file in result)
            {
                value.Add(file.value); //value from package.json
                key.Add(file); //key from package.json
            }


            //Writing Column Names to Woorksheet
            xlsSheet["A1"].Value = "Name";
            xlsSheet["B1"].Value = "Repo";
            xlsSheet["C1"].Value = "Version";
            
           

            //Opening Chrome Browser
            IWebDriver wd = new ChromeDriver();
            for (int i = 0; i < 3; i++)
            {
                string x = "npm init --version";
                x = Console.ReadLine();
                xlsSheet["C" + i].Value = x;
                xlsSheet["B" + i].Value = value[i];
                xlsSheet["A" + i].Value = key[i];

            }

             

            xlsWorkbook.SaveAs(@"C:\Users\sandh\dyte\DYTE.csv");

        }
    }
}
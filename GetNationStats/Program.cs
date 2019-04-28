using Common.Objects;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;

namespace GetNationStats
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            string paramNationId = "";
            string paramApiKey = "";

            if (args.Length % 2 != 0 || args.Length == 0)
            {
                Console.WriteLine("Parameters: -nationid ### -apikey abcd###");
                return;
            }

            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "-nationid":
                        paramNationId = args[i + 1];
                        break;

                    case "-apikey":
                        paramApiKey = args[i + 1];
                        break;
                }
            }

            using (var httpClient = new HttpClient())
            {
                var nationId = paramNationId;
                var key = paramApiKey;

                httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36");
                var response = httpClient.GetStringAsync(new Uri("http://politicsandwar.com/api/nation/id=" + nationId + "&key=" + key)).Result;              
                Nation nation = JsonConvert.DeserializeObject<Nation>(response);

                var cities = new List<City>();
                foreach (var cityid in nation.cityids)
                {
                    var result = httpClient.GetStringAsync(new Uri("http://politicsandwar.com/api/city/id=" + cityid + "&key=" + key)).Result;
                    var cityJSON = JsonConvert.DeserializeObject<City>(result);
                    var infraNeeded = (((cityJSON.Infrastructure - 2500)) * -1);
                    cityJSON.InfrastructureNeeded = (infraNeeded > 0) ? infraNeeded : 0;
                    cities.Add(cityJSON);
                }               

                FileInfo file = new FileInfo($"{ DateTime.Now.ToString("yyyyMMdd.hhmmssstt") }.xlsx");
                ExcelPackage package = new ExcelPackage(file);
                ExcelWorksheet ws = package.Workbook.Worksheets.Add(cities[0].Nation);
                ws.Cells[1, 1].LoadFromCollection(cities, true);

                // Header Row
                ws.Cells[1, 1, 1, 53].Style.Font.Bold = true;
                ws.Cells[1, 1, 1, 53].Style.Font.Color.SetColor(System.Drawing.Color.White);
                ws.Cells[1, 1, 1, 53].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                ws.Cells[1, 1, 1, 53].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkBlue);

                // Total Row
                ws.Cells[cities.Count() + 2, 1].Value = "Totals";
                ws.Cells[cities.Count() + 2, 1, cities.Count() + 2, 53].Style.Font.Bold = true;
                ws.Cells[cities.Count() + 2, 1, cities.Count() + 2, 53].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                ws.Cells[cities.Count() + 2, 1, cities.Count() + 2, 53].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                package.SaveAs(file);

                return;
            }
        }
    }
}
using Common.Objects;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;

namespace GetTargets
{
    internal class Program
    {
        private static void Main()
        {
            using (var httpClient = new HttpClient())
            {
                var allianceId = 1246;

                httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36");
                var response = httpClient.GetStringAsync(new Uri("http://politicsandwar.com/api/alliance/id=" + allianceId)).Result;
                Alliance alliance = JsonConvert.DeserializeObject<Alliance>(response);

                var nations = new List<Nation>();
                foreach (var nationid in alliance.member_id_list)
                {
                    var result = httpClient.GetStringAsync(new Uri("http://politicsandwar.com/api/nation/id=" + nationid)).Result;
                    var nationJSON = JsonConvert.DeserializeObject<Nation>(result);
                    nations.Add(nationJSON);
                }

                string output = GenerateReport(nations);
                File.WriteAllText($"{DateTime.Now.ToString("yyyyMMdd.hhmmssstt")}.csv", output);

                return;
            }
        }

        private static string GenerateReport<T>(List<T> items) where T : class
        {
            var output = "";
            var delimiter = ",";
            var properties = typeof(T).GetProperties()
             .Where(n =>
             n.PropertyType == typeof(string)
             || n.PropertyType == typeof(bool)
             || n.PropertyType == typeof(char)
             || n.PropertyType == typeof(byte)
             || n.PropertyType == typeof(decimal)
             || n.PropertyType == typeof(int)
             || n.PropertyType == typeof(DateTime)
             || n.PropertyType == typeof(DateTime?));
            using (var sw = new StringWriter())
            {
                var header = properties
                .Select(n => n.Name)
                .Aggregate((a, b) => a + delimiter + b);
                sw.WriteLine(header);
                foreach (var item in items)
                {
                    var row = properties
                    .Select(n => n.GetValue(item, null))
                    .Select(n => n == null ? "null" : n.ToString()).Aggregate((a, b) => a + delimiter + b);
                    sw.WriteLine(row);
                }
                output = sw.ToString();
            }
            return output;
        }
    }
}
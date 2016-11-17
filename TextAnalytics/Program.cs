using ClosedXML.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;

namespace TextAnalytics
{
    class Program
    {
        public static void Main(string[] args)
        {
            MainAsync().Wait();
        }

        public static async Task MainAsync()
        {
            var workbook = new XLWorkbook("PTNB Data (Consolidated) FINAL.xlsx");
            var sheet = workbook.Worksheet(1);
            sheet.Cell("I1").Value = "Negativity";
            sheet.Cell("J1").Value = "Neutrality";
            sheet.Cell("K1").Value = "Positivity";
            sheet.Cell("L1").Value = "Label";
            foreach (var row in sheet.Rows(2, sheet.LastRowUsed().RangeAddress.LastAddress.RowNumber))
            //foreach (var row in sheet.Rows(2, 5))
            {
                var cell = row.Cell(7);
                var requestString = cell.GetString();
                // Create the HttpContent for the form to be posted.
                var requestContent = new FormUrlEncodedContent(new[] { new KeyValuePair<string, string>("text", requestString), });
                Sentiment sentiment = await Post(requestContent);
                Console.WriteLine(row.RowNumber());
                row.Cell(9).Value = sentiment.probability.neg;
                row.Cell(9).Style.NumberFormat.Format = "0.0%";
                row.Cell(10).Value = sentiment.probability.neutral;
                row.Cell(10).Style.NumberFormat.Format = "0.0%";
                row.Cell(11).Value = sentiment.probability.pos;
                row.Cell(11).Style.NumberFormat.Format = "0.0%";
                if (sentiment.label.Equals("pos"))
                    row.Cell(12).Value = "Positive";
                else if (sentiment.label.Equals("neg"))
                    row.Cell(12).Value = "Negative";
                else if (sentiment.label.Equals("error"))
                    row.Cell(12).Value = "Error";
                else
                    row.Cell(12).Value = "Neutral";
                //workbook.SaveAs("PTNB Data (Consolidated) FINAL sentiment.xlsx");
            }
            Console.ReadKey();
            workbook.SaveAs("PTNB Data (Consolidated) FINAL sentiment.xlsx");
        }

        public static async Task<Sentiment> Post(FormUrlEncodedContent requestContent)
        {

            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("X-Mashape-Key", "djAbEOTIz5mshxdvNS1VdnROIeNSp1QCVgujsnaBtedWk8Ez1f");
            // Get the response.
            HttpResponseMessage response = await client.PostAsync("https://japerk-text-processing.p.mashape.com/sentiment/", requestContent);

            if (!response.IsSuccessStatusCode)
                return await Task.FromResult(new Sentiment { probability = new Probability { neg = 0, neutral = 0, pos = 0 }, label = "error" });
            // Get the response content.
            HttpContent responseContent = response.Content;

            // Get the stream of the content.
            using (var reader = new StreamReader(await responseContent.ReadAsStreamAsync()))
            {
                // Write the output.
                string json = await reader.ReadToEndAsync();
                Console.WriteLine(json);
                return await Task.FromResult(JsonConvert.DeserializeObject<Sentiment>(json));
            }

        }
    }

    public class Sentiment
    {
        public Probability probability { get; set; }
        public string label { get; set; }
    }

    public class Probability
    {
        public double neg { get; set; }
        public double neutral { get; set; }
        public double pos { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnDownloadData_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet currentSheet = Globals.ThisAddIn.GetActiveworkSheet();
            String URL = "https://httpbin.org/get";


            using (var webClient = new System.Net.WebClient())
            {
                webClient.Headers.Add("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8");
                webClient.Headers.Add("Accept-Encoding", "gzip, deflate, br");
                webClient.Headers.Add("Accept-Language", "en-US,en;q=0.9,de;q=0.8");
                webClient.Headers.Add("Cookie", "_gauges_unique_month=1; _gauges_unique_year=1; _gauges_unique=1");
                webClient.Headers.Add("Refere", "https://l.facebook.com/");
                webClient.Headers.Add("Upgrade-Insecure-Requests", "1");
                webClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36");
                var json = webClient.DownloadString(URL);
                string valueOriginal = Convert.ToString(json);

                // Now parse with JSON.Net
                JObject joResponse = JObject.Parse(json);
                JObject ojObject = (JObject)joResponse["headers"];

                currentSheet.Columns["A"].ColumnWidth = 30;
                currentSheet.Columns["B"].ColumnWidth = 80;

                currentSheet.Range["A1"].Value = "Accept";
                currentSheet.Range["A2"].Value = "Accept-Encoding";
                currentSheet.Range["A3"].Value = "Accept-Language";
                currentSheet.Range["A4"].Value = "Connection";
                currentSheet.Range["A5"].Value = "Cookie";
                currentSheet.Range["A6"].Value = "Host";
                currentSheet.Range["A7"].Value = "Refere";
                currentSheet.Range["A8"].Value = "Upgrade-Insecure-Requests";
                currentSheet.Range["A9"].Value = "User-Agent";

                currentSheet.Range["B1"].Value = ojObject["Accept"].ToString();
                currentSheet.Range["B2"].Value = ojObject["Accept-Encoding"].ToString();
                currentSheet.Range["B3"].Value = ojObject["Accept-Language"].ToString();
                currentSheet.Range["B4"].Value = ojObject["Connection"].ToString();
                currentSheet.Range["B5"].Value = ojObject["Cookie"].ToString();
                currentSheet.Range["B6"].Value = ojObject["Host"].ToString();
                currentSheet.Range["B7"].Value = ojObject["Refere"].ToString();
                currentSheet.Range["B8"].Value = ojObject["Upgrade-Insecure-Requests"].ToString();
                currentSheet.Range["B9"].Value = ojObject["User-Agent"].ToString();



                MessageBox.Show(json);
            }


        }
    }
}

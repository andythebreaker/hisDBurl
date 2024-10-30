using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Http;
using System.Net;
using System.Text.RegularExpressions;
using System.Security.Cryptography;
using System.Xml.Linq;
using Newtonsoft.Json.Linq;
using System.Security.Policy;
using System.IO;

namespace hisDBurl
{
    public partial class hisDBurlGUI
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
        }

        private void setAllText_Click(object sender, RibbonControlEventArgs e)
        {
            var excelApp = Globals.ThisAddIn.Application;
            var worksheet = excelApp.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;

            if (worksheet != null)
            {
                // Set the format of all cells in the worksheet to Text
                worksheet.Cells.NumberFormat = "@";
            }
        }

        private void addUrl(Int16 colNum, string qry1, string qry2, string url1, string url2)
        {
            // Get the active worksheet
            Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet as Worksheet;

            if (worksheet == null) return;

            // Loop through each row until the last used row
            int lastRow = worksheet.Cells[worksheet.Rows.Count, 1].End(XlDirection.xlUp).Row;

            for (int row = 1; row <= lastRow; row++)
            {
                Range firstCell = worksheet.Cells[row, 1];
                Range secondCell = worksheet.Cells[row, colNum];//2];

                // Check if the first cell can be converted to a number
                if (double.TryParse(firstCell.Text.ToString(), out double numberValue))
                {
                    // Get the text of the second column in the same row
                    string secondColText = secondCell.Text.ToString();

                    // Encode the query in Base64 format
                    string encodedQuery = Convert.ToBase64String(
                        System.Text.Encoding.UTF8.GetBytes(
                            //$"{{\"query\":[{{\"field\":\"in_store_no\",\"value\":\"{secondColText}\"}}]}}"
                            qry1 + secondColText + qry2
                        )
                    );

                    // Create the URL with the encoded query
                    string url = url1 + encodedQuery + url2;//$"https://ahonline.drnh.gov.tw/index.php?act=Archive/search/{encodedQuery}%3D%3D";

                    // Find the last used column in this row and set the URL in the next column
                    //int lastColumn = worksheet.Cells[row, worksheet.Columns.Count].End(XlDirection.xlToLeft).Column;
                    //worksheet.Cells[row, lastColumn + 1].Value = url;

                    // Find the last used column in this row and set the hyperlink in the next column
                    int lastColumn = worksheet.Cells[row, worksheet.Columns.Count].End(XlDirection.xlToLeft).Column;
                    Range targetCell = worksheet.Cells[row, lastColumn + 1];

                    // Add the hyperlink to the target cell
                    worksheet.Hyperlinks.Add(targetCell, url, Type.Missing, "Click to view archive", url);
                }
            }
        }

        private void ahcmsGenUrl_Click(object sender, RibbonControlEventArgs e)
        {
            addUrl(2, "{\"query\":[{\"field\":\"in_store_no\",\"value\":\"", "\"}]}", "https://ahonline.drnh.gov.tw/index.php?act=Archive/search/", "%3D%3D");
        }

        private void ndapGenUrl_Click(object sender, RibbonControlEventArgs e)
        {
            addUrl(3, "{\"query\":[{\"field\":\"identifier\",\"value\":\"", "\"}]}", "https://drtpa.th.gov.tw/index.php?act=Archive/search/", "%3D%3D");
        }

        private void ahtwhGenUrl_Click(object sender, RibbonControlEventArgs e)
        {
            addUrl(3, "{\"search\":[{\"field\":\"list_dataid\",\"value\":\"", "\"}]}", "https://onlinearchives.th.gov.tw/index.php?act=Archive/search/undefined/", "%3D");
        }

        public static string GeneratePhpSessionId(int length = 32)
        {
            // Create a byte array to hold the random bytes
            byte[] randomBytes = new byte[length];

            // Fill the array with cryptographically secure random bytes
            using (var rng = new RNGCryptoServiceProvider())
            {
                rng.GetBytes(randomBytes);
            }

            // Convert the byte array to a hexadecimal string
            StringBuilder sessionId = new StringBuilder(length * 2);
            foreach (byte b in randomBytes)
            {
                sessionId.Append(b.ToString("x2")); // Convert to hex
            }

            return sessionId.ToString();
        }

        private async Task<JObject> FetchAccKeyAndPost(string url, string phpSessionId)
        {
            using (HttpClient client = new HttpClient())
            {
                try
                {
                    // Set the PHPSESSID cookie in the header
                    client.DefaultRequestHeaders.Add("Cookie", $"PHPSESSID={phpSessionId}");
                    // Make the HTTP GET request
                    HttpResponseMessage response = await client.GetAsync(url);
                    response.EnsureSuccessStatusCode();
                    // Read the response body as a string
                    string responseBody = await response.Content.ReadAsStringAsync();
                    // Use regex to find the acckey
                    string pattern = @"acckey='(?<key>[^']+)'";
                    Match match = Regex.Match(responseBody, pattern);
                    if (match.Success)
                    {
                        string acckey = match.Groups["key"].Value;
                        // Prepare the POST request
                        var postData = new System.Collections.Generic.Dictionary<string, string>
                {
                    { "act", $"Display/initial/{acckey}" }
                };
                        var content = new FormUrlEncodedContent(postData);
                        client.DefaultRequestHeaders.Referrer = new Uri(url);
                        // Send the POST request
                        HttpResponseMessage postResponse = await client.PostAsync(url, content);
                        postResponse.EnsureSuccessStatusCode();
                        // Get the response body as JSON
                        string jsonResponse = await postResponse.Content.ReadAsStringAsync();
                        // Optionally parse the JSON
                        JObject json = JObject.Parse(jsonResponse);
                        return json;
                    }
                    else
                    {
                        return null;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred: {ex.Message}", "Error");
                    return null;
                }
            }
        }

        private async Task ProcessSheet(Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            Microsoft.Office.Interop.Excel.Range usedRange = worksheet.UsedRange;

            for (int row = 1; row <= usedRange.Rows.Count; row++)
            {
                for (int col = 1; col <= usedRange.Columns.Count; col++)
                {
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[row, col] as Microsoft.Office.Interop.Excel.Range;
                    string cellText = cell.Text.ToString();

                    if (Uri.IsWellFormedUriString(cellText, UriKind.Absolute))
                    {
                        JObject resJson = await FetchAccKeyAndPost(cellText, GeneratePhpSessionId());

                        if (resJson != null)
                        {
                            Microsoft.Office.Interop.Excel.Range rightCell = worksheet.Cells[row, col + 1] as Microsoft.Office.Interop.Excel.Range;
                            //rightCell.Value2 = ProcessJsonObject(resJson, cellText);//resJson.ToString();
                            var the_url = ProcessJsonObject(resJson, cellText);
                            worksheet.Hyperlinks.Add(rightCell, the_url, Type.Missing, "Click to view archive", the_url);
                        }
                    }
                }
            }
        }



        private string ProcessJsonObject(JObject obj, string url)
        {
            if ((int)obj["action"] == 0)
            {
                return obj["info"].ToString();
            }
            else if ((int)obj["action"] == 1)
            {
                var uri = new Uri(url);
                string domain = uri.Host;
                string display = obj["data"]["display"].ToString();
                string resource = obj["data"]["resouse"].ToString();
                return $"https://{domain}/index.php?act=Display/{display}/{resource}";
            }
            else
            {
                throw new ArgumentException("Invalid action type");
            }
        }


        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var excelApp = Globals.ThisAddIn.Application;
            var worksheet = excelApp.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            _ = ProcessSheet(worksheet);
        }

        private void hoDownload_Click(object sender, RibbonControlEventArgs e)
        {
           // FileDownloader fileDownloader = new FileDownloader(hoacc.Text,hopsw.Text,hosa);
            //_ = fileDownloader.DownloadFileFromAPIAsync();
            if (String.IsNullOrEmpty(hoacc.Text) || String.IsNullOrEmpty(hopsw.Text))
            {
                MessageBox.Show($"account password is empty", "Error");
            }
            else
            {
                 FileDownloader fileDownloader = new FileDownloader(hoacc.Text,hopsw.Text,hosa);
                _ = fileDownloader.DownloadFileFromAPIAsync();
            }
        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {//if (String.IsNullOrEmpty(hoacc.Text) || String.IsNullOrEmpty(hopsw.Text)) {
         //    MessageBox.Show($"account password is empty", "Error");
         //  }
         //  else
         // {
         //   FileDownloader fileDownloader = new FileDownloader(hoacc.Text, hopsw.Text, hosa);
         //  fileDownloader.test1();
         //  }
            hosa.Tag = "ok";
        }

        private void hosa_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show((string)hosa.Tag);
        }
    }
}

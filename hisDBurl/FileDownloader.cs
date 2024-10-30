using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Security.Policy;
using System.Security.Cryptography;
using Microsoft.Office.Tools.Ribbon;

namespace hisDBurl
{
    internal class FileDownloader
    {
        private string hoacc;
        private string hopsw;
        private RibbonButton bbb;

        public FileDownloader(string hoacc, string hopsw, RibbonButton bbb)
        {
            this.hoacc = hoacc;
            this.hopsw = hopsw;
            this.bbb = bbb;
            var jsonObject = new JObject
            {
                ["log"] = new JArray { "start" }
            };
            this.bbb.Tag = jsonObject.ToString();
        }

        public void test1() {
            bbb.Label = "ok";
        }

        private static readonly HttpClient client = new HttpClient();

        private string AtobAccPsw() {
            string zzz =  Convert.ToBase64String(
                            System.Text.Encoding.UTF8.GetBytes(
                                $"{{\"account\":\"{this.hoacc}\",\"password\":\"{this.hopsw}\"}}"
                            )
                        );
            //MessageBox.Show(zzz);
            return zzz;
        }

        public static string GeneratePhpSessionId(int length = 32)//COPY FROM MAIN
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

        public async Task DownloadFileFromAPIAsync()
        {
            try
            {

                client.DefaultRequestHeaders.Add("Cookie", $"PHPSESSID={GeneratePhpSessionId()}");
                // Prepare the POST request
                var postData = new System.Collections.Generic.Dictionary<string, string>
                {
                    { "act", $"Landing/signin/{AtobAccPsw()}%3D%3D" }
                };
                var content = new FormUrlEncodedContent(postData);
                //client.DefaultRequestHeaders.Referrer = new Uri(url);
                // Send the POST request
                HttpResponseMessage postResponse = await client.PostAsync("https://ahonline.drnh.gov.tw/index.php", content);
                postResponse.EnsureSuccessStatusCode();
                // Get the response body as JSON
                string xxx = await postResponse.Content.ReadAsStringAsync();
                JObject xx = JObject.Parse(xxx);
                // MessageBox.Show(xxx);
                var bb = JObject.Parse((string)bbb.Tag);
                ((JArray)bb["log"]).Add(xx.ToString());
                bbb.Tag = bb.ToString();
                var lgKey = xx["data"]?["lgkey"]?.ToString();

                HttpResponseMessage response0 = await client.GetAsync($"https://ahonline.drnh.gov.tw/index.php?act=Landing/inter/{lgKey}");
                response0.EnsureSuccessStatusCode();//302found
                //  Read the response body as a string
               // string responseBody0 = await response0.Content.ReadAsStringAsync();
               // MessageBox.Show(responseBody0);
               // MessageBox.Show("????????");

                // Define the initial URL and PHPSESSID cookie
                string initialUrl = "https://ahonline.drnh.gov.tw/index.php?act=Display/package/5013216=BX5aC0/1-2";
               // string cookie = "PHPSESSID=slhr6laimtp2nlonf53g4g9294";

                // Set up the HttpClient headers and cookie
                //client.DefaultRequestHeaders.Add("Cookie", cookie);

                // First GET request to retrieve JSON data
                //var response = await client.GetStringAsync(initialUrl);
                // Make the HTTP GET request
                HttpResponseMessage response = await client.GetAsync(initialUrl);
                response.EnsureSuccessStatusCode();
                // Read the response body as a string
                string responseBody = await response.Content.ReadAsStringAsync();
                JObject jsonResponse = JObject.Parse(responseBody);

                // Check if "action" equals 1
                if (jsonResponse["action"]?.Value<int>() == 1)
                {
                    string data = jsonResponse["data"]?.ToString();
                    if (string.IsNullOrEmpty(data))
                    {
                        Console.WriteLine("Data field is empty.");
                        return;
                    }

                    // Construct the download URL
                    string downloadUrl = $"https://ahonline.drnh.gov.tw/index.php?act=Display/download/{data}";

                    // Second GET request to download the file
                    HttpResponseMessage downloadResponse = await client.GetAsync(downloadUrl);
                    if (downloadResponse.IsSuccessStatusCode)
                    {
                        // Determine the content disposition for the file name
                        //string fileName = downloadResponse.Content.Headers.ContentDisposition?.FileNameStar ??
                        //                downloadResponse.Content.Headers.ContentDisposition?.FileName ??
                        //              "DownloadedFile.pdf";
                        string fileName = "fuckme2.pdf";

                    // Clean the file name
                    fileName = WebUtility.HtmlDecode(fileName);

                        // Set the download path to %tmp% directory
                        string downloadPath = Path.Combine(Path.GetTempPath(), fileName);

                        // Save the file to %tmp%
                        using (var fileStream = new FileStream(downloadPath, FileMode.Create, FileAccess.Write))
                        {
                            await downloadResponse.Content.CopyToAsync(fileStream);
                        }

                        MessageBox.Show($"File downloaded successfully to {downloadPath}");
                    }
                    else
                    {
                        MessageBox.Show("Failed to download the file.");
                    }
                }
                else
                {
                    MessageBox.Show($"Action not equal to 1. No file downloaded.{jsonResponse.ToString()}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message},{ex.TargetSite},{ex.Source},{ex.Data}", "Error");
            }
        }
    }
}

using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Http;
using Newtonsoft.Json;

namespace HomsV2.Functions
{
    class DPR
    {
        private readonly JArray productionData;
        private readonly JObject workCenterData;
        private readonly JObject userData;

        private readonly string templatePath = @"\\apbiphbpsts01\htdocs\homs\resources\DPR\Templates\";
        private readonly string generatedDPRPath = @"\\apbiphbpsts01\htdocs\homs\resources\DPR\Generated\";

        public DPR(JArray productionData, JObject workCenterData, JObject userData)
        {
            this.productionData = productionData;
            this.workCenterData = workCenterData;
            this.userData = userData;

            dataCheck();
        }

        public void dataCheck()
        {
            if (!(this.productionData is JArray prod &&
                this.workCenterData is JObject wc &&
                this.userData is JObject user))
            {
                MessageBox.Show("Invalid JSON structure for DPR generation.");
                Console.WriteLine("Invalid JSON structure for DPR generation.");
                throw new Exception("Invalid JSON structure for DPR generation.");
            }
        }

        public string generateDPRFile()
        {
            // Logic to generate DPR file
            string productionData = this.productionData.ToString(Newtonsoft.Json.Formatting.Indented);
            string workCenterData = this.workCenterData.ToString(Newtonsoft.Json.Formatting.Indented);
            string userData = this.userData.ToString(Newtonsoft.Json.Formatting.Indented);

            //Identify which template to use
            string templateFileName = DetermineTemplateFileName();

            string generatedPath;

            ExcelPackage.License.SetNonCommercialOrganization("DPR");

            string fileName;

            using (var package = new ExcelPackage(templateFileName))
            {
                bool isDaySheetInitialized = false;
                bool isNightSheetInitialized = false;

                JObject[] reversedProductionItems = MergedData(this.productionData)
                                                                    .Reverse()
                                                                    .Cast<JObject>()
                                                                    .ToArray();

                foreach (JObject productionItem in reversedProductionItems)
                {
                    if (productionItem["shift"].ToString() == "ds")
                    {
                        var daySheet = package.Workbook.Worksheets["Dayshift_B"];

                        if (!isDaySheetInitialized)
                        {
                            daySheet.Cells["S2"].Value = this.userData["Full_Name"].ToString();
                            daySheet.Cells["S3"].Value = this.userData["Section"].ToString();

                            DateTime timeCreated = DateTime.Parse(this.productionData[0]["time_created"].ToString());
                            daySheet.Cells["T4"].Value = timeCreated.ToString("MM/dd/yyyy");
                            daySheet.Cells["T5"].Value = this.workCenterData["workcenter"].ToString();

                            daySheet.Cells["V3"].Value = "Run";
                            daySheet.Cells["V4"].Value = this.workCenterData["plant"].ToString();
                            daySheet.Cells["V5"].Value = productionItem["line_name"].ToString();

                            isDaySheetInitialized = true;
                        }

                        processProductionInfo(daySheet, productionItem);
                        processLineStop(daySheet, reversedProductionItems, productionItem);
                    }
                    else
                    {
                        var nightSheet = package.Workbook.Worksheets["Nightshift_Y"];

                        if (!isNightSheetInitialized)
                        {
                            nightSheet.Cells["S2"].Value = this.userData["Full_Name"].ToString();
                            nightSheet.Cells["S3"].Value = this.userData["Section"].ToString();

                            DateTime timeCreated = DateTime.Parse(this.productionData[0]["time_created"].ToString());
                            nightSheet.Cells["T4"].Value = timeCreated.ToString("MM/dd/yyyy");
                            nightSheet.Cells["T5"].Value = this.workCenterData["workcenter"].ToString();

                            nightSheet.Cells["V3"].Value = "Run";
                            nightSheet.Cells["V4"].Value = this.workCenterData["plant"].ToString();
                            nightSheet.Cells["V5"].Value = productionItem["line_name"].ToString();

                            isNightSheetInitialized = true;
                        }

                        processProductionInfo(nightSheet, productionItem);

                    }
                }
                // Save to file

                fileName = this.workCenterData["workcenter"] + "-" + this.productionData[0]["time_created"].ToString() + ".xlsm";
                fileName = fileName.Replace(":", "-");

                var newFile = new FileInfo(generatedDPRPath + fileName);
                generatedPath = newFile.FullName;

                package.SaveAs(newFile);
            }

            return fileName;
        }

        private string DetermineTemplateFileName()
        {
            // Logic to determine the template file name
            string templateFileName = this.workCenterData["dpr_template"].ToString() + ".xlsm";

            if (string.IsNullOrEmpty(templateFileName))
            {
                MessageBox.Show("Template file name is not specified in the work center data.");
                Console.WriteLine("Template file name is not specified in the work center data.");
                throw new Exception("Template file name is not specified in the work center data.");
            }

            string fullPath = templatePath + templateFileName;

            return fullPath;
        }

        private void processProductionInfo(ExcelWorksheet sheet, JToken productionItem)
        {
            try
            {
                int startingRow = 16;
                string startingColumn = "E";

                int row = startingRow;

                while (sheet.Cells[startingColumn + row].Value != null)
                {
                    row++;
                }

                //Production Information

                // Column E: Material (treat as plain string, but strip any leading apostrophe)
                string material = productionItem["material"]?.ToString().TrimStart('\'') ?? "";
                sheet.Cells["E" + row].Value = material;

                // Column F: Plan Quantity
                if (double.TryParse(productionItem["plan_quantity"]?.ToString().TrimStart('\''), out double planQty))
                {
                    sheet.Cells["F" + row].Value = planQty;
                    sheet.Cells["F" + row].Style.Numberformat.Format = "0"; // optional: no decimals
                }
                else
                {
                    sheet.Cells["F" + row].Value = productionItem["plan_quantity"]?.ToString();
                }

                // Column G: Actual Quantity
                if (double.TryParse(productionItem["actual_quantity"]?.ToString().TrimStart('\''), out double actualQty))
                {
                    sheet.Cells["G" + row].Value = actualQty;
                    sheet.Cells["G" + row].Style.Numberformat.Format = "0"; // optional: no decimals
                }
                else
                {
                    sheet.Cells["G" + row].Value = productionItem["actual_quantity"]?.ToString();
                }


                if (DateTime.TryParse(productionItem["start_time"]?.ToString().TrimStart('\''), out DateTime startTime))
                {
                    sheet.Cells["I" + row].Value = startTime;
                    sheet.Cells["I" + row].Style.Numberformat.Format = "HH:mm";
                }
                else
                {
                    sheet.Cells["I" + row].Value = ""; // or some fallback
                }

                if (DateTime.TryParse(productionItem["end_time"]?.ToString().TrimStart('\''), out DateTime endTime))
                {
                    sheet.Cells["J" + row].Value = endTime;
                    sheet.Cells["J" + row].Style.Numberformat.Format = "HH:mm";
                }
                else
                {
                    sheet.Cells["J" + row].Value = "";
                }


                string startTimeStr = startTime.ToString("HH:mm");
                string endTimeStr = endTime.ToString("HH:mm");

                sheet.Cells["K" + row].Value = GetRestTime(this.userData["Section"].ToString(), startTimeStr, endTimeStr);



                // Linestop info logic comes here...

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error processing production info: " + ex.Message);
            }
        }

        private JArray MergedData(JArray originalArray)
        {
            JArray mergedArray = new JArray();

            for (int i = 0; i < originalArray.Count; i += 2)
            {
                JObject first = (JObject)originalArray[i];
                JObject second = (i + 1 < originalArray.Count) ? (JObject)originalArray[i + 1] : null;

                if (second != null)
                {
                    JObject merged = new JObject(second); // Start with second
                    foreach (var prop in first.Properties())
                    {
                        var secondValue = merged[prop.Name];
                        bool isEmpty = secondValue == null ||
                                       secondValue.Type == JTokenType.Null ||
                                       (secondValue.Type == JTokenType.String && (string)secondValue == "") ||
                                       (secondValue.Type == JTokenType.Integer && (int)secondValue == 0) ||
                                       (secondValue.Type == JTokenType.Float && (float)secondValue == 0.0) ||
                                       (secondValue.Type == JTokenType.String && (string)secondValue == "0");

                        if (isEmpty)
                        {
                            merged[prop.Name] = prop.Value;
                        }
                    }
                    mergedArray.Add(merged);
                }
                else
                {
                    mergedArray.Add(first); // Add last unpaired item if odd count
                }
            }

            return mergedArray;
        }

        private string GetRestTime(string section, string startTime, string endTime)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    string url = $@"http://apbiphbpswb02/homs/api/production/getComputedRestTime.php?section={section}&start_time={startTime}&end_time={endTime}";

                    HttpResponseMessage response = client.GetAsync(url).Result;
                    if (response.IsSuccessStatusCode)
                    {
                        string jsonString = response.Content.ReadAsStringAsync().Result;

                        if (string.IsNullOrWhiteSpace(jsonString) || !jsonString.TrimStart().StartsWith("{"))
                        {
                            throw new Exception("Invalid response format. Expected JSON object.");
                        }

                        JObject json = JObject.Parse(jsonString);
                        string status = json["status"]?.ToString();

                        if (status == "success")
                        {
                            int totalRestTime = (int)json["data"]?["total_rest_time"];
                            return totalRestTime.ToString();
                        }
                        else
                        {
                            throw new Exception("GetRestTime - Error: " + json["message"]?.ToString());
                        }
                    }

                    return "0";
                }
            }
            catch (Exception ex)
            {
                try
                {
                    // Just in case MessageBox throws, don't crash the whole thing, baka~
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch { /* swallow UI error just in case */ }

                return "0";
            }
        }

        private void processLineStop(ExcelWorksheet sheet, JObject[] productionItems, JToken item)
        {
            for (int i = 0; i < productionItems.Length; i++)
            {
               Console.WriteLine($"{i} PARTNER 1: " + productionItems[i]["end_time"].ToString());
            }
        }
    }
}

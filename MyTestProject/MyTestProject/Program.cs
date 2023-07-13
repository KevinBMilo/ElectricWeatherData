using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;

namespace MyTestProject
{
    class Program
    {
        public class ReadAndParseJsonFileWithNewtonsoftJson
        {
            private readonly string _sampleJsonFilePath;
            public ReadAndParseJsonFileWithNewtonsoftJson(string sampleJsonFilePath)
            {
                _sampleJsonFilePath = sampleJsonFilePath;
            }
        }

        public class Usage
        {
            public Usage(string timeStamp, int hour, double kiloWatts) {
                TimeStamp = timeStamp;
                Hour = hour;
                KiloWatt = kiloWatts;
            }
            public string? TimeStamp { get; set; }
            public int Hour { get; set; }
            public double KiloWatt { get; set; }
        }
        

        static void Main(string[] args)
        {

            //in the full version, how many minutes in a month?  It depends.  30 days?
            var startDate = "4/1/2023 2:00";
            var endDate = "4/15/2023 4:55";
            var startHour = 7;
            var endHour = 20;

            using StreamReader reader = new(@"C:\Users\oddys\source\repos\MyTestProject\MyTestProject\Assets\Wilmington.JSON");
            var firstJson = reader.ReadToEnd();

            using StreamReader secondReader = new(@"C:\Users\oddys\source\repos\MyTestProject\MyTestProject\Assets\Princeton.JSON");
            var secondJson = reader.ReadToEnd();

            CalculatePeakUsage(startDate, endDate, startHour, endHour);

            ConvertJsonFileToCSV(firstJson);
        }

        public static string CalculatePeakUsage(string startDate, string endDate, int startHour, int endHour)
        {
            var allMonthlyUsageData = GetMonthlyUsageFromExcel();

            //string to date
            var parsedStartDate = DateTime.Parse(startDate);
            var parsedEndDate = DateTime.Parse(endDate);

            //grabs the data to determine the monthly billing cycle
            var billCycle = allMonthlyUsageData
                .Where(sum => DateTime.Parse(sum.TimeStamp) >= parsedStartDate && DateTime.Parse(sum.TimeStamp) <= parsedEndDate);

            //grabs the starting and ending hour for on peak hour
            var peakDemand = billCycle
                .Where(hour => hour.Hour >= startHour && hour.Hour <= endHour);

            
            //optimize this
            var sumsOfAllPeakHours = peakDemand
                .Select(usage => usage.KiloWatt);

            
            //the greatest amount of KiloWatts to determine highest peak hour of the month
            var peakUsage = peakDemand
                .MaxBy(max => max.KiloWatt);

            
            Console.WriteLine($"Peak usage was " + peakUsage?.KiloWatt + " on " + peakUsage?.TimeStamp + " at hour " + peakUsage?.Hour);

            return $"Peak usage was " + peakUsage?.KiloWatt + " on " + peakUsage?.TimeStamp + " at hour " + peakUsage?.Hour;
        }

        public static List<Usage> GetMonthlyUsageFromExcel()
        {
            Application xlsApp = new Application();
            _Workbook wrk = xlsApp.Workbooks.Open(@"C:\Users\oddys\source\repos\MyTestProject\MyTestProject\Assets\rbhs-campus-load-april-2023.xlsx", 0, true, 5, Missing.Value, Missing.Value, true, Excel.XlPlatform.xlWindows, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            var rowCount = xlsApp.Cells.Find("*", Missing.Value, Missing.Value, Missing.Value, XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious, false, Missing.Value, Missing.Value).Row;

            List<Usage> monthlyUsageList = new List<Usage>();

            int i = 2;
            while (i <= rowCount) {

                var timeStampFromExcel = Convert.ToString((xlsApp.Cells[i, 1]).Value());
                var hourFromExcel = Convert.ToInt32((xlsApp.Cells[i, 2]).Value());

                var totalKiloWatts = 0.00;
                for(int k = 0; k < 12; k++) {
                    totalKiloWatts += Convert.ToDouble(xlsApp.Cells[i, 3].Value());
                }

                monthlyUsageList.Add(new Usage(timeStampFromExcel, hourFromExcel, totalKiloWatts));

                i = i + 12;
            }

            //Only for testing purposes
            /*int n = 0;
            while (n < 100)
            {
                Console.WriteLine(monthlyUsageList[n].TimeStamp + " - " + monthlyUsageList[n].Hour + " - " + monthlyUsageList[n].KiloWatt);
                n++;
            }*/

            return monthlyUsageList;
        }

        public static string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";

            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return columnName;
        }

        public static string ConvertJsonFileToCSV(string json)
        {         
            Application xlsApp = new Application();
            _Workbook wrk = xlsApp.Workbooks.Add();

            JObject obj = JObject.Parse(json);

            int headerColumnNumber = 1;

            foreach ( var o in obj)
            {
                xlsApp.Cells[1, headerColumnNumber] = o.Key;
                headerColumnNumber++;
            }

            int row = 2;
            int currentColumn = 1;

            foreach (var o in obj)
            {
                for (int i = 0; i < o.Value?.Count(); i++)
                {
                    xlsApp.Cells[row, currentColumn] = o.Value?[i];
                    row++;
                }
                row = 2;
                currentColumn++;

                //eliminate after testing
                /*if (currentColumn == 5)
                {
                    break;
                }*/
            }

            //Ideally, we would make this its own function to run after the excel file was populated.

            //Unnecessary code.  We can just always go to UTC instead of checking if local is present or not.
            /*string timeHeader;

            int timeColumnNumber = 1;

            bool hasLocalTime = obj.ContainsKey("validTimeLocal");

            if (hasLocalTime == true)
            {
                timeHeader = "validTimeLocal";
            } else
            {
                timeHeader = "validTimeUtc";
            }*/


            int timeColumnNumber = 1;
            foreach (var o in obj)
            {
                //if time is already in first column, forget about this function.
                if (timeColumnNumber > 1)
                {
                    if (o.Key == "validTimeUtc")
                    {
                        string localTimeColumn = GetExcelColumnName(timeColumnNumber);

                        Excel.Range copyRange = xlsApp.Range[localTimeColumn + ":" + localTimeColumn];

                        Excel.Range insertRange = xlsApp.Range["A:A"];
                        insertRange.Insert(XlInsertShiftDirection.xlShiftToRight, copyRange.Cut());
                    }
                }
                timeColumnNumber++;
            }

            Excel.Range changeRange = xlsApp.Range["A:A"];

            for (int i = 2; i < changeRange.Rows.Count; i++)
            {
                var cell = ((xlsApp.Cells[i, 1])?.Value);
                if(cell == null)
                {
                    break;
                }
                double oldTime = ((xlsApp.Cells[i, 1])?.Value);
                DateTime beginEpoch = new DateTime(1970, 1, 1, 0, 0, 0, 0);
                beginEpoch = beginEpoch.AddSeconds(oldTime);
                beginEpoch.AddHours(-4);
                xlsApp.Cells[i, 1] = beginEpoch;
            }

            //Show finished product, so the user can make changes and save the file.
            xlsApp.Visible = true;

            return "This is a placeholder.";
        }

        
    }
}
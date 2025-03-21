using System;
using System.Collections.Concurrent;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace WindowsFormsApp2
{
    public class DataPersistence
    {
        private static readonly BlockingCollection<BusinessInfo> queue = new BlockingCollection<BusinessInfo>();
        private static readonly Task writerTask;
        private static int serialNumber = 0;
        private static readonly string filePath = AppDataHelper.GetAppDataPath();

        static DataPersistence()
        {
            writerTask = Task.Run(ProcessQueue); // Start the background writer
        }

        // ✅ Add Data to the Queue (Non-Blocking)
        public static void SaveDataToCSV(BusinessInfo newData)
        {
            serialNumber++;
            newData.serialNumber = serialNumber;
            queue.Add(newData);
        }

        // ✅ Background Task to Write Data Efficiently
        private static void ProcessQueue()
        {
            bool fileExists = File.Exists(filePath);
            using (StreamWriter sw = new StreamWriter(filePath, append: true, Encoding.UTF8))
            {
                if (!fileExists)
                {
                    string columnHeader = "";
                    // ✅ Write header only if file does not exist
                    for (int index = 0; index < SharedDataTableModel.SelectedFields.Count; index++)
                    {
                        columnHeader += $"{SharedDataTableModel.SelectedFields[index].Name},";
                    }

                    columnHeader = columnHeader.Substring(0, columnHeader.Length - 1);
                    sw.WriteLine(columnHeader);
                }

                foreach (var data in queue.GetConsumingEnumerable(CancellationToken.None))
                {
                    // ✅ Batch save to avoid multiple file writes
                    sw.WriteLine($"" +
                        $"{data.serialNumber}," +
                        $"{EscapeCSVValue(data.SearchTerm)}," +
                        $"{EscapeCSVValue(data.ResultTitle)}," +
                        $"{EscapeCSVValue(data.Category)}," +
                        $"{EscapeCSVValue(data.Address)}," +
                        $"{EscapeCSVValue(data.StreetAddress)}," +
                        $"{EscapeCSVValue(data.City)}," +
                        $"{EscapeCSVValue(data.Zip)}," +
                        $"{EscapeCSVValue(data.Country)}," +
                        $"{EscapeCSVValue(data.ContactNumber)}," +
                        $"{EscapeCSVValue(data.SocialMedias.ContainsKey("emails") ? data.SocialMedias["emails"] : string.Empty)}," +
                        $"{EscapeCSVValue(data.CompanyWebsite)}," +
                        $"{EscapeCSVValue(data.SocialMedias.ContainsKey("facebook") ? data.SocialMedias["facebook"] : string.Empty)}," +
                        $"{EscapeCSVValue(data.SocialMedias.ContainsKey("Linkedin") ? data.SocialMedias["Linkedin"] : string.Empty)}," +
                        $"{EscapeCSVValue(data.SocialMedias.ContainsKey("x") ? data.SocialMedias["x"] : string.Empty)}," +
                        $"{EscapeCSVValue(data.SocialMedias.ContainsKey("youtube") ? data.SocialMedias["youtube"] : string.Empty)}," +
                        $"{EscapeCSVValue(data.SocialMedias.ContainsKey("instagram") ? data.SocialMedias["instagram"] : string.Empty)}," +
                        $"{EscapeCSVValue(data.SocialMedias.ContainsKey("pinterest") ? data.SocialMedias["pinterest"] : string.Empty)}," +
                        $"{EscapeCSVValue(data.Rating)}," +
                        $"{EscapeCSVValue(data.ReviewCount)}," +
                        $"{EscapeCSVValue(data.Claim.ToString())}," +
                        $"{EscapeCSVValue( data.HoursInfo)}, " +
                        $"{EscapeCSVValue(data.BusinessHours.ContainsKey("Monday") ? data.BusinessHours["Monday"] : string.Empty)}," +
                        $"{EscapeCSVValue(data.BusinessHours.ContainsKey("Tuesday") ? data.BusinessHours["Tuesday"] : string.Empty)}," +
                        $"{EscapeCSVValue(data.BusinessHours.ContainsKey("Wednesday") ? data.BusinessHours["Wednesday"] : string.Empty)}," +
                        $"{EscapeCSVValue(data.BusinessHours.ContainsKey("Thursday") ? data.BusinessHours["Thursday"] : string.Empty)}," +
                        $"{EscapeCSVValue(data.BusinessHours.ContainsKey("Friday") ? data.BusinessHours["Friday"] : string.Empty)}," +
                        $"{EscapeCSVValue(data.BusinessHours.ContainsKey("Saturday") ? data.BusinessHours["Saturday"] : string.Empty)}," +
                        $"{EscapeCSVValue(data.BusinessHours.ContainsKey("Sunday") ? data.BusinessHours["Sunday"] : string.Empty)}," +
                        $"{EscapeCSVValue(data.LocatedIn)}," +
                        $"{EscapeCSVValue(data.Attributes)}," +
                        $"{EscapeCSVValue(data.MapLink)}"
                        );
                }
            }
        }

        public static string EscapeCSVValue(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";

            value = value.Replace("\n", " ") // ✅ Replace actual newlines
                         .Replace("\r", " ") // ✅ Replace carriage returns
                         .Replace("%0A", " ") // ✅ Handle URL-encoded newlines
                         .Replace("%0D", " ") // ✅ Handle carriage return encoding
                         .Replace("&#038;", "&"); // ✅ Decode special characters

            if (value.Contains(",") || value.Contains("\""))
            {
                value = "\"" + value.Replace("\"", "\"\"") + "\""; // ✅ Double quotes inside text are escaped
            }
            return value;
        }

        // ✅ Call this on app exit to make sure all data is saved
        public static void StopProcessing()
        {
            queue.CompleteAdding();
            writerTask.Wait();
        }
    }
}

using System.IO;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using System.Data;
using Spire.Xls;

namespace main.function
{
    public class ParseXLSBFunction
    {
        [FunctionName("ParseXLSBFunction")]
        public void Run([BlobTrigger("samples/{name}", Connection = "AzureWebJobsStorage")]Stream myBlob,
                        [Blob("output/{name}.xlsx", FileAccess.Write, Connection = "AzureWebJobsStorage")] Stream outputBlob, 
                        string name, ILogger log)
        {
            System.DateTime start = System.DateTime.Now;
            log.LogInformation($"C# Blob trigger function Processed blob\n Name:{name} \n Size: {myBlob.Length} Bytes");
            Workbook workbook = new Workbook();
            workbook.LoadFromStream(myBlob);
            // Get the "Blank 3-U" worksheet
            Worksheet worksheet = workbook.Worksheets["Blank 3-U"];
            DataTable dt = worksheet.ExportDataTable(worksheet.AllocatedRange, false, true);
            Workbook outputWorkbook = new Workbook();
            Worksheet outputWorksheet = outputWorkbook.Worksheets[0];
            outputWorksheet.InsertDataTable(dt, true, 1, 1);
            outputWorkbook.SaveToStream(outputBlob, FileFormat.Version2013);

            // Dispose the workbook
            workbook.Dispose();
            outputWorkbook.Dispose();
            System.DateTime end = System.DateTime.Now;
            log.LogInformation($" DigiIP C# Azure Function execution time: {end-start}"); 
        }
    }
}

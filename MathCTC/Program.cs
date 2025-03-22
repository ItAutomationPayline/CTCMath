using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;

namespace MathCTC
{
    public class Program
    {
        public static string sourceFolder = @"E:/PAYROLL_SERVER/Automation/Input";
        public static string outputFilePath = @"D:/TDS reconciliation Automation/Output/";
        public static int row2 = 2;
        public static int ErrorCount = 0;
        public static string errors = @"E:/PAYROLL_SERVER/Automation/Errors";
        public static string archived = @"E:/PAYROLL_SERVER/Automation/Archived";
        public static string filePath1 = "";
        public static string destination = @"E:/PAYROLL_SERVER/Automation/output";
        public static string destinationFolder = @"E:/PAYROLL_SERVER/Automation/output";
        public static string ascendcodes;
        public static string ClientName = "";
        public static string ShrinkString(string input)
        {
            if (input != null)
            {
                input = input.ToLower();
                input = input.Replace(" ", "");
                return input;
            }
            return "";
        }
        public static int getColumnNumber(string filepath, string worksheetname, string columnname)
        {
            try
            {
                columnname = columnname.ToLower();
                columnname = columnname.Replace(" ", "");
                using (var package = new ExcelPackage(new FileInfo(filepath)))
                {
                    var inputWorkSheet = package.Workbook.Worksheets[worksheetname];
                    int col = 1;
                    int totalColumns = inputWorkSheet.Dimension.End.Column;
                    for (col = 1; col <= totalColumns; col++)
                    {
                        string temp = inputWorkSheet.Cells[1, col].Text.ToLower();
                        temp = temp.Replace(" ", "");
                        if (columnname.Equals(temp))
                        {
                            return col; // Return the column number if the header matches
                        }
                    }
                    col = -1;
                    if (col == -1)
                    {
                        PathLog(columnname + " column was not found in " + worksheetname + " of " + filepath + " file.");
                        //ErrorCount++;
                    }
                    return col;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                PathLog(columnname + " column was not found in" + worksheetname + " of " + filepath + " file.");
                throw;
            }
        }
        //Method to get position of Sheet
        public static int getSheetNumber(string filepath, string worksheetname)
        {
            try
            {
                worksheetname = ShrinkString(worksheetname);
                using (var package = new ExcelPackage(new FileInfo(filepath)))
                {
                    int worksheetCount = package.Workbook.Worksheets.Count;
                    int i = 0;
                    for (i = worksheetCount - 1; i >= 0; i--)
                    {
                        string temp = package.Workbook.Worksheets[i].Name;
                        temp = ShrinkString(temp);
                        if (temp.Equals(worksheetname))
                        {
                            return i;
                        }
                    }
                    i = 0;
                    if (i == 0)
                    {
                        PathLog(worksheetname + " sheet was not found in " + filepath);
                        //ErrorCount++;
                    }
                    return i;
                }
            }
            catch (Exception e)
            {
                PathLog(e.Message);
                throw;
            }
        }
        public static void Main(string[] args)
        {
            var TDSFiles = Directory.GetFiles(sourceFolder + "/", "*.xlsx")
                .OrderByDescending(f => new FileInfo(f).CreationTime).ToList();

            if (TDSFiles.Count == 0)
            {
                Console.WriteLine("Required file not found. Ensure there is file with '.xlsx' in their name.");
                return;
            }
            filePath1 = filePath1 + TDSFiles.First();
            string filePath = filePath1;
            Console.WriteLine($"Using file: {filePath1}");
            DateTime now = DateTime.Now;
            // Format the month and year as "Month_Year"
            string formattedDate = $"{now:dd_MMMM_yyyy}";
            string foldername = Path.GetFileName(filePath);
            foldername = foldername.Replace(".xlsx", "");
            string filename = Path.GetFileName(filePath.ToLower());
            string[] directories = Directory.GetDirectories(destinationFolder);

            // Extract only the folder names
            //method to find right folder
            string[] folderNames = Array.ConvertAll(directories, dir => Path.GetFileName(dir.ToLower()));
            foreach (string folderName in folderNames)
            {
                if (!folderName.Contains(' '))
                {
                    if (filename.ToLower().Contains(folderName.ToLower()))
                    {
                        Console.WriteLine(folderName);
                        destinationFolder = destinationFolder + "/" + folderName;
                        string[] referencefile = Directory.GetFiles((destinationFolder), "*.xlsx");
                        ascendcodes = destinationFolder + "/" + Path.GetFileName(referencefile[0]);
                        destinationFolder = destinationFolder + "/" + folderName + " " + formattedDate;
                        destination = destinationFolder;
                        Console.WriteLine(ascendcodes);
                        ClientName = folderName;
                        break;
                    }
                }
                //in case of spaces in folder name
                else
                {
                    string[] parts = folderName.Split(' ');
                    int count = parts.Length;
                    int temp = 0;
                    foreach (string part in parts)
                    {
                        if (filename.ToLower().Contains(part.ToLower()))
                        {
                            temp++;
                        }
                    }
                    if (temp == count)
                    {
                        destinationFolder = destinationFolder + "/" + folderName;
                        string[] referencefile = Directory.GetFiles((destinationFolder), "*.xlsx");
                        ascendcodes = destinationFolder + "/" + Path.GetFileName(referencefile[0]);
                        destinationFolder = destinationFolder + "/" + folderName + " " + formattedDate;
                        destination = destinationFolder;
                        Console.WriteLine(ascendcodes);
                        ClientName = folderName;
                        break;
                    }
                }
            }
            if (!Directory.Exists(foldername))
            {
                Directory.CreateDirectory(destinationFolder);
            }
            // Ensure file is fully available by checking in a loop until it's accessible
            for (int retries = 0; retries < 5; retries++)
            {
                try
                {
                    using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        stream.Close();
                        break; // If accessible, break the loop
                    }
                }
                catch (IOException)
                {
                    //await Task.Delay(500); // Wait and retry if file is still being written
                }
            }
            if (filePath.ToLower().Contains("twilio"))
            {
                TwilioCTC.CTC_Master(ascendcodes, filePath, destinationFolder);
            }
            if (filePath.ToLower().Contains("mcafee"))
            {
                McAfee_CTC.CTC_Master(ascendcodes, filePath, destinationFolder);
            }
            if (filePath.ToLower().Contains("alter"))
            {
                Alter_CTC.CTC_Master(ascendcodes, filePath, destinationFolder);
            }
            if (!Directory.Exists(archived))
            {
                Directory.CreateDirectory(archived);
            }
            if (!Directory.Exists(errors))
            {
                Directory.CreateDirectory(errors);
            }
            if (File.Exists(filePath))
            {
                if (ErrorCount == 0)
                {
                    if (File.Exists(archived + "/" + Path.GetFileName(filePath)))
                    {
                        File.Delete(filePath);
                    }
                    else
                    {
                        File.Move(filePath, Path.Combine(archived, Path.GetFileName(filePath)));
                    }
                }
                else
                {
                    if (File.Exists(errors + "/" + Path.GetFileName(filePath)))
                    {
                        File.Delete(filePath);
                    }
                    else
                    {
                        File.Move(filePath, Path.Combine(errors, Path.GetFileName(filePath)));
                    }
                    ErrorCount = 0;
                }
            }
            Log($"Processed file: {Path.GetFileName(filePath)}");
        }
        public static void Log(string message)
        {
            try
            {
                DateTime today = DateTime.Today;
                string _logFilePath = @"E:\PAYROLL_SERVER\Automation\ServiceLogs\" + today.ToString("dd/MMMM/yyyy") + "_PayrollAutomationService.log";
                Directory.CreateDirectory(Path.GetDirectoryName(_logFilePath));
                File.AppendAllText(_logFilePath, $"{DateTime.Now}: {message}{Environment.NewLine}");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                // Fail silently if logging fails to avoid crashing the service
            }
        }
        public static void PathLog(string message)
        {
            try
            {
                DateTime today = DateTime.Today;
                string _logFilePath = destination + "/" + "_PayrollAutomationService_" + today.ToString("dd/MMMM/yyyy") + ".log";
                Directory.CreateDirectory(Path.GetDirectoryName(_logFilePath));
                File.AppendAllText(_logFilePath, $"{DateTime.Now}: {message}{Environment.NewLine}");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                // Fail silently if logging fails to avoid crashing the service
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace MathCTC
{
    public class TwilioCTC
    {
        public static void CTC_Master(string ascendcodes, string filePath, string destinationFolder)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var inputWorkSheet = package.Workbook.Worksheets["Payments and Deductions"];
                    var joinerandChangesSheet = package.Workbook.Worksheets["Joiner and Changes "];
                    int lastRow = inputWorkSheet.Dimension.End.Row;
                    int lastRow2 = joinerandChangesSheet.Dimension.End.Row;
                    int employee_Number = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "HR ID");
                    int payelementdescription = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Pay Element Description");
                    int town = Program.getColumnNumber(filePath, joinerandChangesSheet.ToString(), "PT Location");
                    int hrid = Program.getColumnNumber(filePath, joinerandChangesSheet.ToString(), "Hr id");
                    // int payelementdescription = Program.getColumnNumber(filePath, joinerandChangesSheet.ToString(), "Pay Element Short Code");
                    int witheffectfrom = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Start Date");
                    int annualctc = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Amount");
                    int payfreq = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Pay Frequency");

                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Ctc_Master");
                        outputWorksheet.Cells[1, 1].Value = "Employee Number";
                        outputWorksheet.Cells[1, 2].Value = "With effect From(YYYY - MM - DD)";
                        outputWorksheet.Cells[1, 3].Value = "Annual CTC";
                        outputWorksheet.Cells[1, 4].Value = "001 Basic Salary";
                        outputWorksheet.Cells[1, 5].Value = "Location";
                        outputWorksheet.Cells[1, 6].Value = "003 HRA";
                        outputWorksheet.Cells[1, 7].Value = "004 LTA";
                        outputWorksheet.Cells[1, 8].Value = "006 Cash Allowance";
                        outputWorksheet.Cells[1, 9].Value = "014 Employer NPS";
                        outputWorksheet.Cells[1, 10].Value = "031 Stipend";
                        outputWorksheet.Cells[1, 11].Value = "012 Connect";
                        outputWorksheet.Cells[1, 12].Value = "Employer PF";
                        HashSet<string> HRID = new HashSet<string>();
                        List<string> ptloc=new List<string>();

                        int row2 = 2;
                        for (int row = 2; row <= lastRow2; row++)
                        {
                            var cell = joinerandChangesSheet.Cells[row, hrid];
                            // Get the background color of the cell
                            var bgColor = cell.Style.Fill.BackgroundColor;
                            if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                            {
                                HRID.Add(joinerandChangesSheet.Cells[row, employee_Number].GetValue<string>().Replace(" ",""));
                                ptloc.Add(joinerandChangesSheet.Cells[row,town].Text);
                            }
                        }
                        row2 = 2;
                        foreach (string t in HRID)
                        {
                            outputWorksheet.Cells[row2, 1].Value = t;
                            outputWorksheet.Cells[row2,5].Value=ptloc[row2-2];
                            row2++;
                        }
                        row2 = 2;
                        foreach (string t in HRID) {
                            for (int row = 2; row <= lastRow; row++)
                            {
                                if (inputWorkSheet.Cells[row,employee_Number].Text.Replace(" ", "").Equals(t)&& inputWorkSheet.Cells[row, payelementdescription].Text.ToLower().Contains("basic") && inputWorkSheet.Cells[row, payfreq].Text.ToLower().Contains("annual")) 
                                {
                                    outputWorksheet.Cells[row2, 2].Value = inputWorkSheet.Cells[row, witheffectfrom].Text;
                                    outputWorksheet.Cells[row2, 3].Value = inputWorkSheet.Cells[row, annualctc].Text;
                                    double monthly40 = (inputWorkSheet.Cells[row, annualctc].GetValue<double>()) / 30.0;
                                    monthly40 = Math.Round(monthly40);
                                    outputWorksheet.Cells[row2, 7].Value = 0;
                                    outputWorksheet.Cells[row2, 9].Value = 0;
                                    outputWorksheet.Cells[row2, 10].Value = 0;
                                    outputWorksheet.Cells[row2, 12].Value = Math.Round((monthly40) * 12 / 100);
                                    outputWorksheet.Cells[row2, 7].Value = 0;
                                    outputWorksheet.Cells[row2, 4].Value = monthly40;
                                    outputWorksheet.Cells[row2, 6].Value = Math.Round(outputWorksheet.Cells[row2, 4].GetValue<double>() * 0.4);
                                    string city = outputWorksheet.Cells[row2, 5].GetValue<string>();
                                    city = Program.ShrinkString(city);
                                    bool delhi = city.Contains("delhi");
                                    bool mumbai = city.Contains("mumbai");
                                    bool maharashtra = city.Contains("maharashtra");
                                    bool tamilnadu = city.Contains("tamilnadu");
                                    bool newdelhi = city.Contains("newdelhi");
                                    bool chennai = city.Contains("chennai");
                                    bool kolkata = city.Contains("kolkata");
                                    bool calcutta = city.Contains("calcutta");
                                    if (delhi || mumbai || newdelhi || chennai || kolkata || maharashtra || tamilnadu)
                                    {
                                        outputWorksheet.Cells[row2, 6].Value = outputWorksheet.Cells[row2, 4].GetValue<double>() / 2;
                                    }
                                    outputWorksheet.Cells[row2, 6].Value = Math.Round(outputWorksheet.Cells[row2, 6].GetValue<double>());
                                    outputWorksheet.Cells[row2, 8].Value = Math.Round(((outputWorksheet.Cells[row2, 3].GetValue<double>()) / 12.0) - outputWorksheet.Cells[row2, 4].GetValue<double>() - outputWorksheet.Cells[row2, 6].GetValue<double>() - outputWorksheet.Cells[row2, 7].GetValue<double>() - outputWorksheet.Cells[row2, 9].GetValue<double>());
                                    row2++;
                                }
                            }
                        }
                        row2 = 2;
                        foreach (string t in HRID)
                        {
                            for (int row = 2; row <= lastRow; row++)
                            {
                                if (inputWorkSheet.Cells[row, employee_Number].Text.Replace(" ", "").Equals(t) && inputWorkSheet.Cells[row, payelementdescription].Text.ToLower().Contains("connect") && inputWorkSheet.Cells[row, payfreq].Text.ToLower().Contains("month"))
                                {
                                    outputWorksheet.Cells[row2, 11].Value = inputWorkSheet.Cells[row, annualctc].Text;
                                    row2++;
                                }
                            }
                        }
                        outputWorksheet.Column(3).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(4).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(6).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(7).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(8).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(9).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(10).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(11).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(12).Style.Numberformat.Format = "0.00";
                        string newFileName = Path.Combine(destinationFolder, "NEW_Joiners_Ctc_Master_" + Path.GetFileName(filePath));
                        FileInfo newFileInfo = new FileInfo(newFileName);
                        outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                        outputPackage.SaveAs(newFileInfo);
                        outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
                    }
                    Console.WriteLine("CTC Excel file created successfully!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

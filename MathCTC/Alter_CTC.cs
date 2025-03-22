using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace MathCTC
{
    public class Alter_CTC
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
                    int freqAmt = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Frequency Amount");
                    int payfreq = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Pay Frequency");

                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Ctc_Master");
                        outputWorksheet.Cells[1, 1].Value = "Employee Number";
                        outputWorksheet.Cells[1, 2].Value = "With effect From(YYYY - MM - DD)";
                        outputWorksheet.Cells[1, 3].Value = "Annual CTC";
                        outputWorksheet.Cells[1, 4].Value = "001 Basic Salary";
                        outputWorksheet.Cells[1, 5].Value = "003 HRA";
                        outputWorksheet.Cells[1, 6].Value = "004 LTA";
                        outputWorksheet.Cells[1, 7].Value = "005 Special Allowance";
                        outputWorksheet.Cells[1, 8].Value = "014 Employer NPS";
                        outputWorksheet.Cells[1, 9].Value = "Employer PF";
                        outputWorksheet.Cells[1, 10].Value = "011 Internet Allowance";
                        outputWorksheet.Cells[1, 11].Value = "004 Telephone Reimbursement";
                        outputWorksheet.Cells[1, 12].Value = "006 Meal Coupon";
                        outputWorksheet.Cells[1, 13].Value = "009 Fuel Card";

                        HashSet<string> HRID = new HashSet<string>();
                        List<string> ptloc = new List<string>();

                        int row2 = 2;
                        for (int row = 2; row <= lastRow2; row++)
                        {
                            var cell = joinerandChangesSheet.Cells[row, hrid];
                            // Get the background color of the cell
                            var bgColor = cell.Style.Fill.BackgroundColor;
                            if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                            {
                                HRID.Add(joinerandChangesSheet.Cells[row, employee_Number].GetValue<string>().Replace(" ", ""));
                                ptloc.Add(joinerandChangesSheet.Cells[row, town].Text);
                            }
                        }
                        row2 = 2;
                        foreach (string t in HRID)
                        {
                            outputWorksheet.Cells[row2, 1].Value = t;
                            //outputWorksheet.Cells[row2, 5].Value = ptloc[row2 - 2];
                            outputWorksheet.Cells[row2, 12]. Value = 2200.00;
                            row2++;
                        }
                        row2 = 2;
                        for (row2 = 2; row2 <= lastRow; row2++)
                        {
                            for (int row = 2; row <= lastRow; row++)
                            {
                                if (inputWorkSheet.Cells[row, employee_Number].Text.Equals(outputWorksheet.Cells[row2, 1].Text) && inputWorkSheet.Cells[row, payelementdescription].Text.ToLower().Contains("basic") && inputWorkSheet.Cells[row, payfreq].Text.ToLower().Contains("annual"))
                                {
                                    outputWorksheet.Cells[row2, 2].Value = inputWorkSheet.Cells[row, witheffectfrom].Text;
                                    outputWorksheet.Cells[row2, 3].Value = inputWorkSheet.Cells[row, annualctc].GetValue<double>();
                                    double monthly40 = (inputWorkSheet.Cells[row, annualctc].GetValue<double>()) / 30.0;
                                    double monthly50 = (inputWorkSheet.Cells[row, annualctc].GetValue<double>()) / 24.0;
                                    monthly40 = Math.Round(monthly40);

                                    outputWorksheet.Cells[row2, 8].Value = 0;
                                    outputWorksheet.Cells[row2, 10].Value = 0;
                                    outputWorksheet.Cells[row2, 9].Value = Math.Round((monthly40) * 12 / 100);
                                    outputWorksheet.Cells[row2, 4].Value = monthly50;
                                    outputWorksheet.Cells[row2, 5].Value = Math.Round(outputWorksheet.Cells[row2, 4].GetValue<double>() * 0.5);
                                    //string city = outputWorksheet.Cells[row2, 5].GetValue<string>();
                                    //city = Program.ShrinkString(city);
                                    //bool delhi = city.Contains("delhi");
                                    //bool mumbai = city.Contains("mumbai");
                                    //bool maharashtra = city.Contains("maharashtra");
                                    //bool tamilnadu = city.Contains("tamilnadu");
                                    //bool newdelhi = city.Contains("newdelhi");
                                    //bool chennai = city.Contains("chennai");
                                    //bool kolkata = city.Contains("kolkata");
                                    //bool calcutta = city.Contains("calcutta");
                                    //if (delhi || mumbai || newdelhi || chennai || kolkata || maharashtra || tamilnadu)
                                    //{
                                    //    outputWorksheet.Cells[row2, 5].Value = outputWorksheet.Cells[row2, 4].GetValue<double>() / 2;
                                    //}
                                    outputWorksheet.Cells[row2, 6].Value = outputWorksheet.Cells[row2, 4].GetValue<double>() * 8.33 / 100.0;
                                    outputWorksheet.Cells[row2, 7].Value = Math.Round((outputWorksheet.Cells[row2, 3].GetValue<double>() / 12.0) - outputWorksheet.Cells[row2, 4].GetValue<double>() - outputWorksheet.Cells[row2, 5].GetValue<double>()- outputWorksheet.Cells[row2, 6].GetValue<double>() - outputWorksheet.Cells[row2, 7].GetValue<double>() - outputWorksheet.Cells[row2, 12].GetValue<double>() - outputWorksheet.Cells[row2, 13].GetValue<double>());
                                }
                            }
                        }
                        for (row2 = 2; row2 <= lastRow; row2++)
                        {
                            for (int row = 2; row <= lastRow; row++)
                            {
                                if (inputWorkSheet.Cells[row, employee_Number].Text.Replace(" ", "").Equals(outputWorksheet.Cells[row2, 1].Text) && inputWorkSheet.Cells[row, payelementdescription].Text.ToLower().Contains("internet") && inputWorkSheet.Cells[row, payfreq].Text.ToLower().Contains("annual"))
                                {
                                    outputWorksheet.Cells[row2, 10].Value = inputWorkSheet.Cells[row, freqAmt].Text;
                                }
                            }
                        }
                        for (row2 = 2; row2 <= lastRow; row2++) {
                            for (int row = 2; row <= lastRow; row++)
                            {
                                if (inputWorkSheet.Cells[row, employee_Number].Text.Replace(" ", "").Equals(outputWorksheet.Cells[row2,1].Text) && inputWorkSheet.Cells[row, payelementdescription].Text.ToLower().Contains("telephone") && inputWorkSheet.Cells[row, payfreq].Text.ToLower().Contains("annual"))
                                {
                                    outputWorksheet.Cells[row2, 11].Value = inputWorkSheet.Cells[row, freqAmt].GetValue<double>();
                                   
                                }
                            }
                        }
                        for (row2 = 2; row2 <= lastRow; row2++)
                        {
                            for (int row = 2; row <= lastRow; row++)
                            {
                                if (inputWorkSheet.Cells[row, employee_Number].Text.Replace(" ", "").Equals(outputWorksheet.Cells[row2, 1].Text) && inputWorkSheet.Cells[row, payelementdescription].Text.ToLower().Contains("fuel") && inputWorkSheet.Cells[row, payfreq].Text.ToLower().Contains("annual"))
                                {
                                    outputWorksheet.Cells[row2, 13].Value = inputWorkSheet.Cells[row, freqAmt].GetValue<double>();

                                }
                            }
                        }
                        int endRow= outputWorksheet.Dimension.End.Row;
                        int endCol = outputWorksheet.Dimension.End.Column;
                        for (int row = 2; row <= endRow; row++)
                        {
                            for (int col = 4; col <= endCol; col++)
                            {
                                if (outputWorksheet.Cells[row, col].Text == "" || outputWorksheet.Cells[row, col].Text == null)
                                {
                                    outputWorksheet.Cells[row, col].Value = 0.0;
                                }
                            }
                        }
                        for (int col = 4; col <= endCol; col++) {
                            outputWorksheet.Column(col).Style.Numberformat.Format = "0.00";
                        }
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

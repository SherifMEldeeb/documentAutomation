using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Threading.Tasks;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Charts;


namespace SalesReport
{
    class Program
    {
        static void Main(string[] args)
        {
            Data data = new Data();
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            Console.WriteLine("Welcome To the Reporter...\n\n");
            string[] yearMonths  = new string[] { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };

            getData(ref data, ref yearMonths);
            
            //= ======================================================================================================== =
            ExcelFile report = new ExcelFile();
            ExcelWorksheet sheet = report.Worksheets.Add("Yearly Report");

            zabatellogo(ref sheet, data.compName);
            zabatHeader(ref sheet, ref data, yearMonths) ;

            dataEntry(ref sheet,ref data,ref yearMonths);

            ExcelWorksheet chartsh = report.Worksheets.Add(SheetType.Chart, "Chart");
            ExcelChart chart = chartsh.Charts.Add(ChartType.Pie, 0, 0, 0, 0, LengthUnit.Centimeter);
            chart.DataLabels.LabelPosition = DataLabelPosition.InsideEnd;
            // chart data.
            // period +1 for header.
            chart.SelectData(sheet.Cells.GetSubrangeRelative(7 , 1, 2, data._period+1),true);
            
            report.Save(Globals.savePath + data.compName +".xlsx");

            Console.Write("Succesfully Created... Do you want to make a pdf copy?");
            char pdf = char.Parse(Console.ReadLine());
            if (char.ToLower(pdf) == 'y')
            {
                sheet.PrintOptions.Portrait = true;
                sheet.PrintOptions.PaperType = PaperType.A4;
                
                report.Save(Globals.savePath + data.compName + ".pdf", new PdfSaveOptions() { SelectionType = SelectionType.EntireFile });

            }
            System.Diagnostics.Process.Start("explorer.exe", Globals.savePath);
         }


        static void getData(ref Data data, ref string[] months)
        {
            Console.Write("Enter Company Name: ");
            data.compName = Console.ReadLine();
            
            Console.Write("\nEnter the Begin Month: ");
            data.bMonth = int.Parse(Console.ReadLine());

            Console.Write("\nEnter the End Month: ");
            data.eMonth = int.Parse(Console.ReadLine());

            data.sales = new double[data._period];
            Console.WriteLine("\nEnter Sales Of Month ");
            for (int i = data.bMonth; i <= data.eMonth; i++)
            {
                Console.Write("- {0}: ", months[i - 1]);
                data.sales[i - data.bMonth] = double.Parse(Console.ReadLine());
            }
        }
        static void zabatellogo(ref ExcelWorksheet sheet, string compName)
        {
            sheet.Pictures.Add(Globals.resourcesPath + "img.jpg", "A1", "B4");
            sheet.Cells.GetSubrangeAbsolute(0, 2, 3, 3).Merged = true;
            sheet.Cells["C2"].Value = compName;
            sheet.Cells["C2"].Style.Font.Weight = ExcelFont.BoldWeight;
            sheet.Cells["C2"].Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            sheet.Cells["C2"].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet.Cells["C2"].Style.Font.Size = 20 * 20;
        }

        static void zabatHeader(ref ExcelWorksheet sheet, ref Data data, string[] yearMonths)
        {
            sheet.Cells.GetSubrangeAbsolute(4, 0, 5, 3).Merged = true;
            sheet.Cells.GetSubrangeAbsolute(6, 0, 6, 3).Merged = true;
            sheet.Cells.GetSubrangeAbsolute(4, 0, 6, 3).Style.FillPattern.SetPattern(FillPatternStyle.Solid, SpreadsheetColor.FromName(ColorName.Purple), SpreadsheetColor.FromName(ColorName.Purple));

            for (int i = 4; i < 7; i += 2)
            {
                sheet.Cells[i, 0].Value = i == 4 ? "Monthly Report" : yearMonths[data.bMonth - 1] + " - " + yearMonths[data.eMonth - 1] + " /" + System.DateTime.Now.Year;
                sheet.Cells[i, 0].Style.Font.Size = i == 4 ? 20 * 20 : 10 * 20;
                sheet.Cells[i, 0].Style.Font.Color = i == 4 ? SpreadsheetColor.FromName(ColorName.Yellow) : SpreadsheetColor.FromName(ColorName.White);
                sheet.Cells[i, 0].Style.Font.Weight = ExcelFont.BoldWeight;
                sheet.Cells[i, 0].Style.VerticalAlignment = VerticalAlignmentStyle.Center;
                sheet.Cells[i, 0].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            }
        }

        static void dataEntry(ref ExcelWorksheet sheet, ref Data data, ref string[] yearMonths)
        {
            string[] header = new string[] { "Month", "Sales" };
            const int pos = 7;
            // +7 for the merged cells , +1 for header.
            for (int i = pos; i < data._period + 1 +pos; i++)
            {
                for (int j = 1, k = header.Length+1; j < k; j++)
                {
                    // For Header.
                    if (i == pos)
                    {
                        sheet.Cells[i, j].Value = header[j-1];
                        sheet.Cells[i, j].Style.Font.Weight = ExcelFont.BoldWeight;
                        sheet.Columns[j].Width = (int)LengthUnitConverter.Convert(3, LengthUnit.Centimeter, LengthUnit.ZeroCharacterWidth256thPart);
                    }
                    else
                    {
                        if (j == 1 && i <= data._period + pos)
                        {
                            // Minus one for zero Index and another because of j.
                            sheet.Cells[i, j].Value = yearMonths[(data.bMonth - 2) + i -pos];
                            sheet.Columns[j].Width = (int)LengthUnitConverter.Convert(3, LengthUnit.Centimeter, LengthUnit.ZeroCharacterWidth256thPart);

                        }
                        else if (j == 2 && i <= data._period + pos)
                        {
                            sheet.Cells[i,j].Style.NumberFormat = "\"$\"#,##0";
                            sheet.Cells[i, j].Value = data.sales[(i - 1 -pos)];
                        }
                    }
                    sheet.Cells[i,j].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                }
            }
        }

      
    }
}

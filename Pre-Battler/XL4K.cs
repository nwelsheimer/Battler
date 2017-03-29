using System;
using System.IO;
using System.Xml;
using OfficeOpenXml;
using System.Data;
using System.Drawing;
using OfficeOpenXml.Style;
using General;

namespace Pre_Battler
{
    class xl4k
    {
        /// xl4k - A more robust less stable piece of shit - NDW 3/2/2017
        /// 
        /// Rewrite of the old Joe method of using HTML tags create a spreadsheet. This one isn't very modular, at
        /// least not yet. The ep plus dll lets us write xml direct to an xlsx file, which gives tons of flexibility to creating
        /// dynamic excel files.
        /// TODO
        /// Someone needs to come along and figure out a way to design a spreadsheet template and store in the database, then pull
        /// the template a run time, feed data in to it, and generate output.


        public static string GenerateExport(int sessionId, string fileName = "", string fullPath = "")
        {
            const int startRow = 5;

            if (fullPath != "")
            {
                // lets connect to the server for some data first
                DataTable export = new DataTable();
                export = Global.GetData("usp_PB_ExcelExportSummarySKU @sessionId=" + sessionId).Tables[0];

                //Delete the old file if we are overwriting
                string file = fullPath;
                if (File.Exists(file)) File.Delete(file);
                FileInfo newFile = new FileInfo(file);

                //Don't always know how many regions there will be
                int totalColumns = export.Columns.Count;
                Char lastColumn = Global.LetterToNum(totalColumns+1);

                // ok, we can run the real code now
                using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                {
                    // get handle to the existing worksheet
                    ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets.Add("Summary by SKU");
                    ExcelWorksheet storeDetail = xlPackage.Workbook.Worksheets.Add("Store Detail");
                    var namedStyle = xlPackage.Workbook.Styles.CreateNamedStyle("HyperLink");   //This one is language dependent
                    namedStyle.Style.Font.UnderLine = true;
                    namedStyle.Style.Font.Color.SetColor(Color.Blue);

                    //First worksheet, should have named this better... TAB1
                    if (worksheet != null)
                    {
                        int row = startRow;
                        //Create Headers and format them 
                        worksheet.Cells["A1"].Value = "Summary by SKU Report";
                        using (ExcelRange r = worksheet.Cells["A1:"+lastColumn+"1"])
                        {
                            r.Merge = true;
                            r.Style.Font.SetFromFont(new Font("Britannic Bold", 22, FontStyle.Italic));
                            r.Style.Font.Color.SetColor(Color.White);
                            r.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
                            r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(23, 55, 93));
                        }
                        worksheet.Cells["A2"].Value = "Session: " + fileName;
                        using (ExcelRange r = worksheet.Cells["A2:"+lastColumn+"2"])
                        {
                            r.Merge = true;
                            r.Style.Font.SetFromFont(new Font("Britannic Bold", 18, FontStyle.Italic));
                            r.Style.Font.Color.SetColor(Color.Black);
                            r.Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
                        }

                        worksheet.Cells["A4"].Value = "Site";
                        worksheet.Cells["B4"].Value = "SKU Desc";
                        worksheet.Cells["C4"].Value = "SKU Size";
                        worksheet.Cells["D4"].Value = "SKU Number";
                        for (int i=5; i<=totalColumns; i++)
                        {
                            worksheet.Cells[Global.LetterToNum(i) + "4"].Value = export.Columns[i-1].ColumnName;
                        }
                        worksheet.Cells[Global.LetterToNum(totalColumns+1)+"4"].Value = "Total";
                        worksheet.Cells["A4:" + lastColumn + "4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells["A4:" + lastColumn + "4"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
                        worksheet.Cells["A4:" + lastColumn + "4"].Style.Font.Bold = true;


                        // get the data and fill rows 5 onwards
                        foreach (DataRow dr in export.Rows)
                        {
                            int col = 1;
                            // our query has the columns in the right order, so simply
                            // iterate through the columns
                            for (int i = 0; i < totalColumns; i++)
                            {
                                // do not bother filling cell with blank data (also useful if we have a formula in a cell)
                                worksheet.Cells[row, i+1].Value = dr[i];
                                col++;
                            }
                            row++;
                        }
                        
                        worksheet.Cells[startRow, 3, row - 1, totalColumns].Style.Numberformat.Format = "#,##0";
                        worksheet.Cells[startRow, totalColumns+1, row - 1, totalColumns+1].Style.Numberformat.Format = "#,##0";

                        worksheet.Cells[startRow, totalColumns + 1, row - 1, totalColumns + 1].FormulaR1C1 = "=SUM(RC["+(2-totalColumns)+"]:RC[-1])";

                        //Set column width
                        worksheet.Column(2).Width = 48;
                        worksheet.Column(3).Width = 15;
                        worksheet.Column(4).Width = 15;

                        // lets set the header text 
                        worksheet.HeaderFooter.OddHeader.CenteredText = fileName+ " Pre-battle report";
                        // add the page number to the footer plus the total number of pages
                        worksheet.HeaderFooter.OddFooter.RightAlignedText =
                            string.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
                        // add the sheet name to the footer
                        worksheet.HeaderFooter.OddFooter.CenteredText = ExcelHeaderFooter.SheetName;
                        // add the file path to the footer
                        worksheet.HeaderFooter.OddFooter.LeftAlignedText = ExcelHeaderFooter.FilePath + ExcelHeaderFooter.FileName;
                    }
                    //Second, create detail tab. Let the shoehorning begin TAB2
                    if (storeDetail != null)
                    {
                        //flush the table, reload
                        export = new DataTable();
                        string query = "usp_PB_ExcelExportDetail @sessionId=" + sessionId;
                        export = Global.GetData(query).Tables[0];

                        //reset globals
                        totalColumns = export.Columns.Count;
                        lastColumn = Global.LetterToNum(totalColumns);
                        int row = startRow;

                        //Create Headers and format them 
                        storeDetail.Cells["A1"].Value = "Detail by Store Report";
                        using (ExcelRange r = storeDetail.Cells["A1:" + lastColumn + "1"])
                        {
                            r.Merge = true;
                            r.Style.Font.SetFromFont(new Font("Britannic Bold", 22, FontStyle.Italic));
                            r.Style.Font.Color.SetColor(Color.White);
                            r.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
                            r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(23, 55, 93));
                        }
                        storeDetail.Cells["A2"].Value = "Session: " + fileName;
                        using (ExcelRange r = storeDetail.Cells["A2:" + lastColumn + "2"])
                        {
                            r.Merge = true;
                            r.Style.Font.SetFromFont(new Font("Britannic Bold", 18, FontStyle.Italic));
                            r.Style.Font.Color.SetColor(Color.Black);
                            r.Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                            r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
                        }
                        
                        for (int i = 1; i <= totalColumns; i++)
                        {
                            storeDetail.Cells[Global.LetterToNum(i) + "4"].Value = export.Columns[i - 1].ColumnName;
                        }
                        storeDetail.Cells["A4:" + lastColumn + "4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        storeDetail.Cells["A4:" + lastColumn + "4"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
                        storeDetail.Cells["A4:" + lastColumn + "4"].Style.Font.Bold = true;


                        // get the data and fill rows 5 onwards
                        foreach (DataRow dr in export.Rows)
                        {
                            int col = 1;
                            // our query has the columns in the right order, so simply
                            // iterate through the columns
                            for (int i = 0; i < totalColumns; i++)
                            {
                                // do not bother filling cell with blank data (also useful if we have a formula in a cell)
                                storeDetail.Cells[row, i + 1].Value = dr[i];
                                col++;
                            }
                            row++;
                        }

                        storeDetail.Cells[startRow, 7, row - 1, totalColumns].Style.Numberformat.Format = "#,##0";
                        storeDetail.Cells[startRow, 8, row - 1, 8].Style.Numberformat.Format = "#,##0.0";

                        //Set column width
                        storeDetail.Column(1).Width = 20;
                        storeDetail.Column(2).Width = 12;
                        storeDetail.Column(3).Width = 26;

                        // lets set the header text 
                        storeDetail.HeaderFooter.OddHeader.CenteredText = fileName + " Pre-battle report";
                        // add the page number to the footer plus the total number of pages
                        storeDetail.HeaderFooter.OddFooter.RightAlignedText =
                            string.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
                        // add the sheet name to the footer
                        storeDetail.HeaderFooter.OddFooter.CenteredText = ExcelHeaderFooter.SheetName;
                        // add the file path to the footer
                        storeDetail.HeaderFooter.OddFooter.LeftAlignedText = ExcelHeaderFooter.FilePath + ExcelHeaderFooter.FileName;
                    }
                    //Third, if the shoe don't fit, get a bigger hammer TAB3
                    // set some core property values
                    xlPackage.Workbook.Properties.Title = fileName;
                    xlPackage.Workbook.Properties.Subject = fileName + " Pre-battle report";
                    xlPackage.Workbook.Properties.Keywords = fileName + " Pre-battle report";
                    xlPackage.Workbook.Properties.Category = "ExcelPackage";
                    xlPackage.Workbook.Properties.Comments = "Standard output from pre-battler";

                    // save the new spreadsheet
                    xlPackage.Save();
                }
            
                return newFile.FullName;
            } else
            { return "failed"; }
        }
    }
}

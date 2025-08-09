

using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace uploadfile
{

    public class CreateDocx
    {
       public void CreateWordDocx(string filePath,string wordOutputPath)
        {
            
                var wordApp = new Word.Application();
                var document = wordApp.Documents.Add();

                // Start Excel application
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
                Excel._Worksheet worksheet = workbook.Sheets[1]; // First worksheet
                Excel.Range usedRange = worksheet.UsedRange;
            try
            {
                // Step 3: Add a title
                Word.Paragraph para = document.Paragraphs.Add();
                para.Range.Text = "Processed Excel Data";
                para.Range.Font.Size = 16;
                para.Range.Font.Bold = 1;
                para.Format.SpaceAfter = 24;
                para.Range.InsertParagraphAfter();

                // Get number of rows and columns
                int rows = usedRange.Rows.Count;
                int cols = usedRange.Columns.Count;


                Word.Table table = document.Tables.Add(para.Range, rows, cols);
                table.Borders.Enable = 1;

                for (int r = 1; r < rows; r++)
                {
                    for (int c = 1; c < cols; c++)
                    {
                        //table.Cell(r + 1, c + 1).Range.Text = excelData[r, c];

                        Excel.Range cell = (Excel.Range)usedRange.Cells[r, c];
                        string cellValue = cell.Text?.ToString(); // .Value2 can also be used
                                                                  //Console.Write($"{cellValue}\t");
                        table.Cell(r + 1, c + 1).Range.Text = cellValue;
                    }
                }


                workbook.Close(false);
                excelApp.Quit();

                // Release COM objects to avoid memory leaks
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                // Step 5: Save Word document
                document.SaveAs2(wordOutputPath);
                document.Close();
                wordApp.Quit();
            }
            catch (Exception ex)
            {
                workbook.Close(false);
                excelApp.Quit();

                // Release COM objects to avoid memory leaks
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                // Step 5: Save Word document
                document.SaveAs2(wordOutputPath);
                document.Close();
                wordApp.Quit();
            }

        }
    }

}

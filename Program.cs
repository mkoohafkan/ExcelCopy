using System;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelCopy
{

  static class ExcelCopyRanges
  {

    static void Main(string[] args)
    {
      string srcDir = Path.GetFullPath(args[0]);
      string destDir = Path.GetFullPath(args[1]);
      string templatePath = Path.GetFullPath(args[2]);
      string fromSheet = @"Sheet1";
      string toSheet = @"VISIT";
      string[] ranges = new string[8];
      ranges[0] = @"F4:I6";
      ranges[1] = @"P4:S6";
      ranges[2] = @"Z4:AC6";
      ranges[3] = @"H10:AE16";
      ranges[4] = @"H20:M20";
      ranges[5] = @"Q20:V20";
      ranges[6] = @"Z20:AE20";
      ranges[7] = @"D22:AE24";
      // get list of files from source directory
      string[] fileNames = Directory.GetFiles(srcDir, "*.xlsx",
        SearchOption.TopDirectoryOnly).Select(Path.GetFileName).ToArray();
      Excel.Application excelApplication = new Excel.Application();
      excelApplication.Application.DisplayAlerts = false;
      foreach (string fileName in fileNames)
      {
        Console.WriteLine(String.Join(" ", "Copying", fileName));
        // open source workbook
        Excel.Workbook srcworkBook = excelApplication.Workbooks.Open(Path.Combine(srcDir, fileName));
        Excel.Worksheet srcworkSheet = srcworkBook.Worksheets[fromSheet];
        // open template
        Excel.Workbook destworkBook = excelApplication.Workbooks.Open(templatePath, 0, false);
        Excel.Worksheet destworkSheet = destworkBook.Worksheets[toSheet];
        // copy ranges
        foreach (string range in ranges)
        {
            Excel.Range from = srcworkSheet.Range[range];
            Excel.Range to = destworkSheet.Range[range];
          try
          {
            from.Copy(to);
          }
          catch 
          {
            Console.WriteLine(String.Join(" ", "    Could not copy range", range));
          }
        }
        // save template to new workbook
        srcworkBook.Close();
        destworkBook.SaveAs(Path.Combine(destDir, fileName));
        destworkBook.Close();
      }
      excelApplication.Quit();
    }
  }
}

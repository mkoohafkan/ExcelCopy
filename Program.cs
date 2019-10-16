using System;
using System.Collections;
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
      string templateType = args[3];
      string fromSheet = null; 
      string toSheet = @"ENTRY";
      string[] ranges = null;
      if (templateType == "VISIT")
      {
        Array.Resize(ref ranges, 8);
        ranges[0] = @"F4:I6";
        ranges[1] = @"P4:S6";
        ranges[2] = @"Z4:AC6";
        ranges[3] = @"H10:AE16";
        ranges[4] = @"H20:M20";
        ranges[5] = @"Q20:V20";
        ranges[6] = @"Z20:AE20";
        ranges[7] = @"D22:AE24";
      }
      else if(templateType == "PRECALIBRATION")
      {
        Array.Resize(ref ranges, 27);
        ranges[0] = @"O6";
        ranges[1] = @"D8";
        ranges[2] = @"H8";
        ranges[3] = @"P8";
        ranges[4] = @"AA8:AA9";
        ranges[5] = @"H9";
        ranges[6] = @"F10";
        ranges[7] = @"A17:A18";
        ranges[11] = @"F14:F18";
        ranges[12] = @"K14:K18";
        ranges[13] = @"V13:V18";
        ranges[14] = @"AA13:AA18";
        ranges[15] = @"J23:J36";
        ranges[16] = @"J38";
        ranges[17] = @"M23:M36";
        ranges[18] = @"P25:P38";
        ranges[19] = @"X23:X30";
        ranges[20] = @"AB31";
        ranges[21] = @"X32:X34";
        ranges[22] = @"AC35:AC36";
        ranges[23] = @"X38";
        ranges[24] = @"K41:K44";
        ranges[25] = @"V42:V44";
        ranges[26] = @"X41";
      }
      else if (templateType == "POSTCHECK")
      {
        Array.Resize(ref ranges, 24);
        ranges[0] = @"O6";
        ranges[1] = @"D8";
        ranges[2] = @"H8";
        ranges[3] = @"P8";
        ranges[4] = @"AA8:AA9";
        ranges[5] = @"K9";
        ranges[6] = @"J10";
        ranges[7] = @"O10";
        ranges[8] = @"K11";
        ranges[9] = @"S11";
        ranges[10] = @"A18:A19";
        ranges[11] = @"F15:F19";
        ranges[12] = @"K15:K19";
        ranges[13] = @"V14:V19";
        ranges[14] = @"AA14:AA19";
        ranges[15] = @"I23";
        ranges[16] = @"T23";
        ranges[17] = @"AA23";
        ranges[18] = @"J26:J40";
        ranges[19] = @"M26:M40";
        ranges[20] = @"X35:X37";
        ranges[21] = @"AC38:AC39";
        ranges[22] = @"X40";
        ranges[23] = @"A43";
      }
      // get list of files from source directory
      string[] fileNames = Directory.GetFiles(srcDir, "*.xlsx",
        SearchOption.TopDirectoryOnly).Select(Path.GetFileName).ToArray();
      Excel.Application excelApplication = new Excel.Application();
      excelApplication.Application.DisplayAlerts = false;
      try
      {
        foreach (string fileName in fileNames)
        {
          Console.WriteLine(String.Join(" ", "Copying", fileName));
          // open source workbook
          Excel.Workbook srcworkBook = excelApplication.Workbooks.Open(Path.Combine(srcDir, fileName));
          // get sheet names
          Excel.Sheets srcSheets = srcworkBook.Sheets;
          ArrayList srcSheetNames = new ArrayList();
          foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in srcSheets)
          {
            srcSheetNames.Add(sheet.Name);
          }
          if(srcSheetNames.Contains("VISIT"))
          {
            fromSheet = @"VISIT";
          }
          else if (srcSheetNames.Contains("ENTRY"))
          {
            fromSheet = @"ENTRY";
          }
          else if (srcSheetNames.Contains("PRECHECK"))
          {
            fromSheet = @"PRECHECK";
          }
          else if (srcSheetNames.Contains("POSTCHECK"))
          {
            fromSheet = @"POSTCHECK";
          }
          else if (srcSheetNames.Contains("Sheet1"))
          {
            fromSheet = @"Sheet1";
          } else {
            Console.WriteLine("Could not detect input sheet");
          }
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
              //from.Copy(to);
              from.Copy();
              to.PasteSpecial(Excel.XlPasteType.xlPasteValues);
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
          Console.WriteLine(String.Join(" ", "Sucessfully copied", fileName));
        }
      }
      catch {
        Console.WriteLine("An error occurred.");
      }
      finally
      {
        excelApplication.Quit();
      }
    }
  }
}

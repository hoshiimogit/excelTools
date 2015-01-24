using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTools
{
    class Test
    {
        static void Main(string[] args)
        {
            //準備
            var excelApp = new Excel.Application();
            excelApp.Visible = true; //for debug 画面表示したくなければfalse。
            excelApp.DisplayAlerts = false;

            var wkbooks = excelApp.Workbooks;
            var wkbook = (Excel.Workbook)wkbooks.Open(@"c:\testData\sample.xlsx",
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //テストパターン                                                                        画像の形 セル<>画像サイズ  フィット方向（結果）
            //-----------------------------------------------------------------------------------   --------  -----------       ----------- 
            wkbook = ExcelTools.InsertPicture(wkbook, "sheet3", "C2", "E6", @"c:\testData\a1.jpg");    //縦長    画像大             縦          
            wkbook = ExcelTools.InsertPicture(wkbook, "sheet3", "G2", "H12", @"c:\testData\a1.jpg");   //縦長    画像大             横          
            wkbook = ExcelTools.InsertPicture(wkbook, "sheet3", "C16", "E20", @"c:\testData\a2.jpg");  //横長    画像大             縦          
            wkbook = ExcelTools.InsertPicture(wkbook, "sheet3", "G16", "H26", @"c:\testData\a2.jpg");  //横長    画像大             横          
            wkbook = ExcelTools.InsertPicture(wkbook, "sheet3", "C29", "E33", @"c:\testData\a3.jpg");  //縦長    セル大             縦          
            wkbook = ExcelTools.InsertPicture(wkbook, "sheet3", "G29", "H39", @"c:\testData\a3.jpg");  //縦長    セル大             横          
            wkbook = ExcelTools.InsertPicture(wkbook, "sheet3", "C43", "E47", @"c:\testData\a4.jpg");  //横長    セル大             縦          
            wkbook = ExcelTools.InsertPicture(wkbook, "sheet3", "G43", "H53", @"c:\testData\a4.jpg");  //横長    セル大             横          

            //結果保存
            wkbook.SaveAs(@"c:\testData\result.xlsx",
                Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing,
                Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //後始末
            wkbook.Close(false, Type.Missing, Type.Missing);
            excelApp.Quit();
        }
    }
}

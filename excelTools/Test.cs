
namespace ExcelTools
{
    class Test
    {
        static void Main(string[] args)
        {
            //テストパターン
            ExcelTools.InsertPicture("c:\\excel\\sample.xlsx", "sheet3", "C2", "E6", "c:\\excel\\a1.jpg");  //縦フィット、縦長画像、画像大
            ExcelTools.InsertPicture("c:\\excel\\sample.xlsx", "sheet3", "G2", "H12", "c:\\excel\\a1.jpg"); //横フィット、縦長画像、画像大
            ExcelTools.InsertPicture("c:\\excel\\sample.xlsx", "sheet3", "C16", "E20", "c:\\excel\\a2.jpg");//縦フィット、横長画像、画像大
            ExcelTools.InsertPicture("c:\\excel\\sample.xlsx", "sheet3", "G16", "H26", "c:\\excel\\a2.jpg");//横フィット、横長画像、画像大
            ExcelTools.InsertPicture("c:\\excel\\sample.xlsx", "sheet3", "C29", "E33", "c:\\excel\\a3.jpg");//縦フィット、縦長画像、画像小
            ExcelTools.InsertPicture("c:\\excel\\sample.xlsx", "sheet3", "G29", "H39", "c:\\excel\\a3.jpg");//横フィット、縦長画像、画像小
            ExcelTools.InsertPicture("c:\\excel\\sample.xlsx", "sheet3", "C43", "E47", "c:\\excel\\a4.jpg");//縦フィット、横長画像、画像小
            ExcelTools.InsertPicture("c:\\excel\\sample.xlsx", "sheet3", "G43", "H53", "c:\\excel\\a4.jpg");//横フィット、横長画像、画像小

        }
    }
}

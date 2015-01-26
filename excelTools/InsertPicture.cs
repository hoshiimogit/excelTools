using Excel = Microsoft.Office.Interop.Excel;
using Core = Microsoft.Office.Core;

namespace ExcelTools
{
    class ExcelTools
    {
        /// <summary>
        /// 指定したセルに画像を縦横比を維持したまま中央かつ最大サイズで貼り付ける
        /// Insert Picture in Excel Automatically Sized to Fit Cells
        /// </summary>
        /// <param name="wkbook">excelワークブックオブジェクト</param>
        /// <param name="sheetName">シート名</param>
        /// <param name="topCell">貼り付ける場所（左上セル）</param>
        /// <param name="bottomCell">貼り付ける場所（右下セル）</param>
        /// <param name="picPath">画像ファイルのパス</param>
        /// <returns>更新されたexcelワークブックオブジェクト</returns>
        public static Excel.Workbook InsertPicture(Excel.Workbook wkbook, string sheetName, string topCell, string bottomCell, string picPath)
        {
            var sheets = wkbook.Worksheets;
            var wksheet = (Excel.Worksheet)sheets[sheetName];
            wksheet.Select();

            wksheet.Unprotect("azunyan"); //シート保護の解除

            var range = wksheet.get_Range(topCell, bottomCell);
            range.Select();

            var rangeLeft = float.Parse(range.Left.ToString()); //セルのx座標
            var rangeTop = float.Parse(range.Top.ToString()); //セルのy座標
            var rangeWidth = float.Parse(range.Width.ToString()); //セルの幅
            var rangeHeight = float.Parse(range.Height.ToString()); //セルの高さ

            var shapes = wksheet.Shapes;

            var shape = shapes.AddPicture(picPath,
                Core.MsoTriState.msoFalse, /* false:埋め込む,true:リンクする */
                Core.MsoTriState.msoTrue,  /* false:保存しない,true:保存する */
                rangeLeft, rangeTop, (float)0.0, (float)0.0); //位置だけ指定し、サイズは以下プロパティで再設定する

            shape.ScaleWidth((float)1.0 /* 拡大比率 */, Core.MsoTriState.msoTrue); //一端、実サイズで展開
            shape.ScaleHeight((float)1.0 /* 拡大比率 */, Core.MsoTriState.msoTrue); //一端、実サイズで展開

            var picWidth = shape.Width;//画像の幅
            var picHeight = shape.Height;//画像の高さ
            var picRatio = picWidth / picHeight; //画像の横長比
            var rangeRatio = rangeWidth / rangeHeight; //セルの横長比

            if (rangeRatio > picRatio)
            {
                //セルの方が横長比が大きい場合 
                shape.Height = rangeHeight;// 画像の高さ＝セルの高さ
                shape.Width *= (rangeHeight / picHeight);// 画像の幅＝画像の高さの拡大率に合わせる
                shape.IncrementLeft(rangeWidth / 2 - shape.Width / 2); //横方向に移動する
            }
            else
            {
                //画像の方が横長比が大きい場合
                shape.Width = rangeWidth;//画像の幅＝セルの幅
                shape.Height *= (rangeWidth / picWidth);//画像の高さ＝画像の幅の拡大率に合わせる
                shape.IncrementTop(rangeHeight / 2 - shape.Height / 2); //縦方向に移動する
            }

            wksheet.Protect("azunyan"); //シート保護

            return wkbook;
        }
    }
}

using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Data.Common;

public class QuotationWorksheet : WorksheetBase
{
    // 寫入標題行
    public override string[] header => new string[]
    {
        "項次", "工作項目", "工作說明", "數量", "總價(NT$)", "備註"
    };

    public override int[] widthAlignment => new int[] 
    { 
        10, 29, 54, 9, 18, 12 
    };

    public override int lastRow { get; set; }

    public QuotationWorksheet(ExcelWorksheet sheet) : base(sheet)
    {
        lastRow = 1;
    }

    // 覆寫 Dispose 方法
    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            // 釋放子類別特定的資源（如果有）
        }

        // 呼叫基類的 Dispose 方法
        base.Dispose(disposing);
    }

    public override void FormatPhase()
    {
        // 格式化內容
        FormatCells(sheet.Cells[$"{startCol}{phaseTitleRow + lastRow}:{endCol}{phaseEndRow + lastRow - 1}"], isThin: true, isBorder: true);
        // 使用 RGB 自定義顏色格式化階段標題
        Color titleBgColor = Color.FromArgb(198, 224, 180);  // 淺綠色
        MergeAndAlign(phaseTitleRow + lastRow, 1, phaseTitleRow + lastRow, Array.IndexOf(header, "工作說明") + 1);
        FormatCells(sheet.Cells[$"{startCol}{phaseTitleRow + lastRow}:{endCol}{phaseTitleRow + lastRow}"], bgColor: titleBgColor, isTitle: true);
        sheet.Cells[phaseTitleRow + lastRow, Array.IndexOf(header, "數量") + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        sheet.Cells[phaseTitleRow + lastRow, Array.IndexOf(header, "總價(NT$)") + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;  // 內文靠右
    }

    public void PhaseSumPrice(int phaseCount)
    {
        // 使用 currentRow 來決定列數
        sheet.Cells[phaseTitleRow + lastRow, Array.IndexOf(header, "數量") + 1].Value = 1;
        sheet.Cells[phaseTitleRow + lastRow, Array.IndexOf(header, "總價(NT$)") + 1].Formula = $"SUM({projCostRefSheet}!L{referProjCostRow}:L{referProjCostRow + phaseCount})";
        sheet.Cells[phaseTitleRow + lastRow, Array.IndexOf(header, "總價(NT$)") + 1].Style.Numberformat.Format = "$#,##0";
    }


    public void WriteCostSumFormula(int column)
    {        
        sheet.Cells[currentRow, column].Style.Numberformat.Format = "$#,##0";
        sheet.Cells[currentRow, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right; // 內文靠右
        sheet.Cells[currentRow, column].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
    }

    // 寫入標題行
    public override void WriteAndFormatHeader(int startRow, int endRow)
    {
        FormatCells(sheet.Cells[$"{startCol}{startRow}:{endCol}{endRow - 1}"], isHeader: true, isHair: true);
    }

    public override void WriteFooter()
    {
        string[] columnNames = ["客戶名稱", "專案名稱", "業務部門", "業務代表", "電子信箱", "電話分機", "報價日期", "報價單號", "技術部門", "部門代表", "電子信箱", "電話分機"];
        sheet.Cells[currentRow, 1].Value = $"總計";
        sheet.Cells[currentRow, Array.IndexOf(header, "總價(NT$)") + 1].Formula = $"{projCostRefSheet}!L{referProjCostRow}";
        sheet.Cells[currentRow, Array.IndexOf(header, "總價(NT$)") + 1].Style.Numberformat.Format = "$#,##0";
        sheet.Cells[currentRow, Array.IndexOf(header, "總價(NT$)") + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        Color footerBgColor = Color.FromArgb(248, 203, 173);    // 淺橘紅色
        FormatCells(sheet.Cells[$"{startCol}{currentRow}:{endCol}{currentRow}"], bgColor: footerBgColor, isBorder: true);
        // 深藍色
        footerBgColor = Color.FromArgb(7, 79, 105);     // 深藍色
        sheet.Cells[currentRow + 2, 1].Value = "專案成員人天分配：";
        MergeAndAlign(currentRow + 2, 1, currentRow + 2, 3, isCenter: false);
        FormatCells(sheet.Cells[$"A{currentRow + 2}:C{currentRow + 2}"], bgColor: footerBgColor, fontColor: Color.White, isBorder: true);
        sheet.Cells[currentRow + 3, 1].Value = "處理人員";
        sheet.Cells[currentRow + 3, 2].Value = "天數";
        sheet.Cells[currentRow + 3, 3].Value = "角色";
        // 淺藍色
        sheet.Cells[currentRow + 4, 2].Formula = $"={projCostRefSheet}!F{referProjCostRow}";
        sheet.Cells[currentRow + 4, 3].Value = "專案經理";
        FormatCells(sheet.Cells[currentRow + 4, 2], isRight: true);
        sheet.Cells[currentRow + 5, 2].Formula = $"={projCostRefSheet}!H{referProjCostRow}";
        sheet.Cells[currentRow + 5, 3].Value = "部署者|架構師";
        FormatCells(sheet.Cells[currentRow + 5, 2], isRight: true);
        sheet.Cells[currentRow + 6, 2].Formula = $"={projCostRefSheet}!J{referProjCostRow}";
        sheet.Cells[currentRow + 6, 3].Value = "開發者";
        FormatCells(sheet.Cells[currentRow + 6, 2], isRight: true);
        FormatCells(sheet.Cells[$"A{currentRow + 3}:C{currentRow + 6}"], isThin: true, isBorder: true);
        // 淡藍色
        footerBgColor = Color.FromArgb(202, 237, 251);  // 淡藍色
        FormatCells(sheet.Cells[$"C{currentRow + 3}:C{currentRow + 6}"], bgColor: footerBgColor);
        // 淺藍色
        footerBgColor = Color.FromArgb(97, 203, 243);   // 淺藍色
        FormatCells(sheet.Cells[$"A{currentRow + 3}:C{currentRow + 3}"], bgColor: footerBgColor, isBorder: true);
        // 深藍色
        footerBgColor = Color.FromArgb(7, 79, 105);     // 深藍色
        sheet.Cells[currentRow + 8, 1].Value = "預估專案期間(週)";
        MergeAndAlign(currentRow + 8, 1, currentRow + 8, 3, isCenter: false);
        FormatCells(sheet.Cells[$"A{currentRow + 8}:C{currentRow + 8}"], bgColor: footerBgColor, fontColor: Color.White, isBorder: true);
        sheet.Cells[currentRow + 9, 1].Value = "預估開始日 :";
        MergeAndAlign(currentRow + 9, 1, currentRow + 9, 2, isCenter: false);
        sheet.Cells[currentRow + 10, 1].Value = "預估結束日 :";
        MergeAndAlign(currentRow + 10, 1, currentRow + 10, 2, isCenter: false);
        sheet.Cells[currentRow + 11, 1].Value = "預估週期 :";
        MergeAndAlign(currentRow + 11, 1, currentRow + 11, 2, isCenter: false);
        // 淺藍色
        FormatCells(sheet.Cells[$"A{currentRow + 9}:C{currentRow + 11}"], isThin: true, isBorder: true);
        // 淺藍色
        footerBgColor = Color.FromArgb(97, 203, 243);   // 淺藍色
        FormatCells(sheet.Cells[$"A{currentRow + 9}:B{currentRow + 11}"], bgColor: footerBgColor);
        // 在第 1 行插入 10 行空白行
        sheet.InsertRow(1, 10);
        sheet.Cells[1, 1].Value = "報價單 (供內部使用)";
        sheet.Cells[1, 1].Style.Font.Size = 24;
        MergeAndAlign(1, 1, 1, header.Length, isCenter: true);
        for (int col = 0; col < columnNames.Length; col++)
        {
            sheet.Cells[3 + (col % 7), 1 + ((col / 7) * 3)].Value = $"  {columnNames[col]}：";
        }
        AlignColumnWidth();
    }
}

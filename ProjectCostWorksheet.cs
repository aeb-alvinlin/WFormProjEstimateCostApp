using OfficeOpenXml;
using static OfficeOpenXml.ExcelErrorValue;

public class ProjectCostWorksheet : WorksheetBase
{
    public override string[] header => new string[]
    {
        "專案階段", "工作編號", "工作項目", "工作說明", "工作天數(小計)", "專案經理", "", "部署者", "", "開發者", "", "內部成本小計", "備註"
    };

    public override int[] widthAlignment => new int[]
    {
        9, 14, 30, 45, 15, 8, 11, 8, 11, 8, 11, 15, 10
    };

    public override int lastRow { get; set; }

    public ProjectCostWorksheet(ExcelWorksheet sheet) : base(sheet)
    {
        projCostRefSheet = sheet.Name;
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

    // 寫入標題行
    public override void WriteAndFormatHeader(int startRow, int endRow)
    {
        string[] projcostheader = {
            "", "", "", "", "", "人天", "單價 (NT$)", "人天", "單價 (NT$)", "人天", "單價 (NT$)", "", "", ""
        };
        for (int col = 0; col < header.Length; col++)
        {
            sheet.Cells[lastRow, col + 1].Value = projcostheader[col];
        }
        // 每次寫入後遞增行
        lastRow++;
        MergeAndAlign(startRow, Array.IndexOf(header, "專案經理") + 1, startRow, Array.IndexOf(header, "專案經理") + 2);
        MergeAndAlign(startRow, Array.IndexOf(header, "部署者") + 1, startRow, Array.IndexOf(header, "部署者") + 2);
        MergeAndAlign(startRow, Array.IndexOf(header, "開發者") + 1, startRow, Array.IndexOf(header, "開發者") + 2);
        MergeAndAlign(startRow, Array.IndexOf(header, "專案階段") + 1, startRow + 1, Array.IndexOf(header, "專案階段") + 1);
        MergeAndAlign(startRow, Array.IndexOf(header, "工作編號") + 1, startRow + 1, Array.IndexOf(header, "工作編號") + 1);
        MergeAndAlign(startRow, Array.IndexOf(header, "工作項目") + 1, startRow + 1, Array.IndexOf(header, "工作項目") + 1);
        MergeAndAlign(startRow, Array.IndexOf(header, "工作說明") + 1, startRow + 1, Array.IndexOf(header, "工作說明") + 1);
        MergeAndAlign(startRow, Array.IndexOf(header, "工作天數(小計)") + 1, startRow + 1, Array.IndexOf(header, "工作天數(小計)") + 1);
        MergeAndAlign(startRow, Array.IndexOf(header, "內部成本小計") + 1, startRow + 1, Array.IndexOf(header, "內部成本小計") + 1);
        MergeAndAlign(startRow, Array.IndexOf(header, "備註") + 1, startRow + 1, Array.IndexOf(header, "備註") + 1);
        FormatCells(sheet.Cells[$"{startCol}{startRow}:{endCol}{endRow}"], isHeader: true, isHair: true);
    }

    public override void FormatPhase()
    {
        // 格式化內容
        MergeAndAlign(phaseStartRow + lastRow, 1, phaseEndRow + lastRow - 1, 1);
        referProjCostRow = phaseTitleRow + lastRow;
        FormatCells(sheet.Cells[$"{startCol}{phaseTitleRow + lastRow}:{endCol}{phaseEndRow + lastRow - 1}"], isHair: true, isBorder: true);
        // 使用 RGB 自定義顏色格式化階段標題
        Color titleBgColor = Color.FromArgb(198, 224, 180);  // 淺綠色
        MergeAndAlign(phaseTitleRow + lastRow, 1, phaseTitleRow + lastRow, header.Length);
        FormatCells(sheet.Cells[$"{startCol}{phaseTitleRow + lastRow}:{endCol}{phaseTitleRow + lastRow}"], bgColor: titleBgColor, isTitle: true);
    }

    public void WriteCostSumFormula(int column)
    {
        sheet.Cells[currentRow, column].Formula = $"SUM(F{currentRow}*G{currentRow},H{currentRow}*I{currentRow},J{currentRow}*K{currentRow})";
        sheet.Cells[currentRow, column].Style.Numberformat.Format = "$#,##0";
    }

    public override void WriteFooter()
    {
        sheet.Cells[currentRow, 1].Value = $"總計";
        sheet.Cells[currentRow, 5].Formula = $"SUM(E3:E{currentRow - 1})";
        CenterText(5);
        sheet.Cells[currentRow, 6].Formula = $"SUM(F3:F{currentRow - 1})";
        CenterText(6);
        sheet.Cells[currentRow, 8].Formula = $"SUM(H3:H{currentRow - 1})";
        CenterText(8);
        sheet.Cells[currentRow, 10].Formula = $"SUM(J3:J{currentRow - 1})";
        CenterText(10);
        sheet.Cells[currentRow, 12].Formula = $"SUM(L3:L{currentRow - 1})";
        sheet.Cells[currentRow, 12].Style.Numberformat.Format = "$#,##0";
        CenterText(12, isRight:true);
        Color footerBgColor = Color.FromArgb(248, 203, 173);    // 淺橘紅色
        FormatCells(sheet.Cells[$"{startCol}{currentRow}:{endCol}{currentRow}"], bgColor: footerBgColor, isBorder: true);
        referProjCostRow = phaseTitleRow + lastRow;
        footerBgColor = Color.FromArgb(198, 224, 180);  // 淺綠色
        sheet.Cells[currentRow + 1, 1].Value = "專案成員人天分配：";
        MergeAndAlign(currentRow + 1, 1, currentRow + 1, 3);
        FormatCells(sheet.Cells[$"A{currentRow + 1}:C{currentRow + 4}"], isBorder: true, fontColor:Color.White, bgColor: Color.FromArgb(34, 43, 53)); ;  // 深黑色
        sheet.Cells[currentRow + 1, 12].Value = "建議報價";
        MergeAndAlign(currentRow + 1, 12, currentRow + 1, 13, isCenter: true);
        // 取得今天的日期，並格式化為 "西元年/月/日" 格式
        string getDatetimeToday = DateTime.Now.ToString("yyyy/MM/dd");
        sheet.Cells[currentRow + 2, 1].Value = "A";
        sheet.Cells[currentRow + 2, 2].Value = "專案經理";
        sheet.Cells[currentRow + 2, 3].Formula = $"=F{currentRow}";
        sheet.Cells[currentRow + 2, 12].Formula = $"=L{currentRow}";
        sheet.Cells[currentRow + 2, 12].Style.Numberformat.Format = "$#,##0";
        MergeAndAlign(currentRow + 2, 12, currentRow + 2, 13, isCenter: true);
        FormatCells(sheet.Cells[$"L{currentRow + 1}:M{currentRow + 2}"], bgColor: footerBgColor, isBorder: true);
        sheet.Cells[currentRow + 3, 1].Value = "C";
        sheet.Cells[currentRow + 3, 2].Value = "部署者|架構師";
        sheet.Cells[currentRow + 3, 3].Formula = $"=H{currentRow}";
        sheet.Cells[currentRow + 3, 12].Value = $"製表日期：{getDatetimeToday}";
        FormatCells(sheet.Cells[$"L{currentRow + 3}:M{currentRow + 3}"], isBorder: true);
        sheet.Cells[currentRow + 4, 1].Value = "D";
        sheet.Cells[currentRow + 4, 2].Value = "開發者";
        sheet.Cells[currentRow + 4, 3].Formula = $"=J{currentRow}";
        FormatCells(sheet.Cells[$"A{currentRow + 2}:C{currentRow + 4}"], bgColor: footerBgColor, isBorder: true, isThin: true);
        FormatCells(sheet.Cells[$"B{currentRow + 2}:B{currentRow + 4}"], bgColor: footerBgColor, isBorder: true);
        sheet.Cells[currentRow + 6, 1].Value = "註1.本專案成本表所列的總工天為專員成員的工作總天數，非專案建置所需日曆天";
        AlignColumnWidth();
    }
}

using System.Text.RegularExpressions;
using System.Collections.Generic;
// nuget package "EPPlus" Version 7.0.10
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Collections;


public class TaskItemRow
{
    public string TaskName { get; set; }        // 工作項目
    public string TaskDescription { get; set; } // 工作說明
    public double TotalTaskDays { get; set; }   // 工作天數(小計)
    public double PrjManagerDays { get; set; }  // 專案經理
    public double DeployerDays { get; set; }    // 部署者
    public double DeveloperDays { get; set; }   // 開發者

    public TaskItemRow(string taskName, string taskDescription, double totalTaskDays, double prjmangerDays, double deployerDays, double developerDays)
    {
        TaskName = taskName;
        TaskDescription = taskDescription;
        TotalTaskDays = totalTaskDays;
        PrjManagerDays = prjmangerDays;
        DeployerDays = deployerDays;
        DeveloperDays = developerDays;
    }
}

public class TaskItemsContext
{
    private string taskItemsFilePath;
    private ExcelPackage taskPackage;
    public string errorMessage = null!; 
    public Dictionary< string, List<TaskItemRow> > phaseData;
    public Dictionary< string, int > phaseCount;

    public TaskItemsContext(string filePath)
    {
        taskItemsFilePath = filePath;
        taskPackage = new ExcelPackage(new FileInfo(taskItemsFilePath));
        phaseData = new Dictionary <string, List<TaskItemRow> >();
        phaseCount = new Dictionary <string, int >();
        ReadTaskItemsSourceFile();
    }

    // 定義每個階段的工作表名稱
    public string[] taskItemsHeader = {
        "工作項目",
        "工作說明",
        "工作天數(小計)",
        "專案經理",
        "部署者",
        "開發者",
    };

    // 定義每個階段的工作表名稱
    protected string[] phaseNames = {
        "第一階段-環境調查 (Envisioning)",
        "第二階段-設計規劃 (Planning)",
        "第三階段-發展階段 (Develop)",
        "第四階段-系統部署 (Deploying)",
        "第五階段-結案階段 (Ending)",
        "維護保固階段 (Maintance)",
    };

    public void ReadTaskItemsSourceFile()
    {
        using (ExcelPackage taskPackage = new ExcelPackage(new FileInfo(taskItemsFilePath!)))
        {
            // 讀取所有工作表名稱，存入字串陣列，並排序
            var worksheetNames = new List<string>();
            foreach (var sheet in taskPackage.Workbook.Worksheets)
            {
                worksheetNames.Add(sheet.Name); // 加入工作表名稱
            }

            // 建立一個 worksheetNames 的副本來進行驗證
            var verificationList = new List<string>(worksheetNames);

            // 將工作表名稱依照 phaseName 和編號排序 (無編號的在前, 有編號的按順序排列)
            var sortedWorksheetNames = worksheetNames
                .OrderBy(name =>
                {
                    // 匹配括號中的數字
                    var match = Regex.Match(name, @" \((\d+)\)$");
                    // 按照數字大小排序
                    return match.Success ? int.Parse(match.Groups[1].Value) : 0; 
                })
                // 無編號的排在前
                .ThenBy(name => name) 
                .ToList();

            // 讀取每個階段的工作表，並將其資料存入 Dictionary 中
            foreach (string phaseName in phaseNames)
            {
                // 使用正則表達式檢查是否存在與 phaseName 相似的工作表名稱
                string pattern = $"^{Regex.Escape(phaseName)}( \\(\\d+\\))?$";
                var matchingSheetNames = sortedWorksheetNames.Where(name => Regex.IsMatch(name, pattern)).ToList();

                if (matchingSheetNames.Count == 0)
                {
                    errorMessage = $"檔案中沒有名稱為 {phaseName} 的工作表！";
                    continue;
                }

                foreach (var matchingSheetName in matchingSheetNames)
                {
                    string[] phaseNameParts = matchingSheetName.Split('-');
                    // 擷取分割後的第一部分並使用 Trim() 移除首尾空白字元
                    string phasePart = phaseNameParts[0].Trim();
                    // 找到的表從 verificationList 移除
                    verificationList.Remove(matchingSheetName);

                    var phaseSheet = taskPackage.Workbook.Worksheets[matchingSheetName];
                    var (errorField, phaseRows) = ReadPhaseSheet(matchingSheetName, phaseSheet);

                    // 檢查是否有資料
                    if (errorField != null)
                    {
                        errorMessage = $"發生錯誤，{errorField}！";
                    }
                    else if (phaseRows.Count == 0)
                    {
                        errorMessage = $"工作表 {matchingSheetName} 中沒有有效資料！";
                    }
                    else
                    {
                        phaseData[matchingSheetName] = phaseRows;
                        // 檢查 phasePart 是否已經存在於 phaseCount 字典中
                        if (phaseCount.ContainsKey(phasePart))
                        {
                            // 如果存在，則累加行數
                            phaseCount[phasePart] += phaseRows.Count;
                        }
                        else
                        {
                            // 如果不存在，則初始化
                            phaseCount[phasePart] = phaseRows.Count;
                        }
                    }
                }
            }
            // 檢查是否有未被 remove 出的工作表名稱
            if (verificationList.Count > 0)
            {
                errorMessage = $"未能成功讀入所有工作表，以下工作表未被讀取: {string.Join(", ", verificationList)}";
            }
        }
    }

    // 用來讀取每個階段的工作表並返回該工作表中的每一行資料
    private (string? errorField, List<TaskItemRow> rows) ReadPhaseSheet(string phaseName, ExcelWorksheet sheet)
    {
        var rows = new List<TaskItemRow>();
        int totalRows = sheet.Dimension.End.Row;
        int totalCols = taskItemsHeader.Length;

        // 讀取第一列標題，並建立標題對應的欄號字典
        var headerMap = new Dictionary<string, int>();
        for (int col = 0; col < totalCols; col++)
        {
            string header = sheet.Cells[1, col+1].Text;
            if (!string.IsNullOrEmpty(header))
            {
                headerMap[header] = col; // 建立標題對應的欄號
            }
        }
        // 如果缺少必要欄位則回傳錯誤
        foreach (var header in taskItemsHeader)
        {
            if (!headerMap.ContainsKey(header))
            {
                return ($"在 [{phaseName}] 表中缺少必要欄位: {header}", new List<TaskItemRow>()); 
            }
        }

        if (totalRows >= 2 && sheet.Cells[totalRows, 1].Value != null)
        {
            // 從第二行開始讀取（假設第一行是標題）
            for (int row = 2; row <= totalRows; row++)  // 行從第 2 行開始
            {
                // 根據標題對應的欄位，讀取每一行的資料
                string taskName = sheet.Cells[row, headerMap["工作項目"] + 1].Text;
                string taskDescription = sheet.Cells[row, headerMap["工作說明"] + 1].Text;
                bool totalTask = double.TryParse(sheet.Cells[row, headerMap["工作天數(小計)"] + 1].Text, out double totalTaskDays);
                bool prjmanger = double.TryParse(sheet.Cells[row, headerMap["專案經理"] + 1].Text, out double prjmangerDays);
                bool deployer = double.TryParse(sheet.Cells[row, headerMap["部署者"] + 1].Text, out double deployerDays);
                bool developer = double.TryParse(sheet.Cells[row, headerMap["開發者"] + 1].Text, out double developerDays);


                // 回傳錯誤欄位和空的資料列
                if (string.IsNullOrEmpty(taskName))
                {
                    return ($"在 [{phaseName}] 表中的第 {row} 列的 [工作項目] 不是數值", new List<TaskItemRow>());  
                }
                // 回傳錯誤欄位和空的資料列
                if (string.IsNullOrEmpty(taskDescription))
                {
                    return ($"在 [{phaseName}] 表中的第 {row} 列的 [工作說明] 不是數值", new List<TaskItemRow>());  
                }
                // 回傳錯誤欄位和空的資料列
                if (!totalTask)
                {
                    return ($"在 [{phaseName}] 表中的第 {row} 列的 [工作天數(小計)] 不是數值", new List<TaskItemRow>());  
                }
                // 回傳錯誤欄位和空的資料列
                if (!prjmanger)
                {
                    return ($"在 [{phaseName}] 表中的第 {row} 列的 [專案經理] 不是數值", new List<TaskItemRow>());  
                }
                // 回傳錯誤欄位和空的資料列
                if (!deployer)
                {
                    return ($"在 [{phaseName}] 表中的第 {row} 列的 [部署者] 不是數值", new List<TaskItemRow>());  
                }
                // 回傳錯誤欄位和空的資料列
                if (!developer)
                {
                    return ($"在 [{phaseName}] 表中的第 {row} 列的 [開發者] 不是數值", new List<TaskItemRow>());  
                }

                double sumTaskDays = prjmangerDays + deployerDays + developerDays;

                // 檢查總天數是否一致
                if (sumTaskDays != totalTaskDays)
                {
                    return ($"在 [{phaseName}] 表中的第 {row} 列的 [總工作天數] 不一致", new List<TaskItemRow>());  
                }

                // 如果所有數據都正確，則加入到 rows 中
                rows.Add(new TaskItemRow(taskName, taskDescription, totalTaskDays, prjmangerDays, deployerDays, developerDays));
            }
        }
        // 如果沒有錯誤，回傳 null 和有效的資料列
        return (null, rows);  
    }
}


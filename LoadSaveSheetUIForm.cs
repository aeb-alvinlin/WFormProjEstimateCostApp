using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml; 
using OfficeOpenXml.Style;
using static System.Runtime.InteropServices.JavaScript.JSType;


namespace WFormProjEstimateApp1
{
    public partial class WFormProjEstimate : Form
    {
        string taskListSheetName = "工作項目清單";
        string projCostSheetName = "專案成本表";
        string quotationSheetName = "報價單(供內部使用)";
        string deliverablesSheetName = "專案文件交付清單";

        private string? taskItemsFilePath = null;
        private string? targetFilePath = null;

        public WFormProjEstimate(string? sourceExcelFilePath)
        {
            // Use EPPlus in a noncommercial context according to the Polyform Noncommercial license  
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            if (sourceExcelFilePath != null)
            {
                taskItemsFilePath = sourceExcelFilePath;
                GenerateTargetFilePath();
                GenerateQuotationReport();
                MessageBox.Show($"已讀取「{taskItemsFilePath}」並產生報價試算表「{targetFilePath}」！");
                return;
            }

            InitializeComponent();
        }

        // 開啟來源工作清單 Excel 檔案
        private void OpenTaskItemsSource_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel 檔案 (*.xlsx)|*.xlsx";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    taskItemsFilePath = openFileDialog.FileName;
                    StatusBarLabel.Text = $"已讀入檔案: {taskItemsFilePath}";
                }
            }
        }

        // 產生目標工作報表 Excel 檔案名稱與路徑
        private void GenerateTargetFilePath()
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            do
            {
                // 使用 DateTime 取得當前日期和時間
                string currentDate = DateTime.Now.ToString("yyyyMMddHHmmss");
                // 目標檔案名稱：專案成本表.xlsx - 產生新的空白的Excel檔案做為專案成本表
                string costFileName = @$"專案成本表_{currentDate}.xlsx";
                // 使用取得目標檔案名稱和路徑
                targetFilePath = Path.Combine(desktopPath, costFileName);
                // 檢查直到確認此檔案並沒有相同名稱的檔名存在
            } while (File.Exists(targetFilePath));
        }

        // 產生目標工作報表 Excel 檔案
        private void SaveQuotationReportTarget_Click(object sender, EventArgs e)
        {
            if (taskItemsFilePath == null)
            {
                MessageBox.Show("請先透過選單的「讀取來源檔案」選取工作清單試算表！");
                return;
            }

            GenerateTargetFilePath();
            GenerateQuotationReport();
            MessageBox.Show($"已讀取「{taskItemsFilePath}」並產生報價試算表「{targetFilePath}」！");
            Application.Exit();
        }

        // 儲存目標工作報表 Excel 檔案
        private void GenerateQuotationReport()
        {
            // 呼叫 TaskItemsContext 處理工作表
            TaskItemsContext taskItems = new(taskItemsFilePath!);
            if (!string.IsNullOrEmpty(taskItems.errorMessage))
            {
                MessageBox.Show(taskItems.errorMessage);
                StatusBarLabel.Text = taskItems.errorMessage;
                return;
            }

            using (ExcelPackage costPackage = new ExcelPackage())
            {
                // 啟始類別
                var quotationSheet = costPackage.Workbook.Worksheets.Add(quotationSheetName);
                var projCostSheet = costPackage.Workbook.Worksheets.Add(projCostSheetName);
                var taskListSheet = costPackage.Workbook.Worksheets.Add(taskListSheetName);
                var deliverablesSheet = costPackage.Workbook.Worksheets.Add(deliverablesSheetName);

                // 使用 WorksheetBase 的子類別來協助寫入各工作表的資料
                // 將各工作表的物件放入 using 區塊，確保資源釋放
                using (QuotationWorksheet quotation = new QuotationWorksheet(quotationSheet))
                using (ProjectCostWorksheet projectCost = new ProjectCostWorksheet(projCostSheet))
                using (TaskListWorksheet taskList = new TaskListWorksheet(taskListSheet))
                using (DeliverablesWorkSheet deliverables = new DeliverablesWorkSheet(deliverablesSheet))
                {
                    // 定義 "階段編號" 從 1 開始計算
                    int phaseNumber = 1;
                    // 定義 "序號" 從 1 開始計算
                    int sequenceNumber = 1;
                    // 階段的順序切片
                    string lastPhasePart = null!;

                    // 寫入工作表標題
                    taskList.WriteHeader();
                    projectCost.WriteHeader();
                    quotation.WriteHeader();
                    deliverables.WriteHeader();

                    // 讀取的工作項目將工作簿裡的資料用各工作表各自的方法寫入新工作簿
                    foreach (var phase in taskItems.phaseData)
                    {
                        // phaseName 為 Key 也就是階段名稱
                        string phaseName = phase.Key;
                        // 階段名稱用 "-" 來分割字串
                        string[] phaseNameParts = phaseName.Split('-');
                        // 擷取階段名稱分割後階段的第一部分順序切片出來，並使用 Trim() 移除首尾空白字元
                        string phasePart = phaseNameParts[0].Trim();
                        // taskLists 為 Value 也就是該階段以 TaskItemRow 型別儲存的資料清單
                        List<TaskItemRow> taskLists = phase.Value;
                        // 比對目前的 階段的順序切片 是否就是前次的 階段的順序切片 相同的名稱
                        if (phasePart != lastPhasePart)
                        // 如果目前的 階段的順序切片 是新的階段 將以階段的開始做為寫入該階段的資料的形式
                        {
                            // 先以 SetPhaseTitle 設定 PhaseTitle 階段標題目前的位置
                            WorksheetBase.SetPhaseTitle();
                            // 寫入 "階段名稱"
                            taskList.WriteText(phaseName, 1);
                            projectCost.WriteText(phaseName, 1);
                            quotation.WriteText(phaseName, 1);
                            // 全部一起移到下一列
                            WorksheetBase.MoveSharedRowToNext();
                        }
                        // 定義 "階段編號" 後的 "點"->"大綱編號" 從 1 開始計算
                        int outlineNumber = 1;
                        // 先以 SetPhaseStart 設定 PhaseStart 階段項目目前的位置
                        WorksheetBase.SetPhaseStart();
                        // 開始從 taskLists 清單內逐一取出工作項目資料寫入 "階段項目" 
                        foreach (var item in taskLists)
                        {
                            // 如果目前 大綱編號 為 1 表示這是階段開頭的第一個編號
                            if (item.TaskName != "無")
                            {
                                if (outlineNumber == 1)
                                {
                                    // 只有在階段開頭的第一個編號時才寫入階段編號
                                    taskList.WriteText(phaseNumber, 1);
                                    projectCost.WriteText(phaseNumber, 1);
                                }
                                // 寫入大綱編號
                                taskList.WriteText($"{phaseNumber}.{outlineNumber}", 2);
                                // 寫入 序號
                                quotation.WriteValue(sequenceNumber, 1);
                                projectCost.WriteValue(sequenceNumber, 2);
                                // 寫入 工作天數
                                taskList.WriteText(item.TotalTaskDays, 5, isRight: false); ;
                                projectCost.WriteText(item.TotalTaskDays, 5, isRight: false); ;
                                // 寫入 專案經理
                                projectCost.WriteValue(item.PrjManagerDays, 6, isRight: false); ;
                                projectCost.WriteNumeric(8000, 7);
                                // 寫入 部署者
                                projectCost.WriteValue(item.DeployerDays, 8, isRight: false); ;
                                projectCost.WriteNumeric(8000, 9);
                                // 寫入 負責單位預設值為 "AEB"
                                taskList.WriteText("AEB", 9);
                                // 寫入 開發者
                                projectCost.WriteValue(item.DeveloperDays, 10);
                                projectCost.WriteNumeric(8000, 11);
                                projectCost.WriteCostSumFormula(12);
                            }

                            // 寫入 工作項目
                            taskList.WriteText(item.TaskName, 3, isCenter: false);
                            projectCost.WriteText(item.TaskName, 3, isCenter: false);
                            quotation.WriteText(item.TaskName, 2, isCenter: false);
                            // 寫入 工作說明
                            taskList.WriteText(item.TaskDescription, 4, isCenter: false);
                            projectCost.WriteText(item.TaskDescription, 4, isCenter: false);
                            quotation.WriteText(item.TaskDescription, 3, isCenter: false);

                            // 移到下一列
                            WorksheetBase.MoveSharedRowToNext();
                            // 大綱編號加 1
                            outlineNumber++;
                            // 序號加 1
                            sequenceNumber++;
                        }
                        // 寫入 工作項目
                        WorksheetBase.SetPhaseEnd();
                        // 合併 工作項目
                        taskList.MergeText(3);
                        projectCost.MergeText(3, sheetCalculate: true);
                        quotation.MergeText(2, sheetCalculate: true);
                        // 使用自定義顏色格式化階段 
                        taskList.FormatPhase();
                        // 使用自定義顏色格式化階段 
                        projectCost.FormatPhase();
                        // 使用自定義顏色格式化階段 
                        quotation.FormatPhase();
                        if (phasePart != lastPhasePart)
                        {
                            quotation.PhaseSumPrice(taskItems.phaseCount[phasePart]);
                        }
                        // 階段編號加 1
                        phaseNumber++;
                        // 整個階段完成後，將目前的 phasePart, 也就是階段名稱的階段順序切片部分 指派給 lastPhasePart。用來在下階段判斷是否還是相同階段順序
                        lastPhasePart = phasePart;
                    }
                    // 最後定位在表格的最後一行為 SetPhaseTitle。這是為了讓表尾的文字區段有位置的參考依據
                    WorksheetBase.SetPhaseTitle();
                    // 寫入表尾
                    taskList.WriteFooter();
                    // 寫入表尾
                    projectCost.WriteFooter();
                    // 寫入表尾
                    quotation.WriteFooter();
                    // 寫入表尾
                    deliverables.WriteFooter();
                    // 保存到目標檔案
                    costPackage.SaveAs(new FileInfo(targetFilePath!));
                    StatusBarLabel.Text = $"報價單已儲存: {targetFilePath}";
                    // 結束時將物件設為 null 垃圾回收
                }
            }
        }

        private void WFormProjEstimate_Load(object sender, EventArgs e)
        {
            string[] deliveyItems = [
                "A1 產品簡報",
                "A2 專案建議書",
                "A3 專案成本表",
                "A4 工作項目(Action Item)",
                "A5 工作說明書(SOW)",
                "B1 系統環境調查表",
                "B2 啟動會議簡報",
                "C1 架構流程圖",
                "C2 系統規畫建議書",
                "C3 整理程序說明書",
                "C4 工作分解結構(WBS)",
                "D1 功能驗證報告書",
                "E1 問題處理清單",
                "E2 管理手冊",
                "E3 操作手冊",
                "E4 教育訓練手冊",
                "F1 專案結案報告書",
                "F2 結案會議簡報",
                "G1 工作紀錄/會議紀錄",
                "G2 進度報告",
                "G3 週報",
                "G4 其他(郵件、截圖)",
                "G5 合約",
            ];
            deliverySelectionComboBox.Items.AddRange(deliveyItems);
        }

        private void deliveryListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string selection = deliverySelectionComboBox.Text;
            // 檢查姓名是否為空字串
            if (selection == "")
            {
                MessageBox.Show("請選擇交付項目再按新增！");
                // 離開此事件處理函式
                return;
            }
            if (deliveryListBox.Items.Contains(selection))
            {
                MessageBox.Show("資料已存在!");
            }
            else
            {
                deliveryListBox.Items.Add(selection);
            }
        }
    }
}

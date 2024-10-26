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
        string taskListSheetName = "�u�@���زM��";
        string projCostSheetName = "�M�צ�����";
        string quotationSheetName = "������(�Ѥ����ϥ�)";
        string deliverablesSheetName = "�M�פ���I�M��";

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
                MessageBox.Show($"�wŪ���u{taskItemsFilePath}�v�ò��ͳ����պ��u{targetFilePath}�v�I");
                return;
            }

            InitializeComponent();
        }

        // �}�Ҩӷ��u�@�M�� Excel �ɮ�
        private void OpenTaskItemsSource_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel �ɮ� (*.xlsx)|*.xlsx";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    taskItemsFilePath = openFileDialog.FileName;
                    StatusBarLabel.Text = $"�wŪ�J�ɮ�: {taskItemsFilePath}";
                }
            }
        }

        // ���ͥؼФu�@���� Excel �ɮצW�ٻP���|
        private void GenerateTargetFilePath()
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            do
            {
                // �ϥ� DateTime ���o��e����M�ɶ�
                string currentDate = DateTime.Now.ToString("yyyyMMddHHmmss");
                // �ؼ��ɮצW�١G�M�צ�����.xlsx - ���ͷs���ťժ�Excel�ɮװ����M�צ�����
                string costFileName = @$"�M�צ�����_{currentDate}.xlsx";
                // �ϥΨ��o�ؼ��ɮצW�٩M���|
                targetFilePath = Path.Combine(desktopPath, costFileName);
                // �ˬd����T�{���ɮרèS���ۦP�W�٪��ɦW�s�b
            } while (File.Exists(targetFilePath));
        }

        // ���ͥؼФu�@���� Excel �ɮ�
        private void SaveQuotationReportTarget_Click(object sender, EventArgs e)
        {
            if (taskItemsFilePath == null)
            {
                MessageBox.Show("�Х��z�L��檺�uŪ���ӷ��ɮסv����u�@�M��պ��I");
                return;
            }

            GenerateTargetFilePath();
            GenerateQuotationReport();
            MessageBox.Show($"�wŪ���u{taskItemsFilePath}�v�ò��ͳ����պ��u{targetFilePath}�v�I");
            Application.Exit();
        }

        // �x�s�ؼФu�@���� Excel �ɮ�
        private void GenerateQuotationReport()
        {
            // �I�s TaskItemsContext �B�z�u�@��
            TaskItemsContext taskItems = new(taskItemsFilePath!);
            if (!string.IsNullOrEmpty(taskItems.errorMessage))
            {
                MessageBox.Show(taskItems.errorMessage);
                StatusBarLabel.Text = taskItems.errorMessage;
                return;
            }

            using (ExcelPackage costPackage = new ExcelPackage())
            {
                // �ҩl���O
                var quotationSheet = costPackage.Workbook.Worksheets.Add(quotationSheetName);
                var projCostSheet = costPackage.Workbook.Worksheets.Add(projCostSheetName);
                var taskListSheet = costPackage.Workbook.Worksheets.Add(taskListSheetName);
                var deliverablesSheet = costPackage.Workbook.Worksheets.Add(deliverablesSheetName);

                // �ϥ� WorksheetBase ���l���O�Ө�U�g�J�U�u�@�����
                // �N�U�u�@�������J using �϶��A�T�O�귽����
                using (QuotationWorksheet quotation = new QuotationWorksheet(quotationSheet))
                using (ProjectCostWorksheet projectCost = new ProjectCostWorksheet(projCostSheet))
                using (TaskListWorksheet taskList = new TaskListWorksheet(taskListSheet))
                using (DeliverablesWorkSheet deliverables = new DeliverablesWorkSheet(deliverablesSheet))
                {
                    // �w�q "���q�s��" �q 1 �}�l�p��
                    int phaseNumber = 1;
                    // �w�q "�Ǹ�" �q 1 �}�l�p��
                    int sequenceNumber = 1;
                    // ���q�����Ǥ���
                    string lastPhasePart = null!;

                    // �g�J�u�@����D
                    taskList.WriteHeader();
                    projectCost.WriteHeader();
                    quotation.WriteHeader();
                    deliverables.WriteHeader();

                    // Ū�����u�@���رN�u�@ï�̪���ƥΦU�u�@��U�۪���k�g�J�s�u�@ï
                    foreach (var phase in taskItems.phaseData)
                    {
                        // phaseName �� Key �]�N�O���q�W��
                        string phaseName = phase.Key;
                        // ���q�W�٥� "-" �Ӥ��Φr��
                        string[] phaseNameParts = phaseName.Split('-');
                        // �^�����q�W�٤��Ϋᶥ�q���Ĥ@�������Ǥ����X�ӡA�èϥ� Trim() ���������ťզr��
                        string phasePart = phaseNameParts[0].Trim();
                        // taskLists �� Value �]�N�O�Ӷ��q�H TaskItemRow ���O�x�s����ƲM��
                        List<TaskItemRow> taskLists = phase.Value;
                        // ���ثe�� ���q�����Ǥ��� �O�_�N�O�e���� ���q�����Ǥ��� �ۦP���W��
                        if (phasePart != lastPhasePart)
                        // �p�G�ثe�� ���q�����Ǥ��� �O�s�����q �N�H���q���}�l�����g�J�Ӷ��q����ƪ��Φ�
                        {
                            // ���H SetPhaseTitle �]�w PhaseTitle ���q���D�ثe����m
                            WorksheetBase.SetPhaseTitle();
                            // �g�J "���q�W��"
                            taskList.WriteText(phaseName, 1);
                            projectCost.WriteText(phaseName, 1);
                            quotation.WriteText(phaseName, 1);
                            // �����@�_����U�@�C
                            WorksheetBase.MoveSharedRowToNext();
                        }
                        // �w�q "���q�s��" �᪺ "�I"->"�j���s��" �q 1 �}�l�p��
                        int outlineNumber = 1;
                        // ���H SetPhaseStart �]�w PhaseStart ���q���إثe����m
                        WorksheetBase.SetPhaseStart();
                        // �}�l�q taskLists �M�椺�v�@���X�u�@���ظ�Ƽg�J "���q����" 
                        foreach (var item in taskLists)
                        {
                            // �p�G�ثe �j���s�� �� 1 ��ܳo�O���q�}�Y���Ĥ@�ӽs��
                            if (item.TaskName != "�L")
                            {
                                if (outlineNumber == 1)
                                {
                                    // �u���b���q�}�Y���Ĥ@�ӽs���ɤ~�g�J���q�s��
                                    taskList.WriteText(phaseNumber, 1);
                                    projectCost.WriteText(phaseNumber, 1);
                                }
                                // �g�J�j���s��
                                taskList.WriteText($"{phaseNumber}.{outlineNumber}", 2);
                                // �g�J �Ǹ�
                                quotation.WriteValue(sequenceNumber, 1);
                                projectCost.WriteValue(sequenceNumber, 2);
                                // �g�J �u�@�Ѽ�
                                taskList.WriteText(item.TotalTaskDays, 5, isRight: false); ;
                                projectCost.WriteText(item.TotalTaskDays, 5, isRight: false); ;
                                // �g�J �M�׸g�z
                                projectCost.WriteValue(item.PrjManagerDays, 6, isRight: false); ;
                                projectCost.WriteNumeric(8000, 7);
                                // �g�J ���p��
                                projectCost.WriteValue(item.DeployerDays, 8, isRight: false); ;
                                projectCost.WriteNumeric(8000, 9);
                                // �g�J �t�d���w�]�Ȭ� "AEB"
                                taskList.WriteText("AEB", 9);
                                // �g�J �}�o��
                                projectCost.WriteValue(item.DeveloperDays, 10);
                                projectCost.WriteNumeric(8000, 11);
                                projectCost.WriteCostSumFormula(12);
                            }

                            // �g�J �u�@����
                            taskList.WriteText(item.TaskName, 3, isCenter: false);
                            projectCost.WriteText(item.TaskName, 3, isCenter: false);
                            quotation.WriteText(item.TaskName, 2, isCenter: false);
                            // �g�J �u�@����
                            taskList.WriteText(item.TaskDescription, 4, isCenter: false);
                            projectCost.WriteText(item.TaskDescription, 4, isCenter: false);
                            quotation.WriteText(item.TaskDescription, 3, isCenter: false);

                            // ����U�@�C
                            WorksheetBase.MoveSharedRowToNext();
                            // �j���s���[ 1
                            outlineNumber++;
                            // �Ǹ��[ 1
                            sequenceNumber++;
                        }
                        // �g�J �u�@����
                        WorksheetBase.SetPhaseEnd();
                        // �X�� �u�@����
                        taskList.MergeText(3);
                        projectCost.MergeText(3, sheetCalculate: true);
                        quotation.MergeText(2, sheetCalculate: true);
                        // �ϥΦ۩w�q�C��榡�ƶ��q 
                        taskList.FormatPhase();
                        // �ϥΦ۩w�q�C��榡�ƶ��q 
                        projectCost.FormatPhase();
                        // �ϥΦ۩w�q�C��榡�ƶ��q 
                        quotation.FormatPhase();
                        if (phasePart != lastPhasePart)
                        {
                            quotation.PhaseSumPrice(taskItems.phaseCount[phasePart]);
                        }
                        // ���q�s���[ 1
                        phaseNumber++;
                        // ��Ӷ��q������A�N�ثe�� phasePart, �]�N�O���q�W�٪����q���Ǥ������� ������ lastPhasePart�C�ΨӦb�U���q�P�_�O�_�٬O�ۦP���q����
                        lastPhasePart = phasePart;
                    }
                    // �̫�w��b��檺�̫�@�欰 SetPhaseTitle�C�o�O���F���������r�Ϭq����m���ѦҨ̾�
                    WorksheetBase.SetPhaseTitle();
                    // �g�J���
                    taskList.WriteFooter();
                    // �g�J���
                    projectCost.WriteFooter();
                    // �g�J���
                    quotation.WriteFooter();
                    // �g�J���
                    deliverables.WriteFooter();
                    // �O�s��ؼ��ɮ�
                    costPackage.SaveAs(new FileInfo(targetFilePath!));
                    StatusBarLabel.Text = $"������w�x�s: {targetFilePath}";
                    // �����ɱN����]�� null �U���^��
                }
            }
        }

        private void WFormProjEstimate_Load(object sender, EventArgs e)
        {
            string[] deliveyItems = [
                "A1 ���~²��",
                "A2 �M�׫�ĳ��",
                "A3 �M�צ�����",
                "A4 �u�@����(Action Item)",
                "A5 �u�@������(SOW)",
                "B1 �t�����ҽլd��",
                "B2 �Ұʷ|ĳ²��",
                "C1 �[�c�y�{��",
                "C2 �t�γW�e��ĳ��",
                "C3 ��z�{�ǻ�����",
                "C4 �u�@���ѵ��c(WBS)",
                "D1 �\�����ҳ��i��",
                "E1 ���D�B�z�M��",
                "E2 �޲z��U",
                "E3 �ާ@��U",
                "E4 �Ш|�V�m��U",
                "F1 �M�׵��׳��i��",
                "F2 ���׷|ĳ²��",
                "G1 �u�@����/�|ĳ����",
                "G2 �i�׳��i",
                "G3 �g��",
                "G4 ��L(�l��B�I��)",
                "G5 �X��",
            ];
            deliverySelectionComboBox.Items.AddRange(deliveyItems);
        }

        private void deliveryListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string selection = deliverySelectionComboBox.Text;
            // �ˬd�m�W�O�_���Ŧr��
            if (selection == "")
            {
                MessageBox.Show("�п�ܥ�I���ئA���s�W�I");
                // ���}���ƥ�B�z�禡
                return;
            }
            if (deliveryListBox.Items.Contains(selection))
            {
                MessageBox.Show("��Ƥw�s�b!");
            }
            else
            {
                deliveryListBox.Items.Add(selection);
            }
        }
    }
}

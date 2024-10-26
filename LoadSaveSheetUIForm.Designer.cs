namespace WFormProjEstimateApp1
{
    partial class WFormProjEstimate
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WFormProjEstimate));
            menuStrip1 = new MenuStrip();
            SourceFilelStripMenu = new ToolStripMenuItem();
            OpenTaskItemsSource = new ToolStripMenuItem();
            TargetFileStripMenu = new ToolStripMenuItem();
            SaveQuotationReportTarget = new ToolStripMenuItem();
            statusLable = new StatusStrip();
            StatusBarLabel = new ToolStripStatusLabel();
            mainWindowPanel = new Panel();
            projDepartmentGroup = new GroupBox();
            comboBox4 = new ComboBox();
            comboBox3 = new ComboBox();
            techRepresentativeExtensionLabel = new Label();
            textBox7 = new TextBox();
            techRepresentativeEmailLabel = new Label();
            textBox8 = new TextBox();
            techDepartmentRepresentativelabel = new Label();
            techDepartmentLabel = new Label();
            projSalesGroup = new GroupBox();
            salesRepresentativeComboBox = new ComboBox();
            salesDepartmentComboBox = new ComboBox();
            salesRepresentativeExtensionLabel = new Label();
            textBox6 = new TextBox();
            salesRepresentativeEmailLabel = new Label();
            textBox5 = new TextBox();
            salesRepresentativeLabel = new Label();
            salesDepartmentLabel = new Label();
            projIdentityGroup = new GroupBox();
            projDeliveryGroup = new GroupBox();
            button3 = new Button();
            button2 = new Button();
            button1 = new Button();
            deliveryListBox = new ListBox();
            deliverySelectionComboBox = new ComboBox();
            customerName = new Label();
            customerNametextBox = new TextBox();
            projectNametextBox = new TextBox();
            projectName = new Label();
            menuStrip1.SuspendLayout();
            statusLable.SuspendLayout();
            mainWindowPanel.SuspendLayout();
            projDepartmentGroup.SuspendLayout();
            projSalesGroup.SuspendLayout();
            projIdentityGroup.SuspendLayout();
            projDeliveryGroup.SuspendLayout();
            SuspendLayout();
            // 
            // menuStrip1
            // 
            menuStrip1.Items.AddRange(new ToolStripItem[] { SourceFilelStripMenu, TargetFileStripMenu });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(693, 24);
            menuStrip1.TabIndex = 0;
            menuStrip1.Text = "menuStrip1";
            // 
            // SourceFilelStripMenu
            // 
            SourceFilelStripMenu.DropDownItems.AddRange(new ToolStripItem[] { OpenTaskItemsSource });
            SourceFilelStripMenu.Name = "SourceFilelStripMenu";
            SourceFilelStripMenu.Size = new Size(91, 20);
            SourceFilelStripMenu.Text = "讀取來源檔案";
            // 
            // OpenTaskItemsSource
            // 
            OpenTaskItemsSource.Name = "OpenTaskItemsSource";
            OpenTaskItemsSource.Size = new Size(205, 22);
            OpenTaskItemsSource.Text = "開啟工作清單試算表...(&F)";
            OpenTaskItemsSource.Click += OpenTaskItemsSource_Click;
            // 
            // TargetFileStripMenu
            // 
            TargetFileStripMenu.DropDownItems.AddRange(new ToolStripItem[] { SaveQuotationReportTarget });
            TargetFileStripMenu.Name = "TargetFileStripMenu";
            TargetFileStripMenu.Size = new Size(79, 20);
            TargetFileStripMenu.Text = "產生報價單";
            // 
            // SaveQuotationReportTarget
            // 
            SaveQuotationReportTarget.Name = "SaveQuotationReportTarget";
            SaveQuotationReportTarget.Size = new Size(180, 22);
            SaveQuotationReportTarget.Text = "產生報價單試算表";
            SaveQuotationReportTarget.Click += SaveQuotationReportTarget_Click;
            // 
            // statusLable
            // 
            statusLable.Items.AddRange(new ToolStripItem[] { StatusBarLabel });
            statusLable.Location = new Point(0, 418);
            statusLable.Name = "statusLable";
            statusLable.Size = new Size(693, 22);
            statusLable.TabIndex = 1;
            statusLable.Text = "狀態";
            // 
            // StatusBarLabel
            // 
            StatusBarLabel.Name = "StatusBarLabel";
            StatusBarLabel.Size = new Size(31, 17);
            StatusBarLabel.Text = "狀態";
            // 
            // mainWindowPanel
            // 
            mainWindowPanel.BorderStyle = BorderStyle.FixedSingle;
            mainWindowPanel.Controls.Add(projDepartmentGroup);
            mainWindowPanel.Controls.Add(projSalesGroup);
            mainWindowPanel.Controls.Add(projIdentityGroup);
            mainWindowPanel.ForeColor = SystemColors.ActiveCaptionText;
            mainWindowPanel.Location = new Point(12, 35);
            mainWindowPanel.Name = "mainWindowPanel";
            mainWindowPanel.Size = new Size(658, 372);
            mainWindowPanel.TabIndex = 2;
            // 
            // projDepartmentGroup
            // 
            projDepartmentGroup.Controls.Add(comboBox4);
            projDepartmentGroup.Controls.Add(comboBox3);
            projDepartmentGroup.Controls.Add(techRepresentativeExtensionLabel);
            projDepartmentGroup.Controls.Add(textBox7);
            projDepartmentGroup.Controls.Add(techRepresentativeEmailLabel);
            projDepartmentGroup.Controls.Add(textBox8);
            projDepartmentGroup.Controls.Add(techDepartmentRepresentativelabel);
            projDepartmentGroup.Controls.Add(techDepartmentLabel);
            projDepartmentGroup.Location = new Point(357, 193);
            projDepartmentGroup.Name = "projDepartmentGroup";
            projDepartmentGroup.Size = new Size(274, 164);
            projDepartmentGroup.TabIndex = 8;
            projDepartmentGroup.TabStop = false;
            projDepartmentGroup.Text = "專案團隊";
            // 
            // comboBox4
            // 
            comboBox4.FormattingEnabled = true;
            comboBox4.Location = new Point(84, 64);
            comboBox4.Name = "comboBox4";
            comboBox4.Size = new Size(168, 23);
            comboBox4.TabIndex = 10;
            // 
            // comboBox3
            // 
            comboBox3.FormattingEnabled = true;
            comboBox3.Location = new Point(84, 35);
            comboBox3.Name = "comboBox3";
            comboBox3.Size = new Size(168, 23);
            comboBox3.TabIndex = 10;
            // 
            // techRepresentativeExtensionLabel
            // 
            techRepresentativeExtensionLabel.AutoSize = true;
            techRepresentativeExtensionLabel.Location = new Point(21, 125);
            techRepresentativeExtensionLabel.Name = "techRepresentativeExtensionLabel";
            techRepresentativeExtensionLabel.Size = new Size(55, 15);
            techRepresentativeExtensionLabel.TabIndex = 17;
            techRepresentativeExtensionLabel.Text = "電話分機";
            // 
            // textBox7
            // 
            textBox7.Location = new Point(86, 122);
            textBox7.Name = "textBox7";
            textBox7.Size = new Size(168, 23);
            textBox7.TabIndex = 16;
            // 
            // techRepresentativeEmailLabel
            // 
            techRepresentativeEmailLabel.AutoSize = true;
            techRepresentativeEmailLabel.Location = new Point(20, 96);
            techRepresentativeEmailLabel.Name = "techRepresentativeEmailLabel";
            techRepresentativeEmailLabel.Size = new Size(55, 15);
            techRepresentativeEmailLabel.TabIndex = 15;
            techRepresentativeEmailLabel.Text = "電子信箱";
            // 
            // textBox8
            // 
            textBox8.Location = new Point(85, 93);
            textBox8.Name = "textBox8";
            textBox8.Size = new Size(168, 23);
            textBox8.TabIndex = 14;
            // 
            // techDepartmentRepresentativelabel
            // 
            techDepartmentRepresentativelabel.AutoSize = true;
            techDepartmentRepresentativelabel.Location = new Point(21, 67);
            techDepartmentRepresentativelabel.Name = "techDepartmentRepresentativelabel";
            techDepartmentRepresentativelabel.Size = new Size(55, 15);
            techDepartmentRepresentativelabel.TabIndex = 13;
            techDepartmentRepresentativelabel.Text = "部門代表";
            // 
            // techDepartmentLabel
            // 
            techDepartmentLabel.AutoSize = true;
            techDepartmentLabel.Location = new Point(20, 38);
            techDepartmentLabel.Name = "techDepartmentLabel";
            techDepartmentLabel.Size = new Size(55, 15);
            techDepartmentLabel.TabIndex = 11;
            techDepartmentLabel.Text = "技術部門";
            // 
            // projSalesGroup
            // 
            projSalesGroup.Controls.Add(salesRepresentativeComboBox);
            projSalesGroup.Controls.Add(salesDepartmentComboBox);
            projSalesGroup.Controls.Add(salesRepresentativeExtensionLabel);
            projSalesGroup.Controls.Add(textBox6);
            projSalesGroup.Controls.Add(salesRepresentativeEmailLabel);
            projSalesGroup.Controls.Add(textBox5);
            projSalesGroup.Controls.Add(salesRepresentativeLabel);
            projSalesGroup.Controls.Add(salesDepartmentLabel);
            projSalesGroup.Location = new Point(357, 19);
            projSalesGroup.Name = "projSalesGroup";
            projSalesGroup.Size = new Size(274, 159);
            projSalesGroup.TabIndex = 7;
            projSalesGroup.TabStop = false;
            projSalesGroup.Text = "業務團隊";
            // 
            // salesRepresentativeComboBox
            // 
            salesRepresentativeComboBox.FormattingEnabled = true;
            salesRepresentativeComboBox.Location = new Point(83, 58);
            salesRepresentativeComboBox.Name = "salesRepresentativeComboBox";
            salesRepresentativeComboBox.Size = new Size(168, 23);
            salesRepresentativeComboBox.TabIndex = 10;
            // 
            // salesDepartmentComboBox
            // 
            salesDepartmentComboBox.FormattingEnabled = true;
            salesDepartmentComboBox.Location = new Point(83, 29);
            salesDepartmentComboBox.Name = "salesDepartmentComboBox";
            salesDepartmentComboBox.Size = new Size(168, 23);
            salesDepartmentComboBox.TabIndex = 10;
            // 
            // salesRepresentativeExtensionLabel
            // 
            salesRepresentativeExtensionLabel.AutoSize = true;
            salesRepresentativeExtensionLabel.Location = new Point(19, 119);
            salesRepresentativeExtensionLabel.Name = "salesRepresentativeExtensionLabel";
            salesRepresentativeExtensionLabel.Size = new Size(55, 15);
            salesRepresentativeExtensionLabel.TabIndex = 9;
            salesRepresentativeExtensionLabel.Text = "電話分機";
            // 
            // textBox6
            // 
            textBox6.Location = new Point(84, 116);
            textBox6.Name = "textBox6";
            textBox6.Size = new Size(168, 23);
            textBox6.TabIndex = 8;
            // 
            // salesRepresentativeEmailLabel
            // 
            salesRepresentativeEmailLabel.AutoSize = true;
            salesRepresentativeEmailLabel.Location = new Point(18, 90);
            salesRepresentativeEmailLabel.Name = "salesRepresentativeEmailLabel";
            salesRepresentativeEmailLabel.Size = new Size(55, 15);
            salesRepresentativeEmailLabel.TabIndex = 7;
            salesRepresentativeEmailLabel.Text = "電子信箱";
            // 
            // textBox5
            // 
            textBox5.Location = new Point(83, 87);
            textBox5.Name = "textBox5";
            textBox5.Size = new Size(168, 23);
            textBox5.TabIndex = 6;
            // 
            // salesRepresentativeLabel
            // 
            salesRepresentativeLabel.AutoSize = true;
            salesRepresentativeLabel.Location = new Point(19, 61);
            salesRepresentativeLabel.Name = "salesRepresentativeLabel";
            salesRepresentativeLabel.Size = new Size(55, 15);
            salesRepresentativeLabel.TabIndex = 5;
            salesRepresentativeLabel.Text = "業務代表";
            // 
            // salesDepartmentLabel
            // 
            salesDepartmentLabel.AutoSize = true;
            salesDepartmentLabel.Location = new Point(18, 32);
            salesDepartmentLabel.Name = "salesDepartmentLabel";
            salesDepartmentLabel.Size = new Size(55, 15);
            salesDepartmentLabel.TabIndex = 3;
            salesDepartmentLabel.Text = "業務部門";
            // 
            // projIdentityGroup
            // 
            projIdentityGroup.Controls.Add(projDeliveryGroup);
            projIdentityGroup.Controls.Add(customerName);
            projIdentityGroup.Controls.Add(customerNametextBox);
            projIdentityGroup.Controls.Add(projectNametextBox);
            projIdentityGroup.Controls.Add(projectName);
            projIdentityGroup.Location = new Point(16, 19);
            projIdentityGroup.Name = "projIdentityGroup";
            projIdentityGroup.Size = new Size(322, 338);
            projIdentityGroup.TabIndex = 6;
            projIdentityGroup.TabStop = false;
            projIdentityGroup.Text = "專案與交付項目";
            // 
            // projDeliveryGroup
            // 
            projDeliveryGroup.Controls.Add(button3);
            projDeliveryGroup.Controls.Add(button2);
            projDeliveryGroup.Controls.Add(button1);
            projDeliveryGroup.Controls.Add(deliveryListBox);
            projDeliveryGroup.Controls.Add(deliverySelectionComboBox);
            projDeliveryGroup.Location = new Point(19, 94);
            projDeliveryGroup.Name = "projDeliveryGroup";
            projDeliveryGroup.Size = new Size(284, 238);
            projDeliveryGroup.TabIndex = 6;
            projDeliveryGroup.TabStop = false;
            projDeliveryGroup.Text = "專案交付項目";
            // 
            // button3
            // 
            button3.Location = new Point(186, 209);
            button3.Name = "button3";
            button3.Size = new Size(75, 23);
            button3.TabIndex = 2;
            button3.Text = "刪除項目";
            button3.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            button2.Location = new Point(101, 209);
            button2.Name = "button2";
            button2.Size = new Size(75, 23);
            button2.TabIndex = 2;
            button2.Text = "修改項目";
            button2.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            button1.Location = new Point(16, 209);
            button1.Name = "button1";
            button1.Size = new Size(75, 23);
            button1.TabIndex = 2;
            button1.Text = "新增項目";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // deliveryListBox
            // 
            deliveryListBox.FormattingEnabled = true;
            deliveryListBox.ItemHeight = 15;
            deliveryListBox.Location = new Point(16, 30);
            deliveryListBox.Name = "deliveryListBox";
            deliveryListBox.Size = new Size(245, 124);
            deliveryListBox.TabIndex = 1;
            deliveryListBox.SelectedIndexChanged += deliveryListBox_SelectedIndexChanged;
            // 
            // deliverySelectionComboBox
            // 
            deliverySelectionComboBox.FormattingEnabled = true;
            deliverySelectionComboBox.Location = new Point(16, 173);
            deliverySelectionComboBox.Name = "deliverySelectionComboBox";
            deliverySelectionComboBox.Size = new Size(245, 23);
            deliverySelectionComboBox.TabIndex = 0;
            // 
            // customerName
            // 
            customerName.AutoSize = true;
            customerName.Location = new Point(19, 32);
            customerName.Name = "customerName";
            customerName.Size = new Size(55, 15);
            customerName.TabIndex = 3;
            customerName.Text = "客戶名稱";
            // 
            // customerNametextBox
            // 
            customerNametextBox.Location = new Point(84, 29);
            customerNametextBox.Name = "customerNametextBox";
            customerNametextBox.Size = new Size(219, 23);
            customerNametextBox.TabIndex = 2;
            // 
            // projectNametextBox
            // 
            projectNametextBox.Location = new Point(85, 58);
            projectNametextBox.Name = "projectNametextBox";
            projectNametextBox.Size = new Size(218, 23);
            projectNametextBox.TabIndex = 4;
            // 
            // projectName
            // 
            projectName.AutoSize = true;
            projectName.Location = new Point(20, 61);
            projectName.Name = "projectName";
            projectName.Size = new Size(55, 15);
            projectName.TabIndex = 5;
            projectName.Text = "專案名稱";
            // 
            // WFormProjEstimate
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(693, 440);
            Controls.Add(mainWindowPanel);
            Controls.Add(statusLable);
            Controls.Add(menuStrip1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            MainMenuStrip = menuStrip1;
            Name = "WFormProjEstimate";
            Text = "工作項目成本估算報價單";
            Load += WFormProjEstimate_Load;
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            statusLable.ResumeLayout(false);
            statusLable.PerformLayout();
            mainWindowPanel.ResumeLayout(false);
            projDepartmentGroup.ResumeLayout(false);
            projDepartmentGroup.PerformLayout();
            projSalesGroup.ResumeLayout(false);
            projSalesGroup.PerformLayout();
            projIdentityGroup.ResumeLayout(false);
            projIdentityGroup.PerformLayout();
            projDeliveryGroup.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private MenuStrip menuStrip1;
        private ToolStripMenuItem SourceFilelStripMenu;
        private ToolStripMenuItem OpenTaskItemsSource;
        private StatusStrip statusLable;
        private ToolStripStatusLabel StatusBarLabel;
        private ToolStripMenuItem TargetFileStripMenu;
        private ToolStripMenuItem SaveQuotationReportTarget;
        private Panel mainWindowPanel;
        private GroupBox projDepartmentGroup;
        private GroupBox projSalesGroup;
        private Label salesDepartmentLabel;
        private GroupBox projIdentityGroup;
        private Label salesRepresentativeLabel;
        private Label salesRepresentativeExtensionLabel;
        protected internal TextBox textBox6;
        private Label salesRepresentativeEmailLabel;
        protected internal TextBox textBox5;
        private Label techRepresentativeExtensionLabel;
        protected internal TextBox textBox7;
        private Label techRepresentativeEmailLabel;
        protected internal TextBox textBox8;
        private Label techDepartmentRepresentativelabel;
        private Label techDepartmentLabel;
        private Label customerName;
        private TextBox customerNametextBox;
        protected internal TextBox projectNametextBox;
        private Label projectName;
        private GroupBox projDeliveryGroup;
        private ComboBox deliverySelectionComboBox;
        private Button button3;
        private Button button2;
        private Button button1;
        private ListBox deliveryListBox;
        private ComboBox comboBox4;
        private ComboBox comboBox3;
        private ComboBox salesRepresentativeComboBox;
        private ComboBox salesDepartmentComboBox;
    }
}

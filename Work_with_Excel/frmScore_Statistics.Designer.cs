namespace Work_with_Excel
{
    partial class frmScore_Statistics
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            dgvData = new DataGridView();
            groupBox1 = new GroupBox();
            cbGrade = new ComboBox();
            label2 = new Label();
            cbClass = new ComboBox();
            label4 = new Label();
            txtName = new TextBox();
            label3 = new Label();
            label1 = new Label();
            btnExport = new Button();
            btnImport = new Button();
            btnCAS = new Button();
            btnCreateInvitation = new Button();
            label5 = new Label();
            cbTopStudent = new ComboBox();
            ((System.ComponentModel.ISupportInitialize)dgvData).BeginInit();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // dgvData
            // 
            dgvData.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvData.Location = new Point(12, 62);
            dgvData.Name = "dgvData";
            dgvData.RowHeadersWidth = 62;
            dgvData.Size = new Size(1103, 370);
            dgvData.TabIndex = 9;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(cbGrade);
            groupBox1.Controls.Add(label2);
            groupBox1.Controls.Add(cbClass);
            groupBox1.Controls.Add(label4);
            groupBox1.Controls.Add(txtName);
            groupBox1.Controls.Add(label3);
            groupBox1.Location = new Point(12, 447);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(407, 168);
            groupBox1.TabIndex = 8;
            groupBox1.TabStop = false;
            groupBox1.Text = "Filter";
            // 
            // cbGrade
            // 
            cbGrade.FormattingEnabled = true;
            cbGrade.Location = new Point(83, 77);
            cbGrade.Name = "cbGrade";
            cbGrade.Size = new Size(304, 33);
            cbGrade.TabIndex = 8;
            cbGrade.SelectedValueChanged += cbGrade_SelectedValueChanged;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(6, 80);
            label2.Name = "label2";
            label2.Size = new Size(59, 25);
            label2.TabIndex = 7;
            label2.Text = "Grade";
            // 
            // cbClass
            // 
            cbClass.FormattingEnabled = true;
            cbClass.Location = new Point(83, 124);
            cbClass.Name = "cbClass";
            cbClass.Size = new Size(304, 33);
            cbClass.TabIndex = 6;
            cbClass.SelectedValueChanged += cbClass_SelectedValueChanged;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(6, 127);
            label4.Name = "label4";
            label4.Size = new Size(52, 25);
            label4.TabIndex = 5;
            label4.Text = "Class";
            // 
            // txtName
            // 
            txtName.Location = new Point(83, 30);
            txtName.Name = "txtName";
            txtName.Size = new Size(304, 31);
            txtName.TabIndex = 2;
            txtName.TextChanged += txtName_TextChanged;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(6, 33);
            label3.Name = "label3";
            label3.Size = new Size(59, 25);
            label3.TabIndex = 1;
            label3.Text = "Name";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 15F);
            label1.Location = new Point(457, 9);
            label1.Name = "label1";
            label1.Size = new Size(201, 41);
            label1.TabIndex = 7;
            label1.Text = "Score Statistic";
            // 
            // btnExport
            // 
            btnExport.Location = new Point(811, 522);
            btnExport.Name = "btnExport";
            btnExport.Size = new Size(112, 34);
            btnExport.TabIndex = 6;
            btnExport.Text = "Export";
            btnExport.UseVisualStyleBackColor = true;
            btnExport.Click += btnExport_Click;
            // 
            // btnImport
            // 
            btnImport.Location = new Point(487, 475);
            btnImport.Name = "btnImport";
            btnImport.Size = new Size(112, 34);
            btnImport.TabIndex = 5;
            btnImport.Text = "Import File";
            btnImport.UseVisualStyleBackColor = true;
            btnImport.Click += btnImport_Click;
            // 
            // btnCAS
            // 
            btnCAS.Location = new Point(487, 522);
            btnCAS.Name = "btnCAS";
            btnCAS.Size = new Size(210, 34);
            btnCAS.TabIndex = 10;
            btnCAS.Text = "Calculate Average Score";
            btnCAS.UseVisualStyleBackColor = true;
            btnCAS.Click += btnCAS_Click;
            // 
            // btnCreateInvitation
            // 
            btnCreateInvitation.Location = new Point(487, 569);
            btnCreateInvitation.Name = "btnCreateInvitation";
            btnCreateInvitation.Size = new Size(193, 34);
            btnCreateInvitation.TabIndex = 11;
            btnCreateInvitation.Text = "Create invitation letter";
            btnCreateInvitation.UseVisualStyleBackColor = true;
            btnCreateInvitation.Click += btnCreateInvitation_Click;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(733, 480);
            label5.Name = "label5";
            label5.Size = new Size(72, 25);
            label5.TabIndex = 9;
            label5.Text = "Option:";
            // 
            // cbTopStudent
            // 
            cbTopStudent.FormattingEnabled = true;
            cbTopStudent.Location = new Point(811, 475);
            cbTopStudent.Name = "cbTopStudent";
            cbTopStudent.Size = new Size(304, 33);
            cbTopStudent.TabIndex = 9;
            cbTopStudent.SelectedIndexChanged += cbTopStudent_SelectedIndexChanged;
            // 
            // frmScore_Statistics
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1127, 633);
            Controls.Add(cbTopStudent);
            Controls.Add(label5);
            Controls.Add(btnCreateInvitation);
            Controls.Add(btnCAS);
            Controls.Add(dgvData);
            Controls.Add(groupBox1);
            Controls.Add(label1);
            Controls.Add(btnExport);
            Controls.Add(btnImport);
            Name = "frmScore_Statistics";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Score Statistic";
            FormClosing += Score_Statistics_FormClosing;
            ((System.ComponentModel.ISupportInitialize)dgvData).EndInit();
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private DataGridView dgvData;
        private GroupBox groupBox1;
        private Label label1;
        private Button btnExport;
        private Button btnImport;
        private Button btnCAS;
        private Button btnCreateInvitation;
        private Label label4;
        private TextBox txtName;
        private Label label3;
        private ComboBox cbClass;
        private ComboBox cbGrade;
        private Label label2;
        private Label label5;
        private ComboBox cbTopStudent;
    }
}
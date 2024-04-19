namespace Work_with_Excel
{
    partial class frmCreate_Meeting_Invitation
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
            label1 = new Label();
            label2 = new Label();
            txtAddress = new TextBox();
            label3 = new Label();
            dtpTimeMeeting = new DateTimePicker();
            btnPrintInvitation = new Button();
            SuspendLayout();
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 15F);
            label1.Location = new Point(149, 9);
            label1.Name = "label1";
            label1.Size = new Size(367, 41);
            label1.TabIndex = 0;
            label1.Text = "Create Meeting Invitations";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Segoe UI", 12F);
            label2.Location = new Point(26, 74);
            label2.Name = "label2";
            label2.Size = new Size(98, 32);
            label2.TabIndex = 1;
            label2.Text = "Address";
            // 
            // txtAddress
            // 
            txtAddress.Font = new Font("Segoe UI", 12F);
            txtAddress.Location = new Point(130, 71);
            txtAddress.Multiline = true;
            txtAddress.Name = "txtAddress";
            txtAddress.Size = new Size(490, 46);
            txtAddress.TabIndex = 2;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Font = new Font("Segoe UI", 12F);
            label3.Location = new Point(26, 156);
            label3.Name = "label3";
            label3.Size = new Size(67, 32);
            label3.TabIndex = 3;
            label3.Text = "Time";
            // 
            // dtpTimeMeeting
            // 
            dtpTimeMeeting.CustomFormat = "HH:mm dd/MM/yyyy";
            dtpTimeMeeting.Font = new Font("Segoe UI", 12F);
            dtpTimeMeeting.Format = DateTimePickerFormat.Custom;
            dtpTimeMeeting.Location = new Point(130, 151);
            dtpTimeMeeting.Name = "dtpTimeMeeting";
            dtpTimeMeeting.Size = new Size(490, 39);
            dtpTimeMeeting.TabIndex = 4;
            // 
            // btnPrintInvitation
            // 
            btnPrintInvitation.BackColor = SystemColors.ControlLight;
            btnPrintInvitation.Font = new Font("Segoe UI", 12F);
            btnPrintInvitation.ForeColor = SystemColors.ActiveCaptionText;
            btnPrintInvitation.Location = new Point(190, 238);
            btnPrintInvitation.Name = "btnPrintInvitation";
            btnPrintInvitation.Size = new Size(300, 49);
            btnPrintInvitation.TabIndex = 5;
            btnPrintInvitation.Text = "Print Invitation Letter";
            btnPrintInvitation.UseVisualStyleBackColor = false;
            btnPrintInvitation.Click += btnPrintInvitation_Click;
            // 
            // frmCreate_Meeting_Invitation
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(674, 299);
            Controls.Add(btnPrintInvitation);
            Controls.Add(dtpTimeMeeting);
            Controls.Add(label3);
            Controls.Add(txtAddress);
            Controls.Add(label2);
            Controls.Add(label1);
            Name = "frmCreate_Meeting_Invitation";
            Text = "Create_Meeting_Invitation";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label label1;
        private Label label2;
        private TextBox txtAddress;
        private Label label3;
        private DateTimePicker dtpTimeMeeting;
        private Button btnPrintInvitation;
    }
}
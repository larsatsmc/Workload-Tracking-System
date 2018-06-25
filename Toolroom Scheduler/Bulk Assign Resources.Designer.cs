namespace Toolroom_Scheduler
{
    partial class Bulk_Assign_Resources_Form
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
			this.OKButton = new System.Windows.Forms.Button();
			this.myCancelButton = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.DesignerComboBox = new System.Windows.Forms.ComboBox();
			this.label4 = new System.Windows.Forms.Label();
			this.RoughProgrammerComboBox = new System.Windows.Forms.ComboBox();
			this.FinishProgrammerComboBox = new System.Windows.Forms.ComboBox();
			this.ElectrodeProgrammerComboBox = new System.Windows.Forms.ComboBox();
			this.SuspendLayout();
			// 
			// OKButton
			// 
			this.OKButton.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.OKButton.Location = new System.Drawing.Point(68, 192);
			this.OKButton.Name = "OKButton";
			this.OKButton.Size = new System.Drawing.Size(99, 31);
			this.OKButton.TabIndex = 84;
			this.OKButton.Text = "OK";
			this.OKButton.UseVisualStyleBackColor = true;
			this.OKButton.Click += new System.EventHandler(this.OKButton_Click);
			// 
			// myCancelButton
			// 
			this.myCancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.myCancelButton.Location = new System.Drawing.Point(184, 192);
			this.myCancelButton.Name = "myCancelButton";
			this.myCancelButton.Size = new System.Drawing.Size(99, 31);
			this.myCancelButton.TabIndex = 85;
			this.myCancelButton.Text = "Cancel";
			this.myCancelButton.UseVisualStyleBackColor = true;
			this.myCancelButton.Click += new System.EventHandler(this.CancelButton_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(46, 69);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(129, 16);
			this.label1.TabIndex = 86;
			this.label1.Text = "Rough Programmer:";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(51, 99);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(124, 16);
			this.label2.TabIndex = 87;
			this.label2.Text = "Finish Programmer:";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(28, 130);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(147, 16);
			this.label3.TabIndex = 88;
			this.label3.Text = "Electrode Programmer:";
			// 
			// DesignerComboBox
			// 
			this.DesignerComboBox.FormattingEnabled = true;
			this.DesignerComboBox.Location = new System.Drawing.Point(181, 36);
			this.DesignerComboBox.Name = "DesignerComboBox";
			this.DesignerComboBox.Size = new System.Drawing.Size(134, 24);
			this.DesignerComboBox.TabIndex = 89;
			this.DesignerComboBox.DropDown += new System.EventHandler(this.DesignerComboBox_DropDown);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(109, 39);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(66, 16);
			this.label4.TabIndex = 90;
			this.label4.Text = "Designer:";
			// 
			// RoughProgrammerComboBox
			// 
			this.RoughProgrammerComboBox.FormattingEnabled = true;
			this.RoughProgrammerComboBox.Location = new System.Drawing.Point(181, 66);
			this.RoughProgrammerComboBox.Name = "RoughProgrammerComboBox";
			this.RoughProgrammerComboBox.Size = new System.Drawing.Size(134, 24);
			this.RoughProgrammerComboBox.TabIndex = 91;
			this.RoughProgrammerComboBox.DropDown += new System.EventHandler(this.RoughProgrammerComboBox_DropDown);
			// 
			// FinishProgrammerComboBox
			// 
			this.FinishProgrammerComboBox.FormattingEnabled = true;
			this.FinishProgrammerComboBox.Location = new System.Drawing.Point(181, 96);
			this.FinishProgrammerComboBox.Name = "FinishProgrammerComboBox";
			this.FinishProgrammerComboBox.Size = new System.Drawing.Size(134, 24);
			this.FinishProgrammerComboBox.TabIndex = 92;
			this.FinishProgrammerComboBox.DropDown += new System.EventHandler(this.FinishProgrammerComboBox_DropDown);
			// 
			// ElectrodeProgrammerComboBox
			// 
			this.ElectrodeProgrammerComboBox.FormattingEnabled = true;
			this.ElectrodeProgrammerComboBox.Location = new System.Drawing.Point(181, 127);
			this.ElectrodeProgrammerComboBox.Name = "ElectrodeProgrammerComboBox";
			this.ElectrodeProgrammerComboBox.Size = new System.Drawing.Size(134, 24);
			this.ElectrodeProgrammerComboBox.TabIndex = 93;
			this.ElectrodeProgrammerComboBox.DropDown += new System.EventHandler(this.ElectrodeProgrammerComboBox_DropDown);
			// 
			// Bulk_Assign_Resources_Form
			// 
			this.AcceptButton = this.OKButton;
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(344, 250);
			this.Controls.Add(this.ElectrodeProgrammerComboBox);
			this.Controls.Add(this.FinishProgrammerComboBox);
			this.Controls.Add(this.RoughProgrammerComboBox);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.DesignerComboBox);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.myCancelButton);
			this.Controls.Add(this.OKButton);
			this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.Margin = new System.Windows.Forms.Padding(4);
			this.Name = "Bulk_Assign_Resources_Form";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Bulk Assign Resources";
			this.Load += new System.EventHandler(this.Bulk_Assign_Resources_Form_Load);
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button OKButton;
		private System.Windows.Forms.Button myCancelButton;
		private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox DesignerComboBox;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox RoughProgrammerComboBox;
        private System.Windows.Forms.ComboBox FinishProgrammerComboBox;
        private System.Windows.Forms.ComboBox ElectrodeProgrammerComboBox;
    }
}
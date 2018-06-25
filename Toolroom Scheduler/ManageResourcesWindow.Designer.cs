namespace Toolroom_Scheduler
{
	partial class ManageResourcesForm
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
            this.RoleComboBox = new System.Windows.Forms.ComboBox();
            this.roleListBox = new System.Windows.Forms.ListBox();
            this.AddRoleButton = new System.Windows.Forms.Button();
            this.RemoveRoleButton = new System.Windows.Forms.Button();
            this.automation_Task_ManagerDataSet2 = new Toolroom_Scheduler.Workload_Tracking_SystemDataSet();
            this.RemoveResourceButton = new System.Windows.Forms.Button();
            this.AddResourceButton = new System.Windows.Forms.Button();
            this.resourceListBox = new System.Windows.Forms.ListBox();
            this.addResourceTextBox = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.automation_Task_ManagerDataSet2)).BeginInit();
            this.SuspendLayout();
            // 
            // RoleComboBox
            // 
            this.RoleComboBox.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.RoleComboBox.FormattingEnabled = true;
            this.RoleComboBox.Items.AddRange(new object[] {
            "Tool Makers",
            "Designers",
            "Rough Programmers",
            "Finish Programmers",
            "Electrode Programmers",
            "Rough Mills",
            "Finish Mills",
            "Graphite Mills",
            "Rough CNC Operators",
            "Electrode CNC Operators",
            "Finish CNC Operators"});
            this.RoleComboBox.Location = new System.Drawing.Point(331, 25);
            this.RoleComboBox.Margin = new System.Windows.Forms.Padding(4);
            this.RoleComboBox.Name = "RoleComboBox";
            this.RoleComboBox.Size = new System.Drawing.Size(181, 24);
            this.RoleComboBox.TabIndex = 0;
            this.RoleComboBox.SelectedIndexChanged += new System.EventHandler(this.RoleComboBox_SelectedIndexChanged);
            // 
            // roleListBox
            // 
            this.roleListBox.FormattingEnabled = true;
            this.roleListBox.ItemHeight = 16;
            this.roleListBox.Location = new System.Drawing.Point(331, 55);
            this.roleListBox.Margin = new System.Windows.Forms.Padding(4);
            this.roleListBox.Name = "roleListBox";
            this.roleListBox.Size = new System.Drawing.Size(181, 420);
            this.roleListBox.TabIndex = 2;
            // 
            // AddRoleButton
            // 
            this.AddRoleButton.Location = new System.Drawing.Point(524, 55);
            this.AddRoleButton.Name = "AddRoleButton";
            this.AddRoleButton.Size = new System.Drawing.Size(91, 32);
            this.AddRoleButton.TabIndex = 5;
            this.AddRoleButton.Text = "Add";
            this.AddRoleButton.UseVisualStyleBackColor = true;
            this.AddRoleButton.Click += new System.EventHandler(this.AddRoleButton_Click);
            // 
            // RemoveRoleButton
            // 
            this.RemoveRoleButton.Location = new System.Drawing.Point(524, 93);
            this.RemoveRoleButton.Name = "RemoveRoleButton";
            this.RemoveRoleButton.Size = new System.Drawing.Size(91, 32);
            this.RemoveRoleButton.TabIndex = 6;
            this.RemoveRoleButton.Text = "Remove";
            this.RemoveRoleButton.UseVisualStyleBackColor = true;
            this.RemoveRoleButton.Click += new System.EventHandler(this.RemoveRoleButton_Click);
            // 
            // automation_Task_ManagerDataSet2
            // 
            this.automation_Task_ManagerDataSet2.DataSetName = "Automation_Task_ManagerDataSet2";
            this.automation_Task_ManagerDataSet2.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // RemoveResourceButton
            // 
            this.RemoveResourceButton.Location = new System.Drawing.Point(222, 93);
            this.RemoveResourceButton.Name = "RemoveResourceButton";
            this.RemoveResourceButton.Size = new System.Drawing.Size(91, 32);
            this.RemoveResourceButton.TabIndex = 13;
            this.RemoveResourceButton.Text = "Remove";
            this.RemoveResourceButton.UseVisualStyleBackColor = true;
            this.RemoveResourceButton.Click += new System.EventHandler(this.RemoveResourceButton_Click);
            // 
            // AddResourceButton
            // 
            this.AddResourceButton.Location = new System.Drawing.Point(222, 55);
            this.AddResourceButton.Name = "AddResourceButton";
            this.AddResourceButton.Size = new System.Drawing.Size(91, 32);
            this.AddResourceButton.TabIndex = 12;
            this.AddResourceButton.Text = "Add";
            this.AddResourceButton.UseVisualStyleBackColor = true;
            this.AddResourceButton.Click += new System.EventHandler(this.AddResourceButton_Click);
            // 
            // resourceListBox
            // 
            this.resourceListBox.FormattingEnabled = true;
            this.resourceListBox.ItemHeight = 16;
            this.resourceListBox.Location = new System.Drawing.Point(26, 55);
            this.resourceListBox.Margin = new System.Windows.Forms.Padding(4);
            this.resourceListBox.Name = "resourceListBox";
            this.resourceListBox.Size = new System.Drawing.Size(181, 420);
            this.resourceListBox.TabIndex = 11;
            // 
            // addResourceTextBox
            // 
            this.addResourceTextBox.Location = new System.Drawing.Point(26, 25);
            this.addResourceTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.addResourceTextBox.Name = "addResourceTextBox";
            this.addResourceTextBox.Size = new System.Drawing.Size(181, 22);
            this.addResourceTextBox.TabIndex = 10;
            // 
            // ManageResourcesForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(637, 513);
            this.Controls.Add(this.RemoveResourceButton);
            this.Controls.Add(this.AddResourceButton);
            this.Controls.Add(this.resourceListBox);
            this.Controls.Add(this.addResourceTextBox);
            this.Controls.Add(this.RemoveRoleButton);
            this.Controls.Add(this.AddRoleButton);
            this.Controls.Add(this.roleListBox);
            this.Controls.Add(this.RoleComboBox);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "ManageResourcesForm";
            this.Text = "Manage Resources";
            ((System.ComponentModel.ISupportInitialize)(this.automation_Task_ManagerDataSet2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.ComboBox RoleComboBox;
		private System.Windows.Forms.ListBox roleListBox;
		private System.Windows.Forms.Button AddRoleButton;
		private System.Windows.Forms.Button RemoveRoleButton;
		private Workload_Tracking_SystemDataSet automation_Task_ManagerDataSet2;
        private System.Windows.Forms.Button RemoveResourceButton;
        private System.Windows.Forms.Button AddResourceButton;
        private System.Windows.Forms.ListBox resourceListBox;
        private System.Windows.Forms.TextBox addResourceTextBox;
    }
}
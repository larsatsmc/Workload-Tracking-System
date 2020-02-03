namespace Toolroom_Project_Viewer
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
            this.RemoveResourceButton = new System.Windows.Forms.Button();
            this.AddResourceButton = new System.Windows.Forms.Button();
            this.resourceListBox = new System.Windows.Forms.ListBox();
            this.addResourceTextBox = new System.Windows.Forms.TextBox();
            this.panelControl1 = new DevExpress.XtraEditors.PanelControl();
            this.resourceTypeRadioGroup = new DevExpress.XtraEditors.RadioGroup();
            ((System.ComponentModel.ISupportInitialize)(this.panelControl1)).BeginInit();
            this.panelControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.resourceTypeRadioGroup.Properties)).BeginInit();
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
            "EDM Sinkers",
            "EDM Wires",
            "Rough CNC Operators",
            "Electrode CNC Operators",
            "Finish CNC Operators",
            "EDM Sinker Operators",
            "EDM Wire Operators",
            "Hole Popper Operators",
            "CMM Operators"});
            this.RoleComboBox.Location = new System.Drawing.Point(329, 23);
            this.RoleComboBox.Margin = new System.Windows.Forms.Padding(4);
            this.RoleComboBox.Name = "RoleComboBox";
            this.RoleComboBox.Size = new System.Drawing.Size(181, 25);
            this.RoleComboBox.TabIndex = 0;
            this.RoleComboBox.SelectedIndexChanged += new System.EventHandler(this.RoleComboBox_SelectedIndexChanged);
            // 
            // roleListBox
            // 
            this.roleListBox.FormattingEnabled = true;
            this.roleListBox.ItemHeight = 17;
            this.roleListBox.Location = new System.Drawing.Point(329, 53);
            this.roleListBox.Margin = new System.Windows.Forms.Padding(4);
            this.roleListBox.Name = "roleListBox";
            this.roleListBox.Size = new System.Drawing.Size(181, 412);
            this.roleListBox.TabIndex = 2;
            // 
            // AddRoleButton
            // 
            this.AddRoleButton.Location = new System.Drawing.Point(522, 53);
            this.AddRoleButton.Name = "AddRoleButton";
            this.AddRoleButton.Size = new System.Drawing.Size(91, 32);
            this.AddRoleButton.TabIndex = 5;
            this.AddRoleButton.Text = "Add";
            this.AddRoleButton.UseVisualStyleBackColor = true;
            this.AddRoleButton.Click += new System.EventHandler(this.AddRoleButton_Click);
            // 
            // RemoveRoleButton
            // 
            this.RemoveRoleButton.Location = new System.Drawing.Point(522, 91);
            this.RemoveRoleButton.Name = "RemoveRoleButton";
            this.RemoveRoleButton.Size = new System.Drawing.Size(91, 32);
            this.RemoveRoleButton.TabIndex = 6;
            this.RemoveRoleButton.Text = "Remove";
            this.RemoveRoleButton.UseVisualStyleBackColor = true;
            this.RemoveRoleButton.Click += new System.EventHandler(this.RemoveRoleButton_Click);
            // 
            // RemoveResourceButton
            // 
            this.RemoveResourceButton.Location = new System.Drawing.Point(220, 91);
            this.RemoveResourceButton.Name = "RemoveResourceButton";
            this.RemoveResourceButton.Size = new System.Drawing.Size(91, 32);
            this.RemoveResourceButton.TabIndex = 13;
            this.RemoveResourceButton.Text = "Remove";
            this.RemoveResourceButton.UseVisualStyleBackColor = true;
            this.RemoveResourceButton.Click += new System.EventHandler(this.RemoveResourceButton_Click);
            // 
            // AddResourceButton
            // 
            this.AddResourceButton.Location = new System.Drawing.Point(220, 53);
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
            this.resourceListBox.ItemHeight = 17;
            this.resourceListBox.Location = new System.Drawing.Point(24, 53);
            this.resourceListBox.Margin = new System.Windows.Forms.Padding(4);
            this.resourceListBox.Name = "resourceListBox";
            this.resourceListBox.Size = new System.Drawing.Size(181, 412);
            this.resourceListBox.TabIndex = 11;
            this.resourceListBox.SelectedValueChanged += new System.EventHandler(this.ResourceListBox_SelectedValueChanged);
            // 
            // addResourceTextBox
            // 
            this.addResourceTextBox.Location = new System.Drawing.Point(24, 23);
            this.addResourceTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.addResourceTextBox.Name = "addResourceTextBox";
            this.addResourceTextBox.Size = new System.Drawing.Size(181, 25);
            this.addResourceTextBox.TabIndex = 10;
            // 
            // panelControl1
            // 
            this.panelControl1.Appearance.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.panelControl1.Appearance.Options.UseFont = true;
            this.panelControl1.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
            this.panelControl1.Controls.Add(this.resourceTypeRadioGroup);
            this.panelControl1.Controls.Add(this.RemoveResourceButton);
            this.panelControl1.Controls.Add(this.AddResourceButton);
            this.panelControl1.Controls.Add(this.resourceListBox);
            this.panelControl1.Controls.Add(this.addResourceTextBox);
            this.panelControl1.Controls.Add(this.RemoveRoleButton);
            this.panelControl1.Controls.Add(this.AddRoleButton);
            this.panelControl1.Controls.Add(this.roleListBox);
            this.panelControl1.Controls.Add(this.RoleComboBox);
            this.panelControl1.Location = new System.Drawing.Point(2, 2);
            this.panelControl1.Name = "panelControl1";
            this.panelControl1.Size = new System.Drawing.Size(634, 510);
            this.panelControl1.TabIndex = 14;
            // 
            // resourceTypeRadioGroup
            // 
            this.resourceTypeRadioGroup.Location = new System.Drawing.Point(224, 156);
            this.resourceTypeRadioGroup.Name = "resourceTypeRadioGroup";
            this.resourceTypeRadioGroup.Properties.Columns = 2;
            this.resourceTypeRadioGroup.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
            new DevExpress.XtraEditors.Controls.RadioGroupItem("Person", "Person"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("Machine", "Machine")});
            this.resourceTypeRadioGroup.Properties.ItemsLayout = DevExpress.XtraEditors.RadioGroupItemsLayout.Flow;
            this.resourceTypeRadioGroup.Size = new System.Drawing.Size(85, 51);
            this.resourceTypeRadioGroup.TabIndex = 14;
            this.resourceTypeRadioGroup.EditValueChanged += new System.EventHandler(this.ResourceTypeRadioGroup_EditValueChanged);
            // 
            // ManageResourcesForm
            // 
            this.Appearance.Options.UseFont = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(637, 513);
            this.Controls.Add(this.panelControl1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "ManageResourcesForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Manage Resources";
            this.TopMost = true;
            ((System.ComponentModel.ISupportInitialize)(this.panelControl1)).EndInit();
            this.panelControl1.ResumeLayout(false);
            this.panelControl1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.resourceTypeRadioGroup.Properties)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox RoleComboBox;
        private System.Windows.Forms.ListBox roleListBox;
        private System.Windows.Forms.Button AddRoleButton;
        private System.Windows.Forms.Button RemoveRoleButton;
        //private Workload_Tracking_SystemDataSet automation_Task_ManagerDataSet2;
        private System.Windows.Forms.Button RemoveResourceButton;
        private System.Windows.Forms.Button AddResourceButton;
        private System.Windows.Forms.ListBox resourceListBox;
        private System.Windows.Forms.TextBox addResourceTextBox;
        private DevExpress.XtraEditors.PanelControl panelControl1;
        private DevExpress.XtraEditors.RadioGroup resourceTypeRadioGroup;
    }
}
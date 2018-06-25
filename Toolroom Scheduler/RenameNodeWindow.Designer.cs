namespace Toolroom_Scheduler
{
    partial class RenameNodeWindow
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
			this.BrowseForToolFolderButton = new System.Windows.Forms.Button();
			this.NewNameTextBox = new System.Windows.Forms.TextBox();
			this.OKButton = new System.Windows.Forms.Button();
			this.myCancelButton = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// BrowseForToolFolderButton
			// 
			this.BrowseForToolFolderButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.BrowseForToolFolderButton.Location = new System.Drawing.Point(337, 35);
			this.BrowseForToolFolderButton.Margin = new System.Windows.Forms.Padding(4);
			this.BrowseForToolFolderButton.Name = "BrowseForToolFolderButton";
			this.BrowseForToolFolderButton.Size = new System.Drawing.Size(97, 38);
			this.BrowseForToolFolderButton.TabIndex = 0;
			this.BrowseForToolFolderButton.Text = "Use Tool Folder Name";
			this.BrowseForToolFolderButton.UseVisualStyleBackColor = true;
			// 
			// NewNameTextBox
			// 
			this.NewNameTextBox.Location = new System.Drawing.Point(13, 43);
			this.NewNameTextBox.Margin = new System.Windows.Forms.Padding(4);
			this.NewNameTextBox.Name = "NewNameTextBox";
			this.NewNameTextBox.Size = new System.Drawing.Size(316, 22);
			this.NewNameTextBox.TabIndex = 1;
			// 
			// OKButton
			// 
			this.OKButton.Location = new System.Drawing.Point(114, 86);
			this.OKButton.Name = "OKButton";
			this.OKButton.Size = new System.Drawing.Size(97, 30);
			this.OKButton.TabIndex = 2;
			this.OKButton.Text = "OK";
			this.OKButton.UseVisualStyleBackColor = true;
			// 
			// myCancelButton
			// 
			this.myCancelButton.Location = new System.Drawing.Point(234, 86);
			this.myCancelButton.Name = "myCancelButton";
			this.myCancelButton.Size = new System.Drawing.Size(97, 30);
			this.myCancelButton.TabIndex = 3;
			this.myCancelButton.Text = "Cancel";
			this.myCancelButton.UseVisualStyleBackColor = true;
			// 
			// RenameNodeWindow
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(446, 142);
			this.Controls.Add(this.myCancelButton);
			this.Controls.Add(this.OKButton);
			this.Controls.Add(this.NewNameTextBox);
			this.Controls.Add(this.BrowseForToolFolderButton);
			this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.Margin = new System.Windows.Forms.Padding(4);
			this.Name = "RenameNodeWindow";
			this.Text = "Change Name";
			this.Load += new System.EventHandler(this.RenameNodeWindow_Load);
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BrowseForToolFolderButton;
        private System.Windows.Forms.TextBox NewNameTextBox;
        private System.Windows.Forms.Button OKButton;
        private System.Windows.Forms.Button myCancelButton;
    }
}
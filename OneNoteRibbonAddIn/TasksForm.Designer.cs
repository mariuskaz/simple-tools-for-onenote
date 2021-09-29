namespace OneNoteRibbonAddIn
{
    partial class TasksForm
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
            this.panel3 = new System.Windows.Forms.Panel();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.txtProject = new System.Windows.Forms.TextBox();
            this.txtInfo = new System.Windows.Forms.Label();
            this.todoProjects = new System.Windows.Forms.ComboBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.addLinks = new System.Windows.Forms.CheckBox();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.SystemColors.Control;
            this.panel3.Controls.Add(this.btnCancel);
            this.panel3.Controls.Add(this.btnOK);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel3.Location = new System.Drawing.Point(0, 192);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(402, 56);
            this.panel3.TabIndex = 14;
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.FlatAppearance.BorderColor = System.Drawing.Color.DarkGray;
            this.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(186)));
            this.btnCancel.Location = new System.Drawing.Point(292, 13);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(98, 30);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.btnOK.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnOK.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(186)));
            this.btnOK.Location = new System.Drawing.Point(188, 13);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(98, 30);
            this.btnOK.TabIndex = 0;
            this.btnOK.Text = "Add tasks";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // txtProject
            // 
            this.txtProject.AcceptsReturn = true;
            this.txtProject.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.txtProject.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtProject.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(186)));
            this.txtProject.HideSelection = false;
            this.txtProject.Location = new System.Drawing.Point(27, 64);
            this.txtProject.Multiline = true;
            this.txtProject.Name = "txtProject";
            this.txtProject.Size = new System.Drawing.Size(344, 50);
            this.txtProject.TabIndex = 2;
            this.txtProject.TabStop = false;
            this.txtProject.Text = "Project";
            // 
            // txtInfo
            // 
            this.txtInfo.AutoSize = true;
            this.txtInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(186)));
            this.txtInfo.Location = new System.Drawing.Point(24, 136);
            this.txtInfo.Name = "txtInfo";
            this.txtInfo.Size = new System.Drawing.Size(95, 16);
            this.txtInfo.TabIndex = 21;
            this.txtInfo.Text = "Tasks found: 0";
            // 
            // todoProjects
            // 
            this.todoProjects.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.todoProjects.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(186)));
            this.todoProjects.FormattingEnabled = true;
            this.todoProjects.Location = new System.Drawing.Point(27, 13);
            this.todoProjects.Name = "todoProjects";
            this.todoProjects.Size = new System.Drawing.Size(300, 24);
            this.todoProjects.TabIndex = 22;
            this.todoProjects.Text = "Create new project";
            this.todoProjects.SelectedIndexChanged += new System.EventHandler(this.changeTitle);
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.DimGray;
            this.panel2.Location = new System.Drawing.Point(27, 38);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(300, 1);
            this.panel2.TabIndex = 23;
            // 
            // addLinks
            // 
            this.addLinks.AutoSize = true;
            this.addLinks.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.addLinks.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(186)));
            this.addLinks.Location = new System.Drawing.Point(27, 160);
            this.addLinks.Name = "addLinks";
            this.addLinks.Size = new System.Drawing.Size(158, 22);
            this.addLinks.TabIndex = 2;
            this.addLinks.Text = "Add links to onenote";
            this.addLinks.UseVisualStyleBackColor = true;
            // 
            // TasksForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(402, 248);
            this.Controls.Add(this.addLinks);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.todoProjects);
            this.Controls.Add(this.txtInfo);
            this.Controls.Add(this.txtProject);
            this.Controls.Add(this.panel3);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "TasksForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Todoist";
            this.panel3.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.TextBox txtProject;
        private System.Windows.Forms.Label txtInfo;
        private System.Windows.Forms.ComboBox todoProjects;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.CheckBox addLinks;
    }
}
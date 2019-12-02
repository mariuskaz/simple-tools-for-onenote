using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OneNoteRibbonAddIn
{
    public partial class TasksForm : Form
    {
        public string project;
        public string prefix;

        public TasksForm(string title, string project, string prefix, string task)
        {
            InitializeComponent();
            this.Text = title;
            this.txtInfo.Text = title;
            this.txtProject.Text = project;
            this.txtPrefix.Text = prefix +" - ";
            this.txtTask.Text = task;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            project = txtProject.Text;
            prefix = txtPrefix.Text;
        }
    }
}

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
        public long id;

        public TasksForm(string title, string project, string prefix, string task, IEnumerable<Todoist.Net.Models.Project> projects)
        {
            InitializeComponent();
            this.Text = title;
            this.txtInfo.Text = title;
            this.txtProject.Text = project;
            this.txtPrefix.Text = prefix +" - ";
            this.txtTask.Text = task;
            this.todoProjects.Items.Add("Create new project");
            foreach (var item in projects)
            {
                this.todoProjects.Items.Add(item.Name);
            }
            this.project = this.txtProject.Text;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            project = txtProject.Text;
            prefix = txtPrefix.Text;
        }

        private void changeTitle(object sender, EventArgs e)
        {
            txtProject.Text = todoProjects.Text;
            id = todoProjects.SelectedIndex;
            if (id > 0)
            {
                txtProject.ReadOnly = true;
            }
            else
            {
                txtProject.ReadOnly = false;
                txtProject.Text = project;
            }
            todoProjects.SelectionStart = 2;
        }

    }
}

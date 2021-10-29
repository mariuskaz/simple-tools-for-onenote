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
        public int projectId;
        public bool links;
        public bool updates;

        public TasksForm(string project, string info, IEnumerable<Todoist.Net.Models.Project> projects)
        {
            InitializeComponent();
            this.txtInfo.Text = info;
            this.txtProject.Text = project;
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
            links = addLinks.Checked;
            updates = addOnlyNew.Checked;
        }

        private void changeTitle(object sender, EventArgs e)
        {
            txtProject.Text = todoProjects.Text;
            projectId = todoProjects.SelectedIndex;
            if (projectId == 0)
            {
                txtProject.ReadOnly = false;
                txtProject.Text = project;
            }
            else
            {
                txtProject.ReadOnly = true;
            }
            //todoProjects.SelectionStart = 2;
        }

    }
}

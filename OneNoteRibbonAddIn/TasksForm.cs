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
        public TasksForm(string title, string tasks)
        {
            InitializeComponent();
            this.Text = title;
            this.Tasks.Text = tasks;
        }

        string RemoveHtmlTags(string html)
        {
            return Regex.Replace(html, @"<(.|\n)*?>", "");
        }

    }
}

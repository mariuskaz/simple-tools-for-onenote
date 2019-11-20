using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace OneNoteRibbonAddIn
{
    [ComVisible(false)]
    public partial class MainForm : Form
    {
        private readonly OneNote.Application _oneNoteApp;

        public MainForm(OneNote.Application oneNoteApp)
        {
            _oneNoteApp = oneNoteApp;
            InitializeComponent();
        }


        private void btnLogin_Click(object sender, EventArgs e)
        {
            MessageBox.Show("login");
        }

        private void btncancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }



       

       
    }
}

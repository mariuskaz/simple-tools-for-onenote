using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace OneNoteRibbonAddIn
{
    [ComVisible(false)]
    public partial class LoginForm : Form
    {

        public string email;
        public string password;

        public LoginForm()
        {
            InitializeComponent();
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            email = txtLogin.Text;
            password = txtPassword.Text;
        }


    }
}

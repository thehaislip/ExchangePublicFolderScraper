using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ewsAPI
{
    public partial class GetSingleFolder : Form
    {
        private string _username;
        private string _password;
        private string _email;
        public GetSingleFolder()
        {
            InitializeComponent();
        }

        internal void SetCreds(string username, string password, string email)
        {
            _username = username;
            _password = password;
            _email = email;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var pf = new PublicFolder();
            var f = pf.GetFolderByPath(textBox1.Text,_username,_password,_email);
            var w = 1;
        }
    }
}

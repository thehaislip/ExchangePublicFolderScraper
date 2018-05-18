using ewsAPI.Models;
using MoreLinq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ewsAPI
{
    public partial class ScrapePublicFolders : Form
    {
        public ScrapePublicFolders()
        {
            InitializeComponent();
            uiWatch = new Stopwatch();
        }
        
        private Stopwatch uiWatch;
        private async void button1_Click(object sender, EventArgs e)
        {
            timer1.Interval = (1000) * (1);
            timer1.Tick += new EventHandler(timer_tick);

            button1.Enabled = false;
            progressBar1.Visible = true;
            progressBar1.Style = ProgressBarStyle.Marquee;

            var path = await ShowDialogAsync(saveFileDialog1);

            timer1.Start();
            uiWatch.Start();

            Task<string> task = Task.Run(() => GetPublicFolders(txtUserName.Text, txtPassword.Text, txtEmail.Text, path.ToString()));

            timer1.Stop();
            uiWatch.Stop();

            button1.Enabled = false;
            progressBar1.Visible = false;

            richTextBox1.Text = await task;
        }

        private void timer_tick(object sender, EventArgs e)
        {           
            labelTimer.Text = uiWatch.Elapsed.ToString();
        }

        private async Task<string> ShowDialogAsync(SaveFileDialog fileDialog)
        {
            await Task.Yield();
            string rtn;
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                rtn = Path.GetFullPath(fileDialog.FileName);
            }
            else
            {
                rtn = "";
            }
            return rtn;
        }

        private string GetPublicFolders(string username, string password, string email, string filePath)
        {
            var pf = new PublicFolder();
            try
            {
                var path = Path.GetDirectoryName(filePath);
                var fName = Path.GetFileNameWithoutExtension(filePath);
                var watch = System.Diagnostics.Stopwatch.StartNew();
                var f = pf.GetAllFolders(username, password, email).DistinctBy(e => e.FolderPath);
                watch.Stop();
                var em = watch.Elapsed;
                var csv = CSVWriter.ToCsv<PublicFolderModel>(",", f);
                var stat = $"time: {em.ToString()}; NumberOfItems:{f.Count()}";

                csv.WriteFile($"{path}{fName}.csv");
                return stat;

            }
            catch (Exception ex)
            {

                return ex.Message;
            }
        }
    }
}

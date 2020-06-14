using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CalculaImposto
{
    public partial class TelaConfiguracao : Form
    {
        public TelaConfiguracao()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog vistaFolderBrowserDialog = new FolderBrowserDialog();
            vistaFolderBrowserDialog.RootFolder = Environment.SpecialFolder.Desktop;
            vistaFolderBrowserDialog.Description = "Selecione o Dropbox";
            vistaFolderBrowserDialog.ShowNewFolderButton = false;

            if (vistaFolderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = vistaFolderBrowserDialog.SelectedPath;
            }
 
        }

        private void btnProximo_Click(object sender, EventArgs e)
        {
            FrmCalculaImposto frm = new FrmCalculaImposto();

            frm.Show();

            this.Hide();
        }
    }
}

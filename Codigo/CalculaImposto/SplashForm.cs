using System;
using System.Windows.Forms;

namespace CalculaImposto
{
    public partial class SplashForm : Form
    {
        public SplashForm()
        {
            InitializeComponent();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            panel2.Width += 3;
            if (panel2.Width >= 790)
            {
                timer1.Stop();

                string value = System.Configuration.ConfigurationManager.AppSettings["pastaDropbox"];

                //se app.config tiver o caminho do dropbox salvo
                if ((value!="valor") || (value != null))
                {
                    FrmCalculaImposto frm = new FrmCalculaImposto();

                    frm.Show();

                    this.Hide();
                }
                else
                {

                    TelaConfiguracao frm = new TelaConfiguracao();

                    frm.Show();

                    this.Hide();
                }
            }
        }
    }
}

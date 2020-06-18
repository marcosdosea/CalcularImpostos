using System;
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
            FolderBrowserDialog fBDialog = new FolderBrowserDialog();
            fBDialog.RootFolder = Environment.SpecialFolder.Desktop;
            fBDialog.Description = "Selecione o Dropbox";
            fBDialog.ShowNewFolderButton = false;

            if (fBDialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = fBDialog.SelectedPath;
            }
 
        }

        private void BtnProximo_Click(object sender, EventArgs e)
        {
            try
            {
                //observar se realmente um caminho foi selecionado, caso sim
                if (textBox1.Text !="")
                {
                    string pastaDropbox = textBox1.Text;
                    //atualizar o app.config
                    ConfiguracaoDropbox.UpdateAppSettings("pastaDropbox", pastaDropbox);

                    ConfiguracaoDropbox.UpdateAppConfig("appSettings", "value", pastaDropbox);

                    MessageBox.Show("Caminho do Dropbox salvo com sucesso!");
                    //chama a interface calculaimposto
                    FrmCalculaImposto frm = new FrmCalculaImposto();

                    frm.Show();

                    this.Hide();
                }
                else
                {
                    MessageBox.Show(String.Format("Selecione o caminho do dropbox primeiro!"));
                }
            } catch (Exception ex)
            {
                MessageBox.Show(String.Format("Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }
    }
}

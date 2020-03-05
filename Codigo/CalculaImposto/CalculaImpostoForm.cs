using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CalculaImposto
{
    public partial class FrmCalculaImposto : Form 
    {
        private string caminho;

        private string pastaTemp;

        public FrmCalculaImposto()
        {
            InitializeComponent();
        }

        private void BtnBuscar_Click(object sender, EventArgs e)
        {
            if (openFileDialogNfe.ShowDialog() == DialogResult.OK)
            {
                textBoxFile.Text = openFileDialogNfe.FileName;
            }
        }

        private void BtnBuscarNfe_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Arquivo zip | *.zip";
            openFileDialog.DefaultExt = "zip";
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog.FileName;
                caminho = textBox1.Text;

                if (string.IsNullOrEmpty(openFileDialog.FileName) == false)
                {
                    try
                    {

                        DescompactarZip(caminho);
                       // notasFiscaisBindingSource.DataSource = new List<NotasFiscais>();

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(String.Format("Não foi possível abrir o arquivo. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void DescompactarZip(string caminho)
        {
            pastaTemp = Path.GetTempPath();

            ZipFile.ExtractToDirectory(caminho, pastaTemp);
           
            DirectoryInfo dirInfo = new DirectoryInfo(pastaTemp);

            BuscaArquivos(dirInfo);
        }
        private void BuscaArquivos(DirectoryInfo dir)
        {
            // lista arquivos do diretorio corrente
            foreach (FileInfo file in dir.GetFiles("*.xml"))
            {
                try
                {
                  
                        TNfeProc nfe = new TNfeProc(); 
                        GerenciadorNfe gerenciadorNfe = new GerenciadorNfe();
                        nfe =gerenciadorNfe.LerNFE(file.FullName);
                        //chama novoObjeto para gerar um novo objeto do tipo Notas Fiscais e exibir na Grid
                        NotasFiscais novaNota = novoObjeto(nfe);
                        notasFiscaisBindingSource.DataSource = new List<NotasFiscais>();
                        //olhar como adicionar a List o novo objeto novaNota (retorno do método novoObjeto)

                } catch (Exception ex) 
                {
                    MessageBox.Show(String.Format("Não foi possível ler os arquivos na pasta temporária. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        public NotasFiscais novoObjeto(TNfeProc nfeProc) 
        {
            NotasFiscais notasFiscais = new NotasFiscais();
            notasFiscais.Numero = nfeProc.NFe.infNFe.ide.cNF;
            notasFiscais.DataEmissao = nfeProc.NFe.infNFe.ide.dhEmi;
            decimal converterValorProdutos = Convert.ToDecimal(nfeProc.NFe.infNFe.ide.verProc);
            decimal converterValorFrete = Convert.ToDecimal(nfeProc.NFe.infNFe.ide.cDV);
            decimal converterValorTotal = Convert.ToDecimal(nfeProc.NFe.infNFe.ide.cDV);
            notasFiscais.ValorProdutos = converterValorProdutos;
            notasFiscais.ValorFrete = converterValorFrete;
            notasFiscais.ValorTotal = converterValorTotal;
            return notasFiscais;
        }

        private void notasFiscaisBindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }
    }
}

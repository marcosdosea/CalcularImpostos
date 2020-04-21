using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Windows.Forms;


namespace CalculaImposto
{
    public partial class FrmCalculaImposto : Form
    {
        private string caminho;

        string pastaSaida;

        string subdiretorio;

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
            openFileDialog.Filter = "Arquivos | *.zip;*.xml";
            openFileDialog.DefaultExt = "zip";
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog.FileName;
                caminho = textBox1.Text;

                if (string.IsNullOrEmpty(openFileDialog.FileName) == false)
                {
                    if (caminho.EndsWith(".zip"))
                    {
                        try
                        {

                            ExtrairZip(caminho);
                            DirectorioExiste(pastaSaida);

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(String.Format("Não foi possível abrir o arquivo. Excecão OpenFileDialog. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    } else
                    {
                        //pastaSaida = caminho;
                        ProcessarArquivo();
                    }
                }
            }
        }
        public NotasFiscais NovoObjeto(TNfeProc nfeProc)
        {
            try
            {
                NotasFiscais notasFiscais = new NotasFiscais();
                notasFiscais.Numero = nfeProc.NFe.infNFe.ide.nNF;
                notasFiscais.Fornecedor = nfeProc.NFe.infNFe.emit.IE;

                notasFiscais.DataEmissao = nfeProc.NFe.infNFe.ide.dhEmi;
                DateTime converterData = Convert.ToDateTime(notasFiscais.DataEmissao);
                string strDate = converterData.ToString("dd/MM/yyyy");
                notasFiscais.DataEmissao = strDate;

                decimal converterValorProdutos = Convert.ToDecimal(nfeProc.NFe.infNFe.total.ICMSTot.vProd);
                decimal converterValorFrete = Convert.ToDecimal(nfeProc.NFe.infNFe.total.ICMSTot.vFrete);
                decimal converterValorTotal = Convert.ToDecimal(nfeProc.NFe.infNFe.total.ICMSTot.vNF);

                notasFiscais.ValorProdutos = converterValorProdutos;
                notasFiscais.ValorFrete = converterValorFrete;
                notasFiscais.ValorTotal = converterValorTotal;

                return notasFiscais;
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Não foi possível criar o objeto NotasFiscais. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        private void NotasFiscaisBindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }

        #region Métodos Arquivo
        /// <summary>
        /// Verifica se o diretorio existe. Em caso afirmativo, começa a leitura dos arquivos existentes dentro dele, chamando o outro método ProcessarArquivos()
        /// </summary>
        /// <param name="pasta">Passa o caminho da pasta a ser verificada</param>
        public void DirectorioExiste(string pasta)
        {
            if (Directory.Exists(pasta))
            {
                    ProcessarArquivos();
            }
        }

        public void ProcessarArquivos()
        {
            try
            {
                string[] arquivos = Directory.GetFiles(pastaSaida, "*.xml");
                TNfeProc nfe;
                NotasFiscais novaNota;
                GerenciadorNfe gerenciadorNfe;
                List<NotasFiscais> notaList = new List<NotasFiscais>();
                foreach (var file in arquivos)
                {
                    nfe = new TNfeProc();

                    gerenciadorNfe = new GerenciadorNfe();

                    nfe = gerenciadorNfe.LerNFE(file);

                    novaNota = NovoObjeto(nfe);

                    notaList.Add(novaNota);
                }

                this.notasFiscaisBindingSource.DataSource = notaList;

                this.dataGridView1.DataSource =
                   this.notasFiscaisBindingSource;
             //   this.dataGridView1.Columns["ValorFrete"].DefaultCellStyle.Format = "0:C2";
             //   this.dataGridView1.Columns["ValorProdutos"].DefaultCellStyle.Format = "0:C2";
             //   this.dataGridView1.Columns["ValorTotal"].DefaultCellStyle.Format = "0:C2";
              //   this.dataGridView1.Columns["DataEmissao"].DefaultCellStyle.Format = "dd/MM/yy";
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Não foi possível processar os arquivos do diretorio. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ProcessarArquivo()
        {
            try
            {
                string arquivo = Path.GetFullPath(caminho);
                TNfeProc nfe;
                NotasFiscais novaNota;
                GerenciadorNfe gerenciadorNfe;
             
                nfe = new TNfeProc();

                gerenciadorNfe = new GerenciadorNfe();

                nfe = gerenciadorNfe.LerNFE(arquivo);

                novaNota = NovoObjeto(nfe);

                this.notasFiscaisBindingSource.DataSource = novaNota;

                this.dataGridView1.DataSource =
                   this.notasFiscaisBindingSource;
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Não foi possível processar o arquivo. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void CriarDirectorio(string path)
        {

            try
            {
                if (Directory.Exists(path))
                {
                    ApagarDiretorio(path);
                }

                DirectoryInfo di = Directory.CreateDirectory(path);
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Não foi possível criar Diretorio. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ApagarDiretorio(string path)
        {
            try
            {
                string[] fileEntries = Directory.GetFiles(path);
                foreach (string fileName in fileEntries)
                {
                    ApagarArquivo();
                }
                Directory.Delete(path, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Não foi possível excluir o diretorio. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApagarArquivo()
        {
            try
            {

                string[] arquivos = Directory.GetFiles(pastaSaida, "*.xml");
                foreach (var file in arquivos)
                {
                    File.Delete(file);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Não foi possível excluir o arquivo do diretorio. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ExtrairZip(string pastaZip)
        {
            try
            {
                pastaSaida = "DiretorioTemporario";
                CriarDirectorio(pastaSaida);
                ZipFile.ExtractToDirectory(pastaZip, pastaSaida);
                //checar se existe outra pasta dentro do diretorio extraido
                bool valor = ChecarSubpasta(pastaSaida);
                if (valor.Equals(true))
                {
                    pastaSaida = subdiretorio;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Não foi possível extrair os arquivos. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool ChecarSubpasta(string pasta)
        {
            if (Directory.Exists(pasta))
            {
                string[] diretorios = Directory.GetDirectories(pastaSaida);
                foreach (string dir in diretorios)
                {
                    subdiretorio = dir;
                    return true;
                }
            }
            return false;
        }
        #endregion

       
    }
}

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
using System.Security.AccessControl;
using System.Threading;

namespace CalculaImposto
{
    public partial class FrmCalculaImposto : Form
    {
        private string caminho;

        private string pastaTemp = null;

        private DirectoryInfo dirInfo = null;

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
                     /*   string user = "barbie-12-girl@hotmail.com"; // Troque X pelo nome de usuário de um administrador.
                        System.IO.DirectoryInfo folderInfo = new System.IO.DirectoryInfo(caminho);
                        DirectorySecurity ds = new DirectorySecurity();

                        ds.AddAccessRule(new FileSystemAccessRule(user, FileSystemRights.Modify, AccessControlType.Allow));
                        ds.SetAccessRuleProtection(false, false);
                        folderInfo.SetAccessControl(ds);
                        */
                        DescompactarZip(caminho);
                        // notasFiscaisBindingSource.DataSource = new List<NotasFiscais>();

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(String.Format("Não foi possível abrir o arquivo. Excecão OpenFileDialog. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void DescompactarZip(string caminho)
        {
            try
            {

                pastaTemp = Path.GetTempPath();

                /*   if (Directory.Exists(pastaTemp))
                   {
                       dirInfo = new DirectoryInfo(pastaTemp);
                       // Busca automaticamente todos os arquivos em todos os subdiretórios
                       FileInfo[] Files = dirInfo.GetFiles("*xml", SearchOption.AllDirectories);
                       foreach (FileInfo File in Files)
                       {
                           BuscaArquivos(dirInfo);
                       }
                   }

                   else
                   { */
                    ZipFile.ExtractToDirectory(caminho, pastaTemp);

              //      dirInfo = new DirectoryInfo(pastaTemp);
              
              //      BuscaArquivos(dirInfo);
               // }
               }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Não foi possível descompactar e abrir os arquivos na pasta temporária. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } 
            finally 
            {
                 Directory.GetFiles(pastaTemp, "*", SearchOption.AllDirectories);
                 File.Delete(pastaTemp);
                 Directory.Delete(pastaTemp, true);
             //   DeleteDirectory(pastaTemp);
            }

        }
        public static void DeleteDirectory(string path)
        {
            foreach (string directory in Directory.GetDirectories(path))
            {
                Thread.Sleep(1);
                DeleteDir(directory);
            }
            DeleteDir(path);
        }

        private static void DeleteDir(string dir)
        {
            try
            {
                Thread.Sleep(1);
                Directory.Delete(dir, true);
            }
            catch (IOException)
            {
                DeleteDir(dir);
            }
            catch (UnauthorizedAccessException)
            {
                DeleteDir(dir);
            }
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
                    nfe = gerenciadorNfe.LerNFE(file.FullName);
                    //chama novoObjeto para gerar um novo objeto do tipo Notas Fiscais e exibir na Grid
                    NotasFiscais novaNota = novoObjeto(nfe);
                    //    notasFiscaisBindingSource.DataSource = new List<NotasFiscais>();
                    //   notasFiscaisBindingSource.Add(novaNota);

                    List<NotasFiscais> notaList = new List<NotasFiscais>();
                    notaList.Add(novaNota);

                    this.notasFiscaisBindingSource.DataSource = notaList;

                    this.dataGridView1.DataSource =
                        this.notasFiscaisBindingSource;

                }
                catch (Exception ex)
                {
                    MessageBox.Show(String.Format("Não foi possível ler os arquivos na pasta temporária. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        public NotasFiscais novoObjeto(TNfeProc nfeProc)
        {
            try
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
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Não foi possível criar o objeto NotasFiscais. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        private void notasFiscaisBindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }
    }
}

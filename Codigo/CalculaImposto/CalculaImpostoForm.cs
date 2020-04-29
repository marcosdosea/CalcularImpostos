using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

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
                    }
                    else
                    {
                        //passei a pasta caminho;
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
                // notasFiscais.Numero = nfeProc.NFe.infNFe.ide.nNF;
                notasFiscais.Fornecedor = nfeProc.NFe.infNFe.emit.IE;

                notasFiscais.DataEmissao = nfeProc.NFe.infNFe.ide.dhEmi;
                DateTime converterData = Convert.ToDateTime(notasFiscais.DataEmissao);
                string strDate = converterData.ToString("dd/MM/yyyy");
                notasFiscais.DataEmissao = strDate;

                string valorProdutos = nfeProc.NFe.infNFe.total.ICMSTot.vProd;
                var vp = valorProdutos.Replace('.', ',');
                string valorFrete = nfeProc.NFe.infNFe.total.ICMSTot.vFrete;
                var vf = valorFrete.Replace('.', ',');
                string valorTotal = nfeProc.NFe.infNFe.total.ICMSTot.vNF;
                var vt = valorTotal.Replace('.', ',');

                notasFiscais.ValorProdutos = Math.Round(Convert.ToDecimal(vp), 2);
                notasFiscais.ValorFrete = Math.Round(Convert.ToDecimal(vf), 2);
                notasFiscais.ValorTotal = Math.Round(Convert.ToDecimal(vt), 2);

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
                Imposto imposto;
                GerenciadorNfe gerenciadorNfe;
                List<NotasFiscais> notaList = new List<NotasFiscais>();
                List<Imposto> impostoList = new List<Imposto>();
                int quantidadeProdutosNotaF;
                string arquivo;

                foreach (var file in arquivos)
                {
                    nfe = new TNfeProc();

                    gerenciadorNfe = new GerenciadorNfe();

                    nfe = gerenciadorNfe.LerNFE(file);

                    novaNota = NovoObjeto(nfe);

                    notaList.Add(novaNota);

                    int pos = 0;

                    arquivo = Path.GetFullPath(file);

                    quantidadeProdutosNotaF = ContaItensNotaFiscal(arquivo);

                    for (int i = 0; i <= quantidadeProdutosNotaF - 1; i++)
                    {
                        imposto = ImpostoNotaFiscal(pos, nfe);
                        impostoList.Add(imposto);
                        pos++;
                    }
                }

                this.notasFiscaisBindingSource.DataSource = notaList;

                this.dataGridView1.DataSource =
                   this.notasFiscaisBindingSource;

                this.impostoBindingSource.DataSource = impostoList;
                this.dataGridView2.DataSource =
                    this.impostoBindingSource.DataSource;

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
                Imposto imposto;
                int quantidadeProdutosNotaF = ContaItensNotaFiscal(caminho);
                int pos = 0;
                List<Imposto> impostoList = new List<Imposto>();
                GerenciadorNfe gerenciadorNfe;

                nfe = new TNfeProc();

                gerenciadorNfe = new GerenciadorNfe();

                nfe = gerenciadorNfe.LerNFE(arquivo);

                novaNota = NovoObjeto(nfe);

                for (int i = 0; i <= quantidadeProdutosNotaF - 1; i++)
                {
                    imposto = ImpostoNotaFiscal(pos, nfe);
                    impostoList.Add(imposto);
                    pos++;
                }

                this.notasFiscaisBindingSource.DataSource = novaNota;

                this.dataGridView1.DataSource =
                   this.notasFiscaisBindingSource;

                this.impostoBindingSource.DataSource = impostoList;
                this.dataGridView2.DataSource =
                    this.impostoBindingSource.DataSource;
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

        #region Tarefa 2 - Segunda Grid

        public int ContaItensNotaFiscal(string pasta)
        {
            try
            {
                XmlDocument arquivo = new XmlDocument();
                arquivo.Load(pasta);
                var totalItens = arquivo.GetElementsByTagName("prod");
                int itensNF = 0;
                foreach (XmlElement nodo in totalItens)
                {
                    itensNF++;
                }
                return itensNF;
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Não foi possível contar os produtos da nota fiscal. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }
        }
        /// <summary>
        /// O objetivo desse método é pegar a aliquota de origem no nó ICMS da nota fiscal. 
        /// </summary>
        /// <param name="pasta">Passo o caminho na pasta do arquivo XML que será carregado pelo load</param>
        /// <returns></returns>
        public string RetornaAliquotaOrigemICMS(string pasta, int pos)
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(pasta);
                //string origem = "";
                XmlNodeList elemList = doc.GetElementsByTagName("ICMS10");
                // var orig=elemList[pos];
                var recuperaItem = elemList.Item(pos);
                foreach (XmlNode node in elemList)
                {
                    XmlNodeList ali = doc.GetElementsByTagName("orig");
                    recuperaItem = ali.Item(pos);
                   // orig = ali[pos];  
                }
                string format = recuperaItem.OuterXml;
                return format;
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Não foi possível obter aliquota de origem. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        public Imposto ImpostoNotaFiscal(int pos, TNfeProc nfeProc)
        {
            try
            {

                Imposto imposto = new Imposto();

                imposto.Numero = nfeProc.NFe.infNFe.ide.nNF;

                imposto.NCM = nfeProc.NFe.infNFe.det[pos].prod.NCM;
                imposto.Produto = nfeProc.NFe.infNFe.det[pos].prod.xProd; //peguei o nome
                imposto.TipoReceita = nfeProc.NFe.infNFe.ide.natOp; //natureza operação?

                //  imposto.AliquotaOrigem = Convert.ToDecimal(nfeProc.NFe.infNFe.det[pos].imposto);
                //  imposto.AliquotaDestino = Convert.ToDecimal(nfeProc.NFe.infNFe.det[pos].imposto.ICMSUFDest.pICMSInter);

                //chamar método para obter aliquota de origem ICMS
              /*  string valorAliquotaOrigem = RetornaAliquotaOrigemICMS(caminho, pos);
                var vAO = valorAliquotaOrigem.Replace('.', ',');
                imposto.AliquotaOrigem = Convert.ToDecimal(vAO);*/

                
                return imposto;
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Não foi possível criar o objeto Imposto. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        #endregion
    }
}

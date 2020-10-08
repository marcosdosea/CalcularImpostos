using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Xml;
using Microsoft.TeamFoundation.Common;

namespace CalculaImposto
{
    public partial class FrmCalculaImposto : Form
    {
        #region variáveis 
        private string caminho;

        private string pastaSaida;

        private string subdiretorio;

        private string dataAtualMVAFormatada;

        private string buscaAliquotaOrigem;

        private string aliD = "18";

        private DateTime dataAtualizacaoMVA;

        private decimal MVA;

        private string value = ConfiguracaoDropbox.GetValue("pastaDropbox");

        private string valorProdutoUnitario;

        private string valorICMSOrigem;

        private decimal precoGoverno;
        #endregion
        public FrmCalculaImposto()
        {
            InitializeComponent();
        }

        private void BtnBuscar_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fBDialog = new FolderBrowserDialog();
            fBDialog.RootFolder = Environment.SpecialFolder.Desktop;
            fBDialog.Description = "Selecione o Dropbox";
            fBDialog.ShowNewFolderButton = false;

            if (fBDialog.ShowDialog() == DialogResult.OK)
            {
                textBoxFile.Text = fBDialog.SelectedPath;

                string pastaDropbox = textBoxFile.Text;
                //atualizar o app.config
                ConfiguracaoDropbox.UpdateAppSettings("pastaDropbox", pastaDropbox);

                MessageBox.Show("Caminho do Dropbox salvo com sucesso!");
            }
            else
            {

                MessageBox.Show(string.Format("Selecione o caminho do dropbox primeiro!"));
            }
        }

        private void BtnBuscarNfe_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                Filter = "Arquivos | *.zip;*.xml",
                DefaultExt = "zip",
                RestoreDirectory = true
            };

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
                        //passei a pasta 'caminho';
                        ProcessarArquivo();
                    }
                }
            }
        }
        public string RetornaCnpjFornecedor(string pasta, int pos)
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(pasta);
                XmlNodeList elemList = doc.GetElementsByTagName("emit");
                var recuperaItem = elemList.Item(pos);

                foreach (XmlNode node in elemList)
                {
                    XmlNodeList ali = doc.GetElementsByTagName("CNPJ");
                    recuperaItem = ali.Item(pos);
                }
                
                    string format = recuperaItem.OuterXml;
                    format = format.Replace("<CNPJ xmlns=\"http://www.portalfiscal.inf.br/nfe\">", "");
                    format = format.Replace("</CNPJ>", "");
                   
                    return format;
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Não foi possível obter o CNPJ do fornecedor. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public NotasFiscais NovoObjeto(TNfeProc nfeProc,string arquivo,int posicao)
        {
            try
            {         
                NotasFiscais notasFiscais = new NotasFiscais();
                notasFiscais.Numero = nfeProc.NFe.infNFe.ide.nNF;
                // notasFiscais.Numero = nfeProc.NFe.infNFe.ide.nNF;
                notasFiscais.NomeFornecedor = nfeProc.NFe.infNFe.emit.xNome;
                notasFiscais.CnpjFornecedor = RetornaCnpjFornecedor(arquivo, posicao);
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
                NotasFiscais novaNota;
                List<NotasFiscais> notaList = new List<NotasFiscais>();
                List<Imposto> impostoList = new List<Imposto>();
                int quantidadeProdutosNotaF;

                foreach (var file in arquivos)
                {
                    int pos = 0;
                    novaNota = NovoObjeto(DesserializarNota(file),file,pos);

                    notaList.Add(novaNota);

                    quantidadeProdutosNotaF = ContaItensNotaFiscal(file);
 
                    for (int i = 0; i <= quantidadeProdutosNotaF - 1; i++)
                    {
                        buscaAliquotaOrigem = RetornapAliquotaOrigem(file, i);

                        if (String.IsNullOrWhiteSpace(buscaAliquotaOrigem))
                        {
                            buscaAliquotaOrigem = "12";
                         
                        }
                        Imposto imposto = ImpostoNotaFiscal(i, DesserializarNota(file));
                      
                        if (impostoList.Count==0)
                        {
                            impostoList.Add(imposto);
                        }
                        else 
                        {
                           
                                Boolean verifica = PesquisaImpostoList(impostoList, imposto.NCM, Convert.ToDecimal(buscaAliquotaOrigem));
                                if (verifica.Equals(false))
                                {
                                    impostoList.Add(imposto);
                                }
                         
                        }
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
        private Boolean PesquisaImpostoList(List<Imposto> imposto, string ncm, decimal aliOrigem)
        {
            Boolean resultado = false;
            for (int i = 0; i < imposto.Count; i++)
            {
                if (ncm.Equals(imposto[i].NCM) && aliOrigem.Equals(imposto[i].AliquotaOrigem))
                {
                    resultado = true;
                    break;
                }  
            }
            return resultado;
        }

        public void ProcessarArquivo()
        {
            try
            {
                string arquivo = Path.GetFullPath(caminho);
                NotasFiscais novaNota;
                Imposto imposto;
                int quantidadeProdutosNotaF = ContaItensNotaFiscal(arquivo);
                int pos = 0;
                List<Imposto> impostoList = new List<Imposto>();
              
                novaNota = NovoObjeto(DesserializarNota(arquivo),arquivo,pos);
                for (int i = 0; i <= quantidadeProdutosNotaF - 1; i++)
                {
                    buscaAliquotaOrigem = RetornapAliquotaOrigem(arquivo, pos);
                   
                    imposto = ImpostoNotaFiscal(pos, DesserializarNota(arquivo));
                    Boolean verifica = PesquisaImpostoList(impostoList, imposto.NCM, imposto.AliquotaOrigem);
                    if (verifica.Equals(false))
                    {
                        impostoList.Add(imposto);
                    }
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

        public TNfeProc DesserializarNota(string arquivo)
        {
            try
            {
                TNfeProc nfe = new TNfeProc();

                GerenciadorNfe gerenciadorNfe = new GerenciadorNfe();

                nfe = gerenciadorNfe.LerNFE(arquivo);

                return nfe;
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Não foi possível desserializar nota. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
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
                    ApagarArquivo(path); 
                }
                Directory.Delete(path, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Não foi possível excluir o diretorio. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApagarArquivo(string path)
        {
            try
            {
                //   string[] arquivos = Directory.GetFiles(path, "*.xml");

                string[] arquivos = Directory.GetFiles(path);
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
        /// <summary>
        /// Cada nota fiscal possui um ou mais produtos. Esse método conta a quantidade de produtos em cada nota Fiscal;
        /// </summary>
        /// <param name="pasta">O caminho completo até o arquivo "*.xml" a ser lido</param>
        /// <returns></returns>
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
       
        public string RetornapAliquotaOrigem(string pasta, int pos)
        {
            try
            {
                XmlDocument xml = new XmlDocument();
                xml.Load(pasta);
                XmlNodeList icmsList = xml.GetElementsByTagName("ICMS");
                var recuperaItem = " ";
                var format = "";
               
                foreach (XmlNode det in icmsList[pos])
                {
                    foreach (XmlNode n in det)
                    {
                        if (n.Name == "pICMS")
                        {
                            recuperaItem = n.InnerText; 
                        }             
                    }
                }
                if (recuperaItem != " ")
                {
                    format = recuperaItem;
                    format = format.Replace(".", " ");
                    format = format.TrimEnd('0', ' ');
                }
                
                return format;
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Não foi possível retornar a porcentagem da aliquota de origem na nota fiscal. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        /// <summary>
        /// O objetivo desse método é pegar a aliquota de origem no nó ICMS da nota fiscal. 
        /// </summary>
        /// <param name="pasta">Passo o caminho na pasta do arquivo XML que será carregado pelo load</param>
        /// <returns></returns>
        public string RetornavalorICMS(string pasta, int pos)
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(pasta);
                XmlNodeList elemList = doc.GetElementsByTagName("ICMS");
                var recuperaItem = elemList.Item(pos);

                foreach (XmlNode node in elemList)
                {
                    XmlNodeList ali = doc.GetElementsByTagName("vICMS");
                    recuperaItem = ali.Item(pos);

                }
                if (recuperaItem == null)
                {
                    //se o produto não tiver o vICMS
                    return null;
                }
                else
                {
                    string format = recuperaItem.OuterXml;
                    format = format.Replace("<vICMS xmlns=\"http://www.portalfiscal.inf.br/nfe\">", "");
                    format = format.Replace("</vICMS>", "");
                    format = format.Replace(".", ",");

                    return format;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Não foi possível obter valor ICMS. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public string BuscaArquivoTxt()
        {
            try
            {
                string arquivo = " ";
                var fileList = new DirectoryInfo(value).GetFiles("ResumoNotasFiscais*", SearchOption.AllDirectories);
                foreach (FileInfo file in fileList)
                {

                    arquivo = file.FullName;
                }
                if (arquivo == " ")
                {
                    arquivo = null;
                }
                return arquivo;

            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Não foi possível encontrar o arquivo .txt no diretório. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        public Tuple<string, string, string, string, string, int> LerArquivoTxt(string nomeArquivo, string ncm, string aliO, string aliD)
        {
            try
            {
                string valorMVA = " ";
                string dataAtualizacao = " ";
                string aliquotaOrigem = " ";
                string aliquotaDestino = " ";
                string[] colunas;
                Boolean proximalinha = true;
                string trim;
                string[] teste = File.ReadAllLines(nomeArquivo);
                string ncmArquivo = null;
                int linha = 0;
                for (int j = 0; j < teste.Length; j++)
                {
                    if (teste[j].Contains(ncm))
                    {
                        linha = j;
                        trim = teste[j].TrimEnd(',', ' ');
                        colunas = trim.Split(',');
                        for (int i = 0; i < colunas.Length; i++)
                        {
                            ncmArquivo = colunas[0].Replace(',', ' ');
                            aliquotaOrigem = colunas[1].Replace(',', ' ');
                            aliquotaDestino = colunas[2].Replace(',', ' ');
                            dataAtualizacao = colunas[3].Replace(',', ' ');
                            valorMVA = colunas[4];
                            if (colunas.Length == 6)
                            {
                                valorMVA = valorMVA + "," + colunas[5].Replace(',', ' ');
                            }
                            else
                            {
                                valorMVA = valorMVA.Replace(',', ' ');
                            }

                            proximalinha = false;
                            break;
                        }
                    }
                }

                if (proximalinha == true)
                {
                    dataAtualizacao = "";
                    valorMVA = "";
                }

                return new Tuple<string, string, string, string, string, int>(dataAtualizacao, valorMVA, aliquotaOrigem, aliquotaDestino, ncmArquivo, linha);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Não foi possível ler o arquivo Txt. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        /// <summary>
        /// Cria um novo objeto do tipo Imposto. Cada nota Fiscal possui um ou vários produtos. Cada produto possui imposto. 
        /// </summary>
        /// <param name="pos">Passa a posição do item da lista de produtos na nota fiscal que está sendo lido</param>
        /// <param name="nfeProc">Um objeto do tipo TNfeProc</param>
        /// <returns></returns>
        public Imposto ImpostoNotaFiscal(int pos, TNfeProc nfeProc)
        {
            try
            {
                // checar se já existe um arquivo resumo nota fiscal salvo no dropbox local 
                string recuperaArquivo = BuscaArquivoTxt();

                Imposto imposto = new Imposto();
                imposto.NCM = nfeProc.NFe.infNFe.det[pos].prod.NCM;
                imposto.TipoReceita = nfeProc.NFe.infNFe.ide.natOp; //natureza operação? 
                imposto.AliquotaOrigem = Convert.ToDecimal(buscaAliquotaOrigem);
                // aliD = "18" -> valor atual da alíquota de destino de sergipe;
                imposto.AliquotaDestino = Convert.ToDecimal(aliD);

                if (recuperaArquivo != " " && recuperaArquivo != null)
                {
                    var tupla = LerArquivoTxt(recuperaArquivo, imposto.NCM, buscaAliquotaOrigem, aliD);
                    string mva = tupla.Item2;
                    if (String.IsNullOrEmpty(mva).Equals(false))
                    {
                        imposto.DataAtualizacaoMVA = tupla.Item1;
                        imposto.MVA = Convert.ToDecimal(tupla.Item2);
                    }
                }
                return imposto;
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Não foi possível criar o objeto Imposto. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        public void AtualizaArquivoTXT(string pasta)
        {
            string dataAtualizacao = null;
            for (int i = 0; i <= dataGridView2.RowCount - 1; i++)
            {
                string ncm = dataGridView2.Rows[i].Cells[0].Value.ToString();
                string aliO = dataGridView2.Rows[i].Cells[2].Value.ToString();
                string aliD = dataGridView2.Rows[i].Cells[3].Value.ToString();
                string mva = dataGridView2.Rows[i].Cells[5].Value.ToString();
                //    string dataAtualizacao = dataGridView2.Rows[i].Cells[6].Value;
                var tupla = LerArquivoTxt(pasta, ncm, aliO, aliD);
                string ncmArquivo = tupla.Item5;
                string dataAtualizacaoArquivo = tupla.Item1;
                string mvaArquivo = tupla.Item2;
                string aliOrigemArquivo = tupla.Item3;
                string aliDestinoArquivo = tupla.Item4;
                int linha = tupla.Item6;
                string novaLinha = "";

                if (ncm != ncmArquivo) //se o ncm ainda não tiver gravado no txt
                {
                    StreamWriter s = File.AppendText(pasta);
                    novaLinha = novaLinha + ncm + ","
                + aliO + ","
                + aliD + ","
                + dataGridView2.Rows[i].Cells[4].Value + ","
                + mva;
                    novaLinha = novaLinha + ",";
                    s.WriteLine(novaLinha);
                    s.Close();
                }
                if (dataGridView2.Rows[i].Cells[4].Value != null)
                {
                    dataAtualizacao = dataGridView2.Rows[i].Cells[4].Value.ToString();
                }
                if (ncm.Equals(ncmArquivo) && dataAtualizacaoArquivo != dataAtualizacao)
                {
                    if (dataAtualizacaoArquivo.IsNullOrEmpty().Equals(false))
                    {
                        DateTime dataGrid = DateTime.ParseExact(dataAtualizacao, "dd/MM/yyyy", null);
                        DateTime dataArq = DateTime.ParseExact(dataAtualizacaoArquivo, "dd/MM/yyyy", null);
                        int result = DateTime.Compare(dataArq, dataGrid);

                        if (result < 0)
                        {
                            string sLine = "";
                            sLine = sLine + ncm + "," +
                           aliO + "," +
                           aliD + "," +
                           dataAtualizacao + "," +
                           mva;
                            sLine = sLine + ",";
                            AlteraLinha(pasta, ncm, sLine, linha);
                        }
                    }
                    else
                    {
                        string sLine = "";
                        sLine = sLine + ncm + "," +
                       aliO + "," +
                       aliD + "," +
                       dataAtualizacao + "," +
                       mva;
                        sLine = sLine + ",";
                        AlteraLinha(pasta, ncm, sLine, linha);
                    }
                }
                if (ncm.Equals(ncmArquivo) && dataAtualizacaoArquivo == dataAtualizacao)
                {
                    if (mva != mvaArquivo || aliO != aliOrigemArquivo || aliD != aliDestinoArquivo)
                    {
                        string sLine = "";
                        sLine = sLine + ncm + "," +
                       aliO + "," +
                       aliD + "," +
                       dataAtualizacao + "," +
                       mva;
                        sLine = sLine + ",";
                        AlteraLinha(pasta, ncm, sLine, linha);
                    }
                }
            }
        }

        private void AlteraLinha(string pasta, string ncm, string linhaAtualizada, int posicaolinha)
        {

            string[] lines = File.ReadAllLines(pasta);

            if (lines.Length == 0)
            {
                MessageBox.Show("Seu arquivo está vazio!");
                return;
            }

            using (StreamWriter writer = new StreamWriter(pasta))
            {
                for (int i = 0; i < lines.Length; i++)
                {
                    // Verifica se é a segunda linha e se o conteúdo da mesma é igual ao usuário atual
                    if (i == posicaolinha && lines[i].Contains(ncm))
                    {
                        writer.WriteLine(linhaAtualizada);
                    }
                    else
                    {
                        writer.WriteLine(lines[i]);
                    }
                }
            }
        }

        private string GerarNomeArquivoTXT()
        {
            string nomeArquivo = "ResumoNotasFiscais.txt";
            string caminhoCompleto = "";
            try
            {
                caminhoCompleto = Path.Combine(value, nomeArquivo);
                return caminhoCompleto;
            }
            catch (Exception ex)
            {
                if (value == null)
                    value = "null";
                if (nomeArquivo == null)
                    nomeArquivo = "null";
                MessageBox.Show(string.Format("Você não pode combinar '{0}' e '{1}' porque: {2}{3}", value, nomeArquivo, Environment.NewLine, ex.Message));
                return null;
            }
        }
        private void BtnSalvar_Click(object sender, EventArgs e)
        {

            string recuperaArquivo = BuscaArquivoTxt();
            if (recuperaArquivo != null)
            {
                AtualizaArquivoTXT(recuperaArquivo);
            }
            else
            {
                string caminhoCompleto = GerarNomeArquivoTXT();
                StreamWriter file = new StreamWriter(caminhoCompleto);
                try
                {

                    string sLine = "";
                    for (int i = 0; i <= dataGridView2.RowCount - 1; i++)
                    { //Crio um for do tamanho da quantidade de linhas existente
                        sLine = sLine + dataGridView2.Rows[i].Cells[0].Value + ","  //ncm      
                        + dataGridView2.Rows[i].Cells[2].Value + ","//aliquota de origem 
                        + dataGridView2.Rows[i].Cells[3].Value + "," //aliquota de destino  
                        + dataGridView2.Rows[i].Cells[4].Value + ","//data de atualizacao  
                        + dataGridView2.Rows[i].Cells[5].Value;//mva
                        sLine = sLine + ",";
                        file.WriteLine(sLine);
                        sLine = "";
                    }
                    file.Close();
                    MessageBox.Show("Dados exportados com sucesso.", "Program Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                catch (Exception err)
                {
                    MessageBox.Show(err.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    file.Close();
                }
            }
        }

        /// <summary>
        /// Pega a célula de aliquota de destino, editada na grid pelo contador e atualizar seu valor. 
        /// Pergunta se deseja atualizar toda a coluna de alíquota Destino para esse valor.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView2_CellEndEdit(
object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                //obtendo o valor que foi alterado na grid
                string valorAlteradoGrid = dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                if (e.ColumnIndex.Equals(3)) //se o dado alterado for na coluna 5, ao seja, aliquota destino, segue esse comportamento:
                {
                    aliD = valorAlteradoGrid;
                }
                else if (e.ColumnIndex.Equals(5)) //se o dado alterado for na coluna 7, MVA
                {
                    decimal valor = Convert.ToDecimal(valorAlteradoGrid);
                    if (valor != MVA || valor.Equals(0))
                    {
                        //produtos com mesmo ncm E ORIGEM, seta de uma vez o MESMO MVA
                        MVA = valor;
                        //pega a data do sistema no momento em que o MVA foi alterado
                        dataAtualizacaoMVA = DateTime.Today;
                        dataAtualMVAFormatada = dataAtualizacaoMVA.ToString("dd/MM/yyyy");
                        dataGridView2.Rows[e.RowIndex].Cells[4].Value = dataAtualMVAFormatada;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Não foi possível atualizar o valor. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        private void btnConfirmaSaida_Click(object sender, EventArgs e)
        {
            DialogResult escolha = MessageBox.Show("Tem certeza que deseja sair?", "Mensagem do Sistema", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (escolha == DialogResult.Yes)
            {
                Application.Exit();
            }
        }
        #region Tarefa 3 - Terceira Grid
        private void btnExportar_Click(object sender, EventArgs e)
        {
            SaveFileDialog salvar = new SaveFileDialog();
            // Aplicação Excel
            Excel.Application App;
            Excel.Workbook WorkBook;
            Excel.Worksheet WorkSheet;
            object misValor = System.Reflection.Missing.Value;

            App = new Excel.Application();
            WorkBook = App.Workbooks.Add(misValor);
            WorkSheet = (Excel.Worksheet)WorkBook.Worksheets.get_Item(1);
            int i = 0;
            int j = 0;

            // Passa as celulas do DataGridView para a Pasta do Excel
            for (i = 0; i <= dataGridView3.RowCount - 1; i++)
            {
                for (j = 0; j <= dataGridView3.ColumnCount - 1; j++)
                {
                    DataGridViewCell cell = dataGridView3[j, i];
                    WorkSheet.Cells[i + 1, j + 1] = cell.Value;
                }
            }

            //Gera um nome para novo arquivo
            Random numAleatorio = new Random();
            int valorInteiro = numAleatorio.Next();
            DateTime dataHoje = DateTime.Today;
            string data = dataHoje.ToString("D");
            string nomeArquivo = "Extrato-" + data + " " + valorInteiro + ".xls";

            string[] pasta = { value, nomeArquivo };
            string caminhoCompleto = Path.Combine(pasta);

            // define algumas propriedades da caixa salvar
            salvar.Title = "Exportar para Excel";
            salvar.Filter = "Arquivo do Excel *.xls | *.xls";
            salvar.ShowDialog(); // mostra
            
            //salva na pasta virtual escolhida 
            WorkBook.SaveAs(caminhoCompleto, Excel.XlFileFormat.xlWorkbookNormal, misValor, misValor, misValor, misValor,

            Excel.XlSaveAsAccessMode.xlExclusive, misValor, misValor, misValor, misValor, misValor);
            WorkBook.Close(true, misValor, misValor);
            App.Quit();
        }
        private Boolean verificaTemMVA(string ncm, string aliO)
        {
            Boolean resultado = true;
           
            for (int i = 0; i <= dataGridView2.RowCount - 1; i++)
            {
                string mva = dataGridView2.Rows[i].Cells[5].Value.ToString();
                string ncmGrid = dataGridView2.Rows[i].Cells[0].Value.ToString();
                string aliOrigemGrid = dataGridView2.Rows[i].Cells[2].Value.ToString();
                if (ncm.Equals(ncmGrid) && aliO.Equals(aliOrigemGrid))
                {
                    if(String.IsNullOrWhiteSpace(mva) || Convert.ToDecimal(mva) == 0)
                 
                    resultado = false;
                }  
            }
           
            return resultado;
        }
        private Boolean verificaMVAGrid()
        {
            Boolean resultado = false;
            string mva;
            for (int i = 0; i <= dataGridView2.RowCount - 1; i++)
            {
                mva = dataGridView2.Rows[i].Cells[5].Value.ToString();
             
                if (Convert.ToDecimal (mva)>0)
                {
                    resultado = true;
                    break;
                }
            }
          
            return resultado;
        }
        private string obterMVAGrid(string ncm, string aliO)
        {
            string mva=null;
            for (int i=0; i <= dataGridView2.RowCount-1; i++)
            {
                string ncmGrid = dataGridView2.Rows[i].Cells[0].Value.ToString();
                string aliOrigemGrid = dataGridView2.Rows[i].Cells[2].Value.ToString();
                if (ncm.Equals(ncmGrid) && aliO.Equals(aliOrigemGrid))
                {
                    mva = dataGridView2.Rows[i].Cells[5].Value.ToString();
                }      
            }
            return mva;
        }
        private Tuple<ExtratoImposto, int,Boolean> GerandoExtratoSemMVA(string file)
        {
            ExtratoImposto extrato;
            CalculaIcmsAntecipado icmsAntecipado;
            int linha = 0;
            int pos = 0;
            int cont = 0;
            Boolean temMva=false;
            string recuperapIPI = null;
            decimal somaProdutoComMVA = 0;
            int quantidadeProdutosNotaF = ContaItensNotaFiscal(file);
            for (int i = 0; i <= quantidadeProdutosNotaF - 1; i++)
            {
                buscaAliquotaOrigem = RetornapAliquotaOrigem(file, linha);
                if (String.IsNullOrWhiteSpace(buscaAliquotaOrigem))
                {
                    buscaAliquotaOrigem = "12";
                }
                Imposto imposto = ImpostoNotaFiscal(linha, DesserializarNota(file));
                string ncmNotaFiscal = imposto.NCM;  
                icmsAntecipado = new CalculaIcmsAntecipado();
                Boolean verificarpIPI = verificapIPI(file, linha);//tem que ser pela quantidade de produtos na nota fiscal
                if (verificarpIPI.Equals(true))
                {
                    recuperapIPI = RetornapIPI(file, linha, pos);
                    pos++;
                }
                else
                {
                    recuperapIPI = null;
                }
               
                temMva = verificaTemMVA(ncmNotaFiscal, buscaAliquotaOrigem);
               
                if (temMva.Equals(false)) 
                {
                    var tupla = DadosSegundaGridParaCalculoIcms(linha, file, recuperapIPI);
                    string valorICMSOrigem = tupla.Item1;
                    string valorProduto = tupla.Item2;
                    string Ipi = tupla.Item3;
                    precoGoverno = icmsAntecipado.CalculaPrecoGoverno(0, Ipi, Convert.ToDecimal(valorProduto));
                    somaProdutoComMVA = somaProdutoComMVA + icmsAntecipado.CalculaICMSAntecipado(precoGoverno, Convert.ToDecimal(valorICMSOrigem));
                }
                else
                {
                    cont++;     
                }

                linha++;
            }
            if (somaProdutoComMVA != 0)
            {
                extrato = ExtratoGrid(DesserializarNota(file), somaProdutoComMVA);
            }
            else
            {
                extrato = null;
            }
            return new Tuple<ExtratoImposto, int, Boolean>(extrato, cont, temMva);
        }
        private Tuple<ExtratoImposto, int, Boolean> GerandoExtratoComMVA(string file)
        {
            ExtratoImposto extrato;
            CalculaIcmsAntecipado icmsAntecipado;
            Boolean temMva = false;
            int linha = 0;
            int pos = 0;
            int cont = 0;
            string recuperapIPI = null;
            decimal somaProdutoComMVA = 0;
            int quantidadeProdutosNotaF = ContaItensNotaFiscal(file);
            for (int i = 0; i <= quantidadeProdutosNotaF - 1; i++)
            {
                buscaAliquotaOrigem = RetornapAliquotaOrigem(file, i);
                if (String.IsNullOrWhiteSpace(buscaAliquotaOrigem))
                {
                    buscaAliquotaOrigem = "12";
                  
                }
                Imposto imposto = ImpostoNotaFiscal(i, DesserializarNota(file));
                string ncmNotaFiscal = imposto.NCM;
                temMva = verificaTemMVA(ncmNotaFiscal, buscaAliquotaOrigem);
             
                if (temMva.Equals(true))
                {
                    icmsAntecipado = new CalculaIcmsAntecipado();
                    Boolean verificarpIPI = verificapIPI(file, i);
                    if (verificarpIPI.Equals(true))
                    {
                        recuperapIPI = RetornapIPI(file, i, pos);
                        pos++;
                    }
                    else
                    {
                        recuperapIPI = null;
                    }
                    string mva=obterMVAGrid(ncmNotaFiscal, buscaAliquotaOrigem);
                    var tupla = DadosSegundaGridParaCalculoIcms(i, file,recuperapIPI);
                    string valorICMSOrigem = tupla.Item1;
                    string valorProduto = tupla.Item2;
                    string Ipi = tupla.Item3;
                    precoGoverno = icmsAntecipado.CalculaPrecoGoverno(Convert.ToDecimal(mva), Ipi, Convert.ToDecimal(valorProduto));
                    somaProdutoComMVA = somaProdutoComMVA + icmsAntecipado.CalculaICMSAntecipado(precoGoverno, Convert.ToDecimal(valorICMSOrigem));
                }
                else
                {
                    cont++;
                }
                linha++;
            }
            if (somaProdutoComMVA != 0)
            {
                extrato = ExtratoGrid(DesserializarNota(file), somaProdutoComMVA);
            }
            else
            {
                extrato = null;
            }
            return new Tuple<ExtratoImposto, int, Boolean>(extrato, cont, temMva);
        }
        private void ExtratoNotas(List<ExtratoImposto> extratoList)
        {
            string[] arquivos = Directory.GetFiles(pastaSaida, "*.xml");

            foreach (var file in arquivos)
            {
               
                    var gerarExtratoComMva = GerandoExtratoComMVA(file);
                    ExtratoImposto extratoComMva = gerarExtratoComMva.Item1;
                    int contSemMva = gerarExtratoComMva.Item2;
                    Boolean temMva = gerarExtratoComMva.Item3;
                    if (temMva.Equals(true))
                    {

                        extratoComMva.FormaRecolhimento = "Antecipação com encerramento de fase";
                        extratoList.Add(extratoComMva);
                    }

                if (contSemMva > 0)
                {
                    var gerarExtratoSemMva = GerandoExtratoSemMVA(file);
                    ExtratoImposto extratoSemMVA = gerarExtratoSemMva.Item1;
                    int conta = gerarExtratoSemMva.Item2;
                    Boolean naotemMva = gerarExtratoSemMva.Item3;
                    extratoSemMVA.FormaRecolhimento = "Complementação de alíquota";
                    extratoList.Add(extratoSemMVA);
                    if (temMva.Equals(false))
                    {
                        if (conta > 0)
                        {
                            gerarExtratoComMva = GerandoExtratoComMVA(file);
                            extratoComMva = gerarExtratoComMva.Item1;
                            extratoComMva.FormaRecolhimento = "Antecipação com encerramento de fase";
                            extratoList.Add(extratoComMva);
                        }
                    }
                }
              }
            }
        
        private void btnGerarExtrato_Click(object sender, EventArgs e)
        {
            List<ExtratoImposto> extratoList = new List<ExtratoImposto>();
            if (caminho.EndsWith("xml"))
            //apenas 1 nota fiscal foi aberta, não necessita do foreach para ler cada nota xml desserializada separadamente
            {
                Boolean temMva = false;
                Boolean algumProdutoTemMVA = verificaMVAGrid();
                if (algumProdutoTemMVA.Equals(true))
                {
                        var gerarExtratoComMva = GerandoExtratoComMVA(caminho);
                        ExtratoImposto extratoComMva = gerarExtratoComMva.Item1;
                        int contSemMva = gerarExtratoComMva.Item2;
                        temMva = gerarExtratoComMva.Item3;
                        if (temMva.Equals(true))
                        {
                        extratoComMva.FormaRecolhimento = "Antecipação com encerramento de fase";
                        extratoList.Add(extratoComMva);
                        }
                                        
                        if (contSemMva > 0)
                        {
                            var gerarExtratoSemMva = GerandoExtratoSemMVA(caminho);
                            ExtratoImposto extratoSemMVA = gerarExtratoSemMva.Item1;

                            extratoSemMVA.FormaRecolhimento = "Complementação de alíquota";
                            extratoList.Add(extratoSemMVA);
                        }
                    }
                else
                {
                    var gerarExtratoSemMva = GerandoExtratoSemMVA(caminho);
                    ExtratoImposto extratoSemMVA = gerarExtratoSemMva.Item1;
                    Boolean naotemMva = gerarExtratoSemMva.Item3;
                    if (naotemMva.Equals(false))
                    {
                        extratoSemMVA.FormaRecolhimento = "Complementação de alíquota";
                        extratoList.Add(extratoSemMVA);
                    }
                    int conta = gerarExtratoSemMva.Item2;
                    if (temMva.Equals(false) && conta > 0)
                    {
                        var gerarExtratoComMva = GerandoExtratoComMVA(caminho);
                        ExtratoImposto extratoComMva = gerarExtratoComMva.Item1;
                        extratoComMva.FormaRecolhimento = "Antecipação com encerramento de fase";
                        extratoList.Add(extratoComMva);
                    }
                }
                  }   
            else
            {
                ExtratoNotas(extratoList);
            }

            this.extratoImpostoBindingSource.DataSource = extratoList;

            this.dataGridView3.DataSource =
                this.extratoImpostoBindingSource.DataSource;
        }
       
        public Tuple<string, string,string> DadosSegundaGridParaCalculoIcms(int pos, string caminho, string pIPI)
        {
            try
            {
                valorProdutoUnitario = RetornaValorProdutoUnitario(caminho, pos).Replace(".", ",");
                valorICMSOrigem = RetornavalorICMS(caminho, pos);
                return new Tuple<string, string,string>(valorICMSOrigem, valorProdutoUnitario,pIPI);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Não foi possível recuperar os dados da segunda grid para o extrato. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

        }
        public ExtratoImposto ExtratoGrid(TNfeProc nfe, decimal soma)
        {
            try
            {
                ExtratoImposto extrato = new ExtratoImposto();
                extrato.NumeroNota = nfe.NFe.infNFe.ide.nNF;
                string valorTotalNota = nfe.NFe.infNFe.total.ICMSTot.vNF;
                if (valorTotalNota != null)
                {
                    string formatvalorTotalNota = valorTotalNota.Replace(".", ",");
                    extrato.ValorTotalNota = Convert.ToDecimal(formatvalorTotalNota);
                }
                extrato.ValorICMSCalculado = Convert.ToDecimal(nfe.NFe.infNFe.total.ICMSTot.vICMS.Replace(".", ","));
                extrato.ValorAnalisado = soma;

                return extrato;
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Não foi possível criar o extrato. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public string RetornaValorProdutoUnitario(string pasta, int pos)
        {
            try
            {
                
                XmlDocument doc = new XmlDocument();
                doc.Load(pasta);
                int quantProdutosNotaFiscal = ContaItensNotaFiscal(pasta);
                XmlNodeList elemList = doc.GetElementsByTagName("prod");
             
                var recuperaItem = elemList.Item(pos);
                string format = "";
                foreach (XmlNode node in elemList)
                {
                    XmlNodeList ali = doc.GetElementsByTagName("vProd");
                    recuperaItem = ali.Item(pos);
                }
                format = recuperaItem.OuterXml;
                format = format.Replace("<vProd xmlns=\"http://www.portalfiscal.inf.br/nfe\">", "");
                format = format.Replace("</vProd>", "");
                return format;

            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Não foi possível retornar o valor do produto unitário. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public Boolean verificapIPI(string pasta, int pos)
        {
            XmlDocument xml = new XmlDocument();
            xml.Load(pasta);
            XmlNodeList detList = xml.GetElementsByTagName("det");
            Boolean busca = false;
            foreach (XmlNode det in detList[pos])
            {
                foreach (XmlNode n in det)
                {
                    if (n.Name == "IPI")
                    {
                        busca = true;
                    }
                }
            }
            return busca;
        }

        public string RetornapIPI(string pasta, int pos, int posProduto)
        {
            try
            {
                XmlDocument xml = new XmlDocument();
                xml.Load(pasta);
                XmlNodeList detList = xml.GetElementsByTagName("det");
                XmlNodeList ipi = null;
                var recuperaItem = " ";
                var format = "";
                foreach (XmlNode det in detList[pos])
                {

                    foreach (XmlNode n in det)
                    {
                        if (n.Name == "IPI")
                        {
                            ipi = xml.GetElementsByTagName("IPI");
                            foreach (XmlNode busca in ipi)
                            {
                                XmlNodeList ipitrib = xml.GetElementsByTagName("IPITrib");
                                for (int i = 0; i < ipitrib.Count; i++)
                                {
                                    if (i == posProduto)
                                    {
                                        recuperaItem = ipitrib[i]["pIPI"].InnerText;
                                    }
                                }
                            }
                        }
                    }
                }

                if (recuperaItem != null)
                {
                    format = recuperaItem;
                    format = format.Replace(".", " ");
                    format = format.TrimEnd('0', ' ');
                }
                return format;
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Não foi possível retornar o IPI. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        #endregion
    }
}
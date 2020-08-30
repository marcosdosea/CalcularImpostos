using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Xml;
using Microsoft.VisualStudio.Services.Common;
using System.ComponentModel;
using System.Windows.Documents;
using System.Linq;
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

        private string recuperaNCM;

        private DateTime dataAtualizacaoMVA;

        private decimal MVA;

        private string value = ConfiguracaoDropbox.GetValue("pastaDropbox");

        private string pIPI;

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
                NotasFiscais novaNota;
                Imposto imposto;
                List<NotasFiscais> notaList = new List<NotasFiscais>();
                List<Imposto> impostoList = new List<Imposto>();
                int quantidadeProdutosNotaF;

                foreach (var file in arquivos)
                {

                    novaNota = NovoObjeto(DesserializarNota(file));

                    notaList.Add(novaNota);

                    int pos = 0;

                    quantidadeProdutosNotaF = ContaItensNotaFiscal(file);

                    buscaAliquotaOrigem = RetornaAliquotaOrigemICMS(file, pos);

                    for (int i = 0; i <= quantidadeProdutosNotaF - 1; i++)
                    {
                        imposto = ImpostoNotaFiscal(pos, DesserializarNota(file));
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
                NotasFiscais novaNota;
                Imposto imposto;
                int quantidadeProdutosNotaF = ContaItensNotaFiscal(arquivo);
                int pos = 0;
                List<Imposto> impostoList = new List<Imposto>();

                novaNota = NovoObjeto(DesserializarNota(arquivo));

                buscaAliquotaOrigem = RetornaAliquotaOrigemICMS(arquivo, pos);

                for (int i = 0; i <= quantidadeProdutosNotaF - 1; i++)
                {
                    imposto = ImpostoNotaFiscal(pos, DesserializarNota(arquivo));
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
                    ApagarArquivo(path); //adicionei esse parametro aqui
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
                XmlNodeList elemList = doc.GetElementsByTagName("ICMS");
                var recuperaItem = elemList.Item(pos);

                foreach (XmlNode node in elemList)
                {
                    XmlNodeList ali = doc.GetElementsByTagName("pICMS");
                    recuperaItem = ali.Item(pos);
                }
                if (recuperaItem == null)
                {
                    //se o produto não tiver o pICMS
                    return null;
                }
                else
                {
                    string format = recuperaItem.OuterXml;
                    format = format.Replace("<pICMS xmlns=\"http://www.portalfiscal.inf.br/nfe\">", "");
                    format = format.Replace("</pICMS>", "");
                    format = format.TrimEnd('0', ' ');
                    format = format.Replace(".", " ");
                    return format;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Não foi possível obter aliquota de origem. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                return arquivo;

            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Não foi possível encontrar o arquivo .txt no diretório. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        public Tuple<string, string, string, string, string> LerArquivoTxt(string nomeArquivo, string ncm, string aliO, string aliD)
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
                for (int j = 0; j < teste.Length; j++)
                {
                    if (teste[j].Contains(ncm))
                    {
                        trim = teste[j].TrimEnd(',', ' ');
                        colunas = trim.Split(',');
                        for (int i = 0; i < colunas.Length; i++)
                        {
                            ncmArquivo = colunas[0].Replace(',', ' ');
                            aliquotaOrigem = colunas[1].Replace(',', ' ');
                            aliquotaDestino = colunas[2].Replace(',', ' ');
                            dataAtualizacao = colunas[3].Replace(',', ' ');
                            valorMVA = colunas[4].Replace(',', ' ');
                            proximalinha = false;
                            break;
                        }
                    }
                }

                if (proximalinha == true)
                {
                    dataAtualizacao = null;
                    valorMVA = null;
                }

                return new Tuple<string, string, string, string, string>(dataAtualizacao, valorMVA, aliquotaOrigem, aliquotaDestino, ncmArquivo);
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
                imposto.Numero = nfeProc.NFe.infNFe.ide.nNF;
                imposto.NCM = nfeProc.NFe.infNFe.det[pos].prod.NCM;
                imposto.Produto = nfeProc.NFe.infNFe.det[pos].prod.xProd; //peguei o nome
                imposto.TipoReceita = nfeProc.NFe.infNFe.ide.natOp; //natureza operação? 
                imposto.AliquotaOrigem = Convert.ToDecimal(buscaAliquotaOrigem);
                // aliD = "18" -> valor atual da alíquota de destino de sergipe;
                imposto.AliquotaDestino = Convert.ToDecimal(aliD);

                if (recuperaArquivo != null)
                {
                    var tupla = LerArquivoTxt(recuperaArquivo, imposto.NCM, buscaAliquotaOrigem, aliD);
                    string mva = tupla.Item2;

                    if (mva != null)
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
            for (int i = 0; i <= dataGridView2.RowCount - 1; i++)
            {
                string ncm = dataGridView2.Rows[i].Cells[1].Value.ToString();
                string aliO = dataGridView2.Rows[i].Cells[4].Value.ToString();
                string aliD = dataGridView2.Rows[i].Cells[5].Value.ToString();
                string mva = dataGridView2.Rows[i].Cells[7].Value.ToString();
                string dataAtualizacao = dataGridView2.Rows[i].Cells[6].Value.ToString();
                var tupla = LerArquivoTxt(pasta, ncm, aliO, aliD);
                string ncmArquivo = tupla.Item5;
                string dataAtualizacaoArquivo = tupla.Item1;
                string mvaArquivoTxt = tupla.Item2;
                string aliqOrigem = tupla.Item3;
                string aliqDestino = tupla.Item4;

                string novaLinha = "";

                if (ncm != ncmArquivo) //se o ncm ainda não tiver gravado no txt
                {
                    if (mvaArquivoTxt != mva || aliqOrigem != aliO || aliqDestino != aliD || dataAtualizacaoArquivo != dataAtualizacao)
                    {
                        //insere parte alterada na grid, abaixo dos dados que já tem salvo no txt
                        StreamWriter s = File.AppendText(pasta);

                        novaLinha = novaLinha + ncm + ","
                    + aliO + ","
                    + aliD + ","
                    + dataAtualizacao + ","
                    + mva;
                        novaLinha = novaLinha + ",";
                        s.WriteLine(novaLinha);

                        s.Close();
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
                        sLine = sLine + dataGridView2.Rows[i].Cells[1].Value + ","  //ncm      
                        + dataGridView2.Rows[i].Cells[4].Value + ","//aliquota de origem 
                        + dataGridView2.Rows[i].Cells[5].Value + "," //aliquota de destino  
                        + dataGridView2.Rows[i].Cells[6].Value + ","//data de atualizacao  
                        + dataGridView2.Rows[i].Cells[7].Value;//mva
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
            // dataAtualizacaoMVA = DateTime.Today;
            try
            {
                //obtendo o valor que foi alterado na grid
                string valorAlteradoGrid = dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                if (e.ColumnIndex.Equals(5)) //se o dado alterado for na coluna 5, ao seja, aliquota destino, segue esse comportamento:
                {
                    aliD = valorAlteradoGrid;

                    //pergunto se deseja atualizar todos os campos de alíquota Destino com esse valor, caso sim:
                    DialogResult dialogResult = MessageBox.Show("Deseja atualizar todos os campos de alíquota destino para esse valor?", "Atenção", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        // Limpa as células selecionadas na grid
                        dataGridView2.ClearSelection();

                        // Faz toda coluna not sortable.
                        for (int i = 0; i < dataGridView2.Columns.Count; i++)
                        {
                            dataGridView2.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                        }

                        // Seta o modo selection para Column.
                        dataGridView2.SelectionMode = DataGridViewSelectionMode.FullColumnSelect;

                        // Eu seleciono a coluna AliquotaDestino  
                        if (dataGridView2.Columns.Count > 0)  // Checa se eu tenho pelo menos uma coluna.
                        {
                            dataGridView2.Columns[5].Selected = true;
                        }
                        for (int i = 0; i < dataGridView2.RowCount; i++) //Crio um for do tamanho da quantidade de linhas existente
                        {
                            dataGridView2.Rows[i].Cells[5].Value = Convert.ToDecimal(aliD); //agora basta inserir o valor da alíquota de destino em todas as células dessa coluna
                        }

                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        MessageBox.Show(String.Format("Valor alterado com sucesso! Novo valor de aliquota destino: {0}", aliD));
                    }
                }
                else if (e.ColumnIndex.Equals(7)) //se o dado alterado for na coluna 7, MVA
                {
                    decimal valor = Convert.ToDecimal(valorAlteradoGrid);
                    if (valor != MVA)
                    {
                        //produtos com mesmo ncm E ORIGEM, seta de uma vez o MESMO MVA
                        MVA = valor;
                        //pega a data do sistema no momento em que o MVA foi alterado
                        dataAtualizacaoMVA = DateTime.Today;
                        dataAtualMVAFormatada = dataAtualizacaoMVA.ToString("dd/MM/yyyy"); //não está inserindo a data s2
                        dataGridView2.Rows[e.RowIndex].Cells[6].Value = dataAtualMVAFormatada;

                        DialogResult dialogResult = MessageBox.Show("Deseja atualizar todos os campos de MVA, com mesmo NCM e origem, para esse valor?", "Atenção", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                        {
                            // Limpa as células selecionadas na grid
                            dataGridView2.ClearSelection();

                            // Faz toda coluna not sortable.
                            for (int i = 0; i < dataGridView2.Columns.Count; i++)
                            {
                                dataGridView2.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                            }
                            // Seta o modo selection para Column.
                            dataGridView2.SelectionMode = DataGridViewSelectionMode.FullColumnSelect;

                            // Eu seleciono a coluna mva  
                            if (dataGridView2.Columns.Count > 0)
                            { // Checa se eu tenho pelo menos uma coluna.
                                dataGridView2.Columns[7].Selected = true;
                                //   dataGridView2.Columns[1].Selected = true; //NCM
                                //   dataGridView2.Columns[4].Selected = true; //ORIGEM
                            }
                            for (int i = 0; i < dataGridView2.RowCount; i++)
                            { //Crio um for do tamanho da quantidade de linhas existente
                                //CHECAR PRIMEIRO SE O NCM E ORIGEM da linha alterada é igual
                                int linha = e.RowIndex;  //PEGAR O INDEX DA LINHA ALTERADA 
                                string ncm = dataGridView2.Rows[linha].Cells[1].Value.ToString();//ncm DA LINHA ALTERADA 
                                string origem = dataGridView2.Rows[linha].Cells[4].Value.ToString();//aliquota de origem DA LINHA ALTERADA 
                                string destino = dataGridView2.Rows[linha].Cells[5].Value.ToString();
                                string ncmResto = dataGridView2.Rows[i].Cells[1].Value.ToString(); //ncm do resto  
                                string origemResto = dataGridView2.Rows[i].Cells[4].Value.ToString();//aliquota do resto 
                                string destinoResto = dataGridView2.Rows[i].Cells[5].Value.ToString();
                                if (ncmResto.Equals(ncm) && (origemResto.Equals(origem)) && (destinoResto.Equals(destino)))
                                {
                                    dataGridView2.Rows[i].Cells[7].Value = MVA; //inserir o valor mva em todas as células dessa coluna
                                    dataGridView2.Rows[i].Cells[6].Value = dataAtualMVAFormatada;
                                }
                            }
                        }
                        else if (dialogResult == DialogResult.No)
                        {
                            MessageBox.Show(String.Format("Valor alterado com sucesso! Novo valor de MVA: {0}", MVA));
                        }
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

            //salva na pasta virtual escolhida 
            WorkBook.SaveAs(caminhoCompleto, Excel.XlFileFormat.xlWorkbookNormal, misValor, misValor, misValor, misValor,

            Excel.XlSaveAsAccessMode.xlExclusive, misValor, misValor, misValor, misValor, misValor);
            WorkBook.Close(true, misValor, misValor);
            App.Quit();
        }
        private Boolean verificaMVA(int pos, string file, string pipi)
        {
            Boolean resultado = true;
            var tupla = DadosSegundaGridParaCalculoIcms(pos, file, pipi);
            string mva = tupla.Item3;
            if (mva.IsNullOrEmpty() || Convert.ToDecimal(mva) <= 0)
            {
                resultado = false;
            }
            return resultado;
        }
        private Tuple<ExtratoImposto, int> GerandoExtratoSemMVA(string file)
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
                icmsAntecipado = new CalculaIcmsAntecipado();
                Boolean verificarpIPI = verificapIPI(file, linha);
                if (verificarpIPI.Equals(true))
                {
                    recuperapIPI = RetornapIPI(file, linha, pos);
                    pos++;
                }
                else
                {
                    recuperapIPI = null;
                }
                temMva = verificaMVA(linha, file, recuperapIPI);
                if (temMva.Equals(false))
                {
                    var tupla = DadosSegundaGridParaCalculoIcms(linha, file, recuperapIPI);
                    string pIPI = tupla.Item1;
                    string valorICMSOrigem = tupla.Item2;
                    string mva = tupla.Item3;
                    string valorProduto = tupla.Item4;
                    precoGoverno = icmsAntecipado.CalculaPrecoGoverno(Convert.ToDecimal(mva), pIPI, Convert.ToDecimal(valorProduto));
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
            return new Tuple<ExtratoImposto, int>(extrato, cont);
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
                icmsAntecipado = new CalculaIcmsAntecipado();
                Boolean verificarpIPI = verificapIPI(file, linha);
                if (verificarpIPI.Equals(true))
                {
                    recuperapIPI = RetornapIPI(file, linha, pos);
                    pos++;
                }
                else
                {
                    recuperapIPI = null;
                }
                temMva = verificaMVA(linha, file, recuperapIPI);
                if (temMva.Equals(true))
                {
                    var tupla = DadosSegundaGridParaCalculoIcms(linha, file, recuperapIPI);
                    string pIPI = tupla.Item1;
                    string valorICMSOrigem = tupla.Item2;
                    string mva = tupla.Item3;
                    string valorProduto = tupla.Item4;
                    temMva = true;
                    precoGoverno = icmsAntecipado.CalculaPrecoGoverno(Convert.ToDecimal(mva), pIPI, Convert.ToDecimal(valorProduto));
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
        private void btnGerarExtrato_Click(object sender, EventArgs e)
        {
            List<ExtratoImposto> extratoList = new List<ExtratoImposto>();

            if (caminho.EndsWith("xml"))
            //apenas 1 nota fiscal foi aberta, não necessita do foreach para ler cada nota xml desserializada separadamente
            {
                var tupla = GerandoExtratoComMVA(caminho);
                ExtratoImposto extrato = tupla.Item1;
                int cont = tupla.Item2;
                Boolean temMva = tupla.Item3;
                if (temMva.Equals(true))
                {
                    extratoList.Add(extrato);
                }
                if (cont > 0)
                {
                    var gerarExtratoSemMva = GerandoExtratoSemMVA(caminho);
                    ExtratoImposto extratoSemMVA = gerarExtratoSemMva.Item1;
                    extratoList.Add(extratoSemMVA);
                    int conta = gerarExtratoSemMva.Item2;
                    if (conta > 0)
                    {
                        tupla = GerandoExtratoComMVA(caminho);
                        extrato = tupla.Item1;
                        extratoList.Add(extrato);
                    }
                }
            }
            else
            {
                string[] arquivos = Directory.GetFiles(pastaSaida, "*.xml");
                foreach (var file in arquivos)
                {
                    var tupla = GerandoExtratoComMVA(file);
                    ExtratoImposto extrato = tupla.Item1;
                    int cont = tupla.Item2;
                    Boolean temMva = tupla.Item3;
                    if (temMva.Equals(true))
                    {
                        extratoList.Add(extrato);
                    }
                    if (cont > 0)
                    {
                        var gerarExtratoSemMva = GerandoExtratoSemMVA(file);
                        ExtratoImposto extratoSemMVA = gerarExtratoSemMva.Item1;
                        extratoList.Add(extratoSemMVA);
                        int conta = gerarExtratoSemMva.Item2;
                        if (temMva.Equals(false))
                        {
                            if (conta > 0)
                            {
                                tupla = GerandoExtratoComMVA(file);
                                extrato = tupla.Item1;
                                extratoList.Add(extrato);
                            }
                        }
                    }
                }
            }

            this.extratoImpostoBindingSource.DataSource = extratoList;

            this.dataGridView3.DataSource =
                this.extratoImpostoBindingSource.DataSource;
        }

        public Tuple<string, string, string, string> DadosSegundaGridParaCalculoIcms(int pos, string caminho, string pIPI)
        {
            try
            {
                string mva = dataGridView2.Rows[pos].Cells[7].Value.ToString();
                valorProdutoUnitario = RetornaValorProdutoUnitario(caminho, pos).Replace(".", ",");
                valorICMSOrigem = RetornaAliquotaOrigemICMS(caminho, pos);
                return new Tuple<string, string, string, string>(pIPI, valorICMSOrigem, mva, valorProdutoUnitario);
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
            XmlNodeList elemList = xml.GetElementsByTagName("imposto");
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
                XmlNodeList elemList = xml.GetElementsByTagName("imposto");
                XmlNodeList ipi = null;
                var recuperaItem = " ";
                var format = "";
                XmlNode node = elemList[pos];
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
                                    // break;
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
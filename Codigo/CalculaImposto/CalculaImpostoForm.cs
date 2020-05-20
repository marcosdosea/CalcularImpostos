using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Windows.Forms;
using System.Xml;
using Dropbox.Api;
using Excel = Microsoft.Office.Interop.Excel;

namespace CalculaImposto
{
    public partial class FrmCalculaImposto : Form
    {
        private string caminho;

        private string pastaSaida;

        private string subdiretorio;

        private string pastaExportarGrid;

        private string dataAtualMVAFormatada;

        private string buscaAliquotaOrigem;

        private string aliD = "18";

        private DateTime dataAtualizacaoMVA;

        private decimal MVA;
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

                    buscaAliquotaOrigem = RetornaAliquotaOrigemICMS(arquivo, pos);

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

                buscaAliquotaOrigem = RetornaAliquotaOrigemICMS(caminho, pos);

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
                    ApagarArquivo(path); //adicionei esse parãmetro aqui
                }
                Directory.Delete(path, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Não foi possível excluir o diretorio. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApagarArquivo(string path) //coloquei esse parãmetro path
        {
            try
            {
             //   string[] arquivos = Directory.GetFiles(path, "*.xml");

                string[] arquivos = Directory.GetFiles(path); //não sei se apaga os arquivos dentro da pasta, se eles não forem xml
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
                    XmlNodeList ali = doc.GetElementsByTagName("orig");
                    recuperaItem = ali.Item(pos);
                }
                string format = recuperaItem.OuterXml;
                format = format.Replace("<orig xmlns=\"http://www.portalfiscal.inf.br/nfe\">", "");
                format = format.Replace("</orig>", "");
                return format;
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Não foi possível obter aliquota de origem. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                Imposto imposto = new Imposto();

                imposto.Numero = nfeProc.NFe.infNFe.ide.nNF;

                imposto.NCM = nfeProc.NFe.infNFe.det[pos].prod.NCM;
                imposto.Produto = nfeProc.NFe.infNFe.det[pos].prod.xProd; //peguei o nome
                imposto.TipoReceita = nfeProc.NFe.infNFe.ide.natOp; //natureza operação?


                imposto.AliquotaOrigem = Convert.ToDecimal(buscaAliquotaOrigem);

                // aliD = "18" -> valor atual da alíquota de destino de sergipe;

                imposto.AliquotaDestino = Convert.ToDecimal(String.Format("{0:p}", aliD));

                //se o campo da datagriv mva for alterado, altera a dataAtualizacao, caso não, não faz nada! observação importante
                imposto.MVA = MVA;
             
                imposto.DataAtualizacaoMVA = dataAtualMVAFormatada;
              
                return imposto;
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Não foi possível criar o objeto Imposto. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        /// <summary>
        /// Salva documento "*.xls" no dropbox, conta `saomarcosmateriais@gmail.com`
        /// </summary>
        /// <param name="nomeArquivo">Nome do arquivo exportado da grid com a extensão "*.xls"</param>
        /// <param name="caminhoCompleto">Caminho completo do diretório até chegar ao arquivo exportado da grid, com sua extensão "*.xls"</param>
        private void Dropbox(string nomeArquivo, string caminhoCompleto)
        {
            //o token está expirando... 
            DropboxApi.DropboxApi dropbox = new DropboxApi.DropboxApi();
            _ = dropbox.Upload(new DropboxClient("qiPNSnvudfAAAAAAAAAAFd3eYqmhfFFlV8E6SjIc2EdWtYGOJsiehCsE8VAc62jz"), "/ResumoNotasFiscais",
             nomeArquivo, caminhoCompleto);
            mensagemSucesso();
        }
        /// <summary>
        /// Exporta a segunda grid para o formato "*.xls", salva o arquivo localmente em uma pasta gerada e chama o método Dropbox.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSalvar_Click(object sender, EventArgs e)
        {
            //cria pasta na aplicacao para salvar localmente o arquivo excel
            pastaExportarGrid = "ResumoNotasFiscais";
            CriarDirectorio(pastaExportarGrid);

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
            for (i = 0; i <= dataGridView2.RowCount - 1; i++)
            {
                for (j = 0; j <= dataGridView2.ColumnCount - 1; j++)
                {
                    DataGridViewCell cell = dataGridView2[j, i];
                    WorkSheet.Cells[i + 1, j + 1] = cell.Value;
                }
            }
          
            //Gera um nome para novo arquivo
            Random numAleatorio = new Random(); 
            int valorInteiro = numAleatorio.Next();
            DateTime dataHoje = DateTime.Today;
            string data = dataHoje.ToString("D");
            string nomeArquivo = "ResumoNotasFiscais-" + data+ " "+valorInteiro+ ".xls";
            
            //Pega caminho do arquivo criado, dentro da nova pasta criada na aplicação 
            string CaminhoRaiz = System.Environment.CurrentDirectory; //Pasta debug
            string[] pasta = { CaminhoRaiz, pastaExportarGrid, nomeArquivo};
            string caminhoCompleto = Path.Combine(pasta); 

            //salva localmente
            WorkBook.SaveAs(caminhoCompleto, Excel.XlFileFormat.xlWorkbookNormal, misValor, misValor, misValor, misValor,

            Excel.XlSaveAsAccessMode.xlExclusive, misValor, misValor, misValor, misValor, misValor);
            WorkBook.Close(true, misValor, misValor);
            App.Quit(); 

            //Salva no dropbox
            Dropbox(nomeArquivo, caminhoCompleto);

        }
        
        private void mensagemSucesso()
        {
            MessageBox.Show("Arquivo salvo no dropbox");
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
                            dataGridView2.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;

                        // Seta o modo selection para Column.
                        dataGridView2.SelectionMode = DataGridViewSelectionMode.FullColumnSelect;

                        // Eu seleciono a coluna AliquotaDestino  
                        if (dataGridView2.Columns.Count > 0)  // Checa se eu tenho pelo menos uma coluna.
                            dataGridView2.Columns[5].Selected = true;
                        for (int i = 0; i < dataGridView2.RowCount; i++) //Crio um for do tamanho da quantidade de linhas existente
                            dataGridView2.Rows[i].Cells[5].Value = aliD; //agora basta inserir o valor da alíquota de destino em todas as células dessa coluna
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        MessageBox.Show(String.Format("Valor alterado com sucesso! Novo valor de aliquota destino: {0}", aliD));
                    }
                } else if (e.ColumnIndex.Equals(7)) //se o dado alterado for na coluna 7, MVA
                {
                    decimal valor = Convert.ToDecimal(valorAlteradoGrid);
                    if (valor !=MVA)
                    {
                        //produtos com mesmo ncm E ORIGEM, seta de uma vez o MESMO MVA
                        MVA = Convert.ToDecimal(valorAlteradoGrid);
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
                                dataGridView2.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;

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
                                string ncm = dataGridView2.Rows[linha].Cells[1].Value.ToString();//PEGAR O ncm DA LINHA ALTERADA 
                                string origem = dataGridView2.Rows[linha].Cells[4].Value.ToString();//PEGAR a aliquota de origem DA LINHA ALTERADA 

                                string ncmResto = dataGridView2.Rows[i].Cells[1].Value.ToString(); //PEGAR O ncm do resto  
                                string origemResto = dataGridView2.Rows[i].Cells[4].Value.ToString();//PEGAR a aliquota do resto 
                                if (ncmResto.Equals(ncm) && (origemResto.Equals(origem)))
                                {
                                    dataGridView2.Rows[i].Cells[7].Value = MVA; //agora basta inserir o valor mva em todas as células dessa coluna
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
    }
}

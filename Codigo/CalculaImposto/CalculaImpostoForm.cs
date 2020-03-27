﻿using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Windows.Forms;


namespace CalculaImposto
{
    public partial class FrmCalculaImposto : Form
    {
        private string caminho;

        string pastaSaida;

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

                        ExtrairZip(caminho);
                        DirectorioExiste(pastaSaida);

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(String.Format("Não foi possível abrir o arquivo. Excecão OpenFileDialog. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        public NotasFiscais novoObjeto(TNfeProc nfeProc)
        {
            try
            {
                NotasFiscais notasFiscais = new NotasFiscais();
                //  notasFiscais.Numero = nfeProc.NFe.infNFe.ide.cNF;
                notasFiscais.Numero = nfeProc.NFe.infNFe.ide.nNF;
                //estou pegando a inscrição estadual de quem emite a nota fiscal, é o fornecedor?
                notasFiscais.Fornecedor = nfeProc.NFe.infNFe.emit.IE;
                notasFiscais.DataEmissao = nfeProc.NFe.infNFe.ide.dhEmi;
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

        private void notasFiscaisBindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }

        #region Métodos Arquivo
        public void DirectorioExiste(string path)
        {
            if (Directory.Exists(path))
            {
                // Processa a lista de arquivos encontrado no diretorio.
                string[] fileEntries = Directory.GetFiles(path);
                foreach (string fileName in fileEntries) 
                {
                    ProcessarArquivo();
                }
            }
        }

        public void ProcessarArquivo()
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
                 
                    novaNota = novoObjeto(nfe);
                        
                    notaList.Add(novaNota);
                }
                this.notasFiscaisBindingSource.DataSource = notaList;

                this.dataGridView1.DataSource =
                   this.notasFiscaisBindingSource;
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Não foi possível processar os arquivos do diretorio. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            try {
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
            }
            catch (Exception ex) 
            {
                MessageBox.Show(String.Format("Não foi possível extrair os arquivos. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion
    }
}

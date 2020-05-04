namespace CalculaImposto
{
    partial class FrmCalculaImposto
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle10 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle11 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle12 = new System.Windows.Forms.DataGridViewCellStyle();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.numeroDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fornecedorDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataEmissaoDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ValorFrete = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.valorProdutosDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.valorTotalDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.notasFiscaisBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.btnBuscarNfe = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.btnSalvar = new System.Windows.Forms.Button();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.NumeroNotaFiscal = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.nCMDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Produto = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tipoReceitaDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.aliquotaOrigemDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.aliquotaDestinoDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DataAtualizacaoMVA = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mVADataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.impostoBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.dataGridView3 = new System.Windows.Forms.DataGridView();
            this.numeroNotaDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.valorTotalNotaDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.valorICMSCalculadoDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.valorRecolherDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.valorAnalisadoDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.formaRecolhimentoDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.extratoImpostoBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.btnBuscar = new System.Windows.Forms.Button();
            this.textBoxFile = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.openFileDialogNfe = new System.Windows.Forms.OpenFileDialog();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.notasFiscaisBindingSource)).BeginInit();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.impostoBindingSource)).BeginInit();
            this.tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.extratoImpostoBindingSource)).BeginInit();
            this.tabPage4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Location = new System.Drawing.Point(-3, 1);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(774, 433);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dataGridView1);
            this.tabPage1.Controls.Add(this.btnBuscarNfe);
            this.tabPage1.Controls.Add(this.textBox1);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(766, 407);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Notas Fiscais";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.numeroDataGridViewTextBoxColumn,
            this.fornecedorDataGridViewTextBoxColumn,
            this.dataEmissaoDataGridViewTextBoxColumn,
            this.ValorFrete,
            this.valorProdutosDataGridViewTextBoxColumn,
            this.valorTotalDataGridViewTextBoxColumn});
            this.dataGridView1.DataSource = this.notasFiscaisBindingSource;
            this.dataGridView1.Location = new System.Drawing.Point(6, 32);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(735, 317);
            this.dataGridView1.TabIndex = 6;
            // 
            // numeroDataGridViewTextBoxColumn
            // 
            this.numeroDataGridViewTextBoxColumn.DataPropertyName = "Numero";
            this.numeroDataGridViewTextBoxColumn.HeaderText = "Numero";
            this.numeroDataGridViewTextBoxColumn.Name = "numeroDataGridViewTextBoxColumn";
            this.numeroDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // fornecedorDataGridViewTextBoxColumn
            // 
            this.fornecedorDataGridViewTextBoxColumn.DataPropertyName = "Fornecedor";
            this.fornecedorDataGridViewTextBoxColumn.HeaderText = "Fornecedor";
            this.fornecedorDataGridViewTextBoxColumn.Name = "fornecedorDataGridViewTextBoxColumn";
            this.fornecedorDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // dataEmissaoDataGridViewTextBoxColumn
            // 
            this.dataEmissaoDataGridViewTextBoxColumn.DataPropertyName = "DataEmissao";
            dataGridViewCellStyle7.Format = "G";
            dataGridViewCellStyle7.NullValue = null;
            this.dataEmissaoDataGridViewTextBoxColumn.DefaultCellStyle = dataGridViewCellStyle7;
            this.dataEmissaoDataGridViewTextBoxColumn.HeaderText = "DataEmissao";
            this.dataEmissaoDataGridViewTextBoxColumn.Name = "dataEmissaoDataGridViewTextBoxColumn";
            this.dataEmissaoDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // ValorFrete
            // 
            this.ValorFrete.DataPropertyName = "ValorFrete";
            dataGridViewCellStyle8.NullValue = null;
            this.ValorFrete.DefaultCellStyle = dataGridViewCellStyle8;
            this.ValorFrete.HeaderText = "ValorFrete";
            this.ValorFrete.Name = "ValorFrete";
            this.ValorFrete.ReadOnly = true;
            // 
            // valorProdutosDataGridViewTextBoxColumn
            // 
            this.valorProdutosDataGridViewTextBoxColumn.DataPropertyName = "ValorProdutos";
            dataGridViewCellStyle9.NullValue = null;
            this.valorProdutosDataGridViewTextBoxColumn.DefaultCellStyle = dataGridViewCellStyle9;
            this.valorProdutosDataGridViewTextBoxColumn.HeaderText = "ValorProdutos";
            this.valorProdutosDataGridViewTextBoxColumn.Name = "valorProdutosDataGridViewTextBoxColumn";
            this.valorProdutosDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // valorTotalDataGridViewTextBoxColumn
            // 
            this.valorTotalDataGridViewTextBoxColumn.DataPropertyName = "ValorTotal";
            dataGridViewCellStyle10.NullValue = null;
            this.valorTotalDataGridViewTextBoxColumn.DefaultCellStyle = dataGridViewCellStyle10;
            this.valorTotalDataGridViewTextBoxColumn.HeaderText = "ValorTotal";
            this.valorTotalDataGridViewTextBoxColumn.Name = "valorTotalDataGridViewTextBoxColumn";
            this.valorTotalDataGridViewTextBoxColumn.ReadOnly = true;
            // 
            // notasFiscaisBindingSource
            // 
            this.notasFiscaisBindingSource.DataSource = typeof(CalculaImposto.NotasFiscais);
            this.notasFiscaisBindingSource.CurrentChanged += new System.EventHandler(this.NotasFiscaisBindingSource_CurrentChanged);
            // 
            // btnBuscarNfe
            // 
            this.btnBuscarNfe.Location = new System.Drawing.Point(666, 3);
            this.btnBuscarNfe.Name = "btnBuscarNfe";
            this.btnBuscarNfe.Size = new System.Drawing.Size(75, 23);
            this.btnBuscarNfe.TabIndex = 5;
            this.btnBuscarNfe.Text = "Importar";
            this.btnBuscarNfe.UseVisualStyleBackColor = true;
            this.btnBuscarNfe.Click += new System.EventHandler(this.BtnBuscarNfe_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(139, 6);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(521, 20);
            this.textBox1.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(117, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Arquivos Notas Fiscais:";
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.btnSalvar);
            this.tabPage2.Controls.Add(this.dataGridView2);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(766, 407);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Imposto NCM";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // btnSalvar
            // 
            this.btnSalvar.Location = new System.Drawing.Point(673, 366);
            this.btnSalvar.Name = "btnSalvar";
            this.btnSalvar.Size = new System.Drawing.Size(88, 24);
            this.btnSalvar.TabIndex = 1;
            this.btnSalvar.Text = "Salvar";
            this.btnSalvar.UseVisualStyleBackColor = true;
            this.btnSalvar.Click += new System.EventHandler(this.btnSalvar_Click);
            // 
            // dataGridView2
            // 
            this.dataGridView2.AutoGenerateColumns = false;
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.NumeroNotaFiscal,
            this.nCMDataGridViewTextBoxColumn,
            this.Produto,
            this.tipoReceitaDataGridViewTextBoxColumn,
            this.aliquotaOrigemDataGridViewTextBoxColumn,
            this.aliquotaDestinoDataGridViewTextBoxColumn,
            this.DataAtualizacaoMVA,
            this.mVADataGridViewTextBoxColumn});
            this.dataGridView2.DataSource = this.impostoBindingSource;
            this.dataGridView2.Location = new System.Drawing.Point(11, 6);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.Size = new System.Drawing.Size(759, 354);
            this.dataGridView2.TabIndex = 0;
            // 
            // NumeroNotaFiscal
            // 
            this.NumeroNotaFiscal.DataPropertyName = "Numero";
            this.NumeroNotaFiscal.HeaderText = "Número NotaFiscal";
            this.NumeroNotaFiscal.Name = "NumeroNotaFiscal";
            // 
            // nCMDataGridViewTextBoxColumn
            // 
            this.nCMDataGridViewTextBoxColumn.DataPropertyName = "NCM";
            this.nCMDataGridViewTextBoxColumn.HeaderText = "NCM";
            this.nCMDataGridViewTextBoxColumn.Name = "nCMDataGridViewTextBoxColumn";
            // 
            // Produto
            // 
            this.Produto.DataPropertyName = "Produto";
            this.Produto.HeaderText = "Produto";
            this.Produto.Name = "Produto";
            // 
            // tipoReceitaDataGridViewTextBoxColumn
            // 
            this.tipoReceitaDataGridViewTextBoxColumn.DataPropertyName = "TipoReceita";
            this.tipoReceitaDataGridViewTextBoxColumn.HeaderText = "TipoReceita";
            this.tipoReceitaDataGridViewTextBoxColumn.Name = "tipoReceitaDataGridViewTextBoxColumn";
            // 
            // aliquotaOrigemDataGridViewTextBoxColumn
            // 
            this.aliquotaOrigemDataGridViewTextBoxColumn.DataPropertyName = "AliquotaOrigem";
            this.aliquotaOrigemDataGridViewTextBoxColumn.HeaderText = "AliquotaOrigem";
            this.aliquotaOrigemDataGridViewTextBoxColumn.Name = "aliquotaOrigemDataGridViewTextBoxColumn";
            // 
            // aliquotaDestinoDataGridViewTextBoxColumn
            // 
            this.aliquotaDestinoDataGridViewTextBoxColumn.DataPropertyName = "AliquotaDestino";
            dataGridViewCellStyle11.NullValue = null;
            this.aliquotaDestinoDataGridViewTextBoxColumn.DefaultCellStyle = dataGridViewCellStyle11;
            this.aliquotaDestinoDataGridViewTextBoxColumn.HeaderText = "AliquotaDestino";
            this.aliquotaDestinoDataGridViewTextBoxColumn.Name = "aliquotaDestinoDataGridViewTextBoxColumn";
            // 
            // DataAtualizacaoMVA
            // 
            this.DataAtualizacaoMVA.DataPropertyName = "DataAtualizacaoMVA";
            dataGridViewCellStyle12.Format = "d";
            dataGridViewCellStyle12.NullValue = null;
            this.DataAtualizacaoMVA.DefaultCellStyle = dataGridViewCellStyle12;
            this.DataAtualizacaoMVA.HeaderText = "Data AtualizacaoMVA";
            this.DataAtualizacaoMVA.Name = "DataAtualizacaoMVA";
            // 
            // mVADataGridViewTextBoxColumn
            // 
            this.mVADataGridViewTextBoxColumn.DataPropertyName = "MVA";
            this.mVADataGridViewTextBoxColumn.HeaderText = "MVA";
            this.mVADataGridViewTextBoxColumn.Name = "mVADataGridViewTextBoxColumn";
            // 
            // impostoBindingSource
            // 
            this.impostoBindingSource.DataSource = typeof(CalculaImposto.Imposto);
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.dataGridView3);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(766, 407);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Extrato";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // dataGridView3
            // 
            this.dataGridView3.AutoGenerateColumns = false;
            this.dataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView3.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.numeroNotaDataGridViewTextBoxColumn,
            this.valorTotalNotaDataGridViewTextBoxColumn,
            this.valorICMSCalculadoDataGridViewTextBoxColumn,
            this.valorRecolherDataGridViewTextBoxColumn,
            this.valorAnalisadoDataGridViewTextBoxColumn,
            this.formaRecolhimentoDataGridViewTextBoxColumn});
            this.dataGridView3.DataSource = this.extratoImpostoBindingSource;
            this.dataGridView3.Location = new System.Drawing.Point(3, 6);
            this.dataGridView3.Name = "dataGridView3";
            this.dataGridView3.Size = new System.Drawing.Size(767, 354);
            this.dataGridView3.TabIndex = 0;
            // 
            // numeroNotaDataGridViewTextBoxColumn
            // 
            this.numeroNotaDataGridViewTextBoxColumn.DataPropertyName = "NumeroNota";
            this.numeroNotaDataGridViewTextBoxColumn.HeaderText = "NumeroNota";
            this.numeroNotaDataGridViewTextBoxColumn.Name = "numeroNotaDataGridViewTextBoxColumn";
            // 
            // valorTotalNotaDataGridViewTextBoxColumn
            // 
            this.valorTotalNotaDataGridViewTextBoxColumn.DataPropertyName = "ValorTotalNota";
            this.valorTotalNotaDataGridViewTextBoxColumn.HeaderText = "ValorTotalNota";
            this.valorTotalNotaDataGridViewTextBoxColumn.Name = "valorTotalNotaDataGridViewTextBoxColumn";
            // 
            // valorICMSCalculadoDataGridViewTextBoxColumn
            // 
            this.valorICMSCalculadoDataGridViewTextBoxColumn.DataPropertyName = "ValorICMSCalculado";
            this.valorICMSCalculadoDataGridViewTextBoxColumn.HeaderText = "ValorICMSCalculado";
            this.valorICMSCalculadoDataGridViewTextBoxColumn.Name = "valorICMSCalculadoDataGridViewTextBoxColumn";
            // 
            // valorRecolherDataGridViewTextBoxColumn
            // 
            this.valorRecolherDataGridViewTextBoxColumn.DataPropertyName = "ValorRecolher";
            this.valorRecolherDataGridViewTextBoxColumn.HeaderText = "ValorRecolher";
            this.valorRecolherDataGridViewTextBoxColumn.Name = "valorRecolherDataGridViewTextBoxColumn";
            // 
            // valorAnalisadoDataGridViewTextBoxColumn
            // 
            this.valorAnalisadoDataGridViewTextBoxColumn.DataPropertyName = "ValorAnalisado";
            this.valorAnalisadoDataGridViewTextBoxColumn.HeaderText = "ValorAnalisado";
            this.valorAnalisadoDataGridViewTextBoxColumn.Name = "valorAnalisadoDataGridViewTextBoxColumn";
            // 
            // formaRecolhimentoDataGridViewTextBoxColumn
            // 
            this.formaRecolhimentoDataGridViewTextBoxColumn.DataPropertyName = "FormaRecolhimento";
            this.formaRecolhimentoDataGridViewTextBoxColumn.HeaderText = "FormaRecolhimento";
            this.formaRecolhimentoDataGridViewTextBoxColumn.Name = "formaRecolhimentoDataGridViewTextBoxColumn";
            // 
            // extratoImpostoBindingSource
            // 
            this.extratoImpostoBindingSource.DataSource = typeof(CalculaImposto.ExtratoImposto);
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.btnBuscar);
            this.tabPage4.Controls.Add(this.textBoxFile);
            this.tabPage4.Controls.Add(this.label1);
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage4.Size = new System.Drawing.Size(766, 407);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "Configurações";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // btnBuscar
            // 
            this.btnBuscar.Location = new System.Drawing.Point(669, 6);
            this.btnBuscar.Name = "btnBuscar";
            this.btnBuscar.Size = new System.Drawing.Size(75, 23);
            this.btnBuscar.TabIndex = 2;
            this.btnBuscar.Text = "Buscar";
            this.btnBuscar.UseVisualStyleBackColor = true;
            this.btnBuscar.Click += new System.EventHandler(this.BtnBuscar_Click);
            // 
            // textBoxFile
            // 
            this.textBoxFile.Location = new System.Drawing.Point(154, 9);
            this.textBoxFile.Name = "textBoxFile";
            this.textBoxFile.ReadOnly = true;
            this.textBoxFile.Size = new System.Drawing.Size(509, 20);
            this.textBoxFile.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(142, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Pasta Drive Virtual Impostos:";
            // 
            // openFileDialogNfe
            // 
            this.openFileDialogNfe.FileName = "openFileDialog";
            // 
            // FrmCalculaImposto
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(774, 438);
            this.Controls.Add(this.tabControl1);
            this.Name = "FrmCalculaImposto";
            this.Text = "Cálculo de Impostos";
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.notasFiscaisBindingSource)).EndInit();
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.impostoBindingSource)).EndInit();
            this.tabPage3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.extratoImpostoBindingSource)).EndInit();
            this.tabPage4.ResumeLayout(false);
            this.tabPage4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.Button btnBuscar;
        private System.Windows.Forms.TextBox textBoxFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog openFileDialogNfe;
        private System.Windows.Forms.Button btnBuscarNfe;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.BindingSource notasFiscaisBindingSource;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.BindingSource impostoBindingSource;
        private System.Windows.Forms.DataGridView dataGridView3;
        private System.Windows.Forms.DataGridViewTextBoxColumn numeroNotaDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn valorTotalNotaDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn valorICMSCalculadoDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn valorRecolherDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn valorAnalisadoDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn formaRecolhimentoDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn diferencaDataGridViewTextBoxColumn;
        private System.Windows.Forms.BindingSource extratoImpostoBindingSource;
        private System.Windows.Forms.DataGridViewTextBoxColumn numeroDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn fornecedorDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataEmissaoDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn ValorFrete;
        private System.Windows.Forms.DataGridViewTextBoxColumn valorProdutosDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn valorTotalDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn NumeroNotaFiscal;
        private System.Windows.Forms.DataGridViewTextBoxColumn nCMDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn Produto;
        private System.Windows.Forms.DataGridViewTextBoxColumn tipoReceitaDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn aliquotaOrigemDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn aliquotaDestinoDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn DataAtualizacaoMVA;
        private System.Windows.Forms.DataGridViewTextBoxColumn mVADataGridViewTextBoxColumn;
        private System.Windows.Forms.Button btnSalvar;
    }
}


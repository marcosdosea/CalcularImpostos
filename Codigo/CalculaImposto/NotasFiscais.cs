﻿namespace CalculaImposto
{
    public class NotasFiscais
    {
        public string Numero { get; set;  }
        public string NomeFornecedor { get; set; }
        public string CnpjFornecedor { get; set; }
        public string DataEmissao { get; set; }
        public decimal ValorProdutos { get; set; }
        public decimal ValorFrete { get; set; }
        public decimal ValorTotal { get; set; }
    }
}

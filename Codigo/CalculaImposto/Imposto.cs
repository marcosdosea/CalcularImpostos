﻿namespace CalculaImposto
{
    public class Imposto
    {
        public string NCM { get; set; }
        public string TipoReceita { get; set; }
        public decimal AliquotaOrigem { get; set; }
        public decimal AliquotaDestino { get; set; }
        public decimal MVA { get; set; }
        public string DataAtualizacaoMVA { get; set; }
    }
}

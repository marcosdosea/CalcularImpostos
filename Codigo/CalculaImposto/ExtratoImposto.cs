namespace CalculaImposto
{
    public class ExtratoImposto
    {
        public string NumeroNota { get; set; }
        public decimal ValorTotalNota { get; set; }
        public decimal ValorICMSCalculado { get; set; }
        public decimal ValorRecolher { get; set; }
        public decimal ValorAnalisado { get; set; }
        public string FormaRecolhimento { get; set; }
        public decimal Diferenca { get; set; }
    }
}

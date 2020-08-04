using Microsoft.TeamFoundation.Common;
using System;

namespace CalculaImposto
{
    public class CalculaIcmsAntecipado
    {
        private decimal ICMSAntecipado;
        public decimal CalculaPrecoGoverno(decimal mva, string pIPI, decimal valorCompra)
        {
            decimal PrecoGoverno;
           
            //ver se o NCM possui MVA >0
            if (mva > 0)
            {
                //se mva > 0 então ele é substituto
                if (pIPI.IsNullOrEmpty()) //o produto não tem pIPI
                {
                    PrecoGoverno = valorCompra * (mva / 100);
                    
                }
                else
                {
                    //o produto tem pIPI
                    PrecoGoverno = ((valorCompra + Convert.ToDecimal(pIPI)) * (mva / 100)) + Convert.ToDecimal(pIPI);
                  
                }
            }
            else //if (mva == 0)//se o mva for igual a zero é porque o produto é normal
            {
                if (pIPI.IsNullOrEmpty()) //o produto não tem pIPI
                {
                    PrecoGoverno = valorCompra;
                   
                }
                else //o produto tem pIPI
                {
                    PrecoGoverno = valorCompra + Convert.ToDecimal(pIPI);
                  
                }
            }
            //truncar duas casas decimais após a virgula do preço do governo
            return PrecoGoverno;
        }
        public decimal CalculaICMSAntecipado(decimal precoGoverno, decimal valorICMCompra)
        {
            //truncar duas casas decimais após a vírgula do ICMSAntecipado
            ICMSAntecipado = ((precoGoverno * (17 / 100)) - valorICMCompra);
            
            return ICMSAntecipado;
        }
    }
}
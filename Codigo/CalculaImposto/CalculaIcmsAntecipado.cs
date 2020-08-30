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
            if (mva > 0) //adiciona uma linha com os produtos que tem mva
            {
                //se mva > 0 então ele é substituto
                if (pIPI.IsNullOrEmpty()) //o produto não tem pIPI
                {
                    PrecoGoverno = valorCompra * (mva / 100);
                    
                }
                else //adiciona uma linha para os produtos que NÃo tem mva
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
            return PrecoGoverno;
        }
        public decimal CalculaICMSAntecipado(decimal precoGoverno, decimal ICMCompra)
        {
            ICMSAntecipado = precoGoverno * ((ICMCompra/100)-(17 / 100));
            
            return Math.Round(ICMSAntecipado,2);
        }
    }
}
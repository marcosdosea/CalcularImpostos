using Microsoft.TeamFoundation.Common;
using System;
using System.Windows;

namespace CalculaImposto
{
    public class CalculaIcmsAntecipado
    {
        private decimal ICMSAntecipado;
        public decimal CalculaPrecoGoverno(decimal mva, string pIPI, decimal valorCompra)
        {
            decimal PrecoGoverno;
           
            if (mva > 0) //adiciona uma linha com os produtos que tem mva
            {
                //se mva > 0 então ele é substituto
                if (pIPI.IsNullOrEmpty()) //o produto não tem pIPI
                {
                   
                    PrecoGoverno = valorCompra * (mva / 100);
                  //  MessageBox.Show("Produto sem IPI e com MVA = " + PrecoGoverno);

                }
                else 
                {
                    
                    PrecoGoverno = ((valorCompra + Convert.ToDecimal(pIPI))* (mva / 100)) + Convert.ToDecimal(pIPI);
                 //   MessageBox.Show("Produto com IPI e com MVA = " + PrecoGoverno);

                }
            }
            else //se o mva for igual a zero,nulo ou -1 é porque o produto é normal
            {
                if (pIPI.IsNullOrEmpty()) //o produto não tem pIPI
                {
                   
                    PrecoGoverno = valorCompra;
                  ///  MessageBox.Show("Produto sem IPI e sem MVA= "+ PrecoGoverno);
                }
                else //o produto tem pIPI
                {
                    PrecoGoverno = valorCompra + Convert.ToDecimal(pIPI);
                 //   MessageBox.Show("valorCompra= " + valorCompra);
                 //   MessageBox.Show("IPI= " + pIPI);
                 //   MessageBox.Show("Produto com IPI e sem MVA = "+PrecoGoverno);
                }
            }     
            return PrecoGoverno;
        }
        public decimal CalculaICMSAntecipado(decimal precoGoverno, decimal valorICMCompra)
        {
            decimal multiplica = precoGoverno * 18/100;
            ICMSAntecipado = multiplica - valorICMCompra;
         //   MessageBox.Show("Valor valorICMCompra = " + valorICMCompra);
          //  MessageBox.Show("Valor ICMSAntecipado = " + ICMSAntecipado);
            return Math.Round(ICMSAntecipado,2);
        }
    }
}
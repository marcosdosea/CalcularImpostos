using Microsoft.TeamFoundation.Common;
using System;
using System.Windows.Forms;
using System.Xml;

namespace CalculaImposto
{
    public class CalculaIcmsAntecipado
    {
       // private decimal PrecoGoverno;
        private decimal ICMSAntecipado;

     /*   public decimal CalculaPrecoGoverno(decimal mva, string pIPI, decimal valorCompra)
        {
           // string pIPI = RetornaIPI(pasta, pos);
            //mva é a porcentagem!!!!! consertar isso s2 
            //ver se o NCM possui MVA >0
            if (mva > 0)
            {
                //se mva > 0 então ele é substituto
                if (pIPI == null) //o produto não tem pIPI
                {
                    PrecoGoverno = valorCompra * (mva / 100);
                }
                else
                {
                    //o produto tem pIPI
                    PrecoGoverno = ((valorCompra + Convert.ToDecimal(pIPI)) * (mva / 100)) + Convert.ToDecimal(pIPI);
                }
            }
            else if (mva == 0)//se o mva for igual a zero é porque o produto é normal
            {
                if (pIPI == null) //o produto não tem pIPI
                {
                    PrecoGoverno = valorCompra;
                }
                else //o produto tem pIPI
                {

                    PrecoGoverno = valorCompra + Convert.ToDecimal(pIPI);
                }
            }
            else//mva é null, não foi preenchido?
            {
               // System.Windows.MessageBox.Show("Preencha o MVA antes. MVA sem valor declarado. Caso o valor do MVA seja zero, digite 0.");
                return -1;
            }
            return PrecoGoverno;
        }
        */
        public decimal CalculaPrecoGoverno(decimal mva, string pIPI, decimal valorCompra)
        {
            decimal PrecoGoverno;
            MessageBox.Show("MVA PARA CÁLCULO: " + mva.ToString());
            MessageBox.Show("pIPI: " + pIPI);
            MessageBox.Show("valor Compra: " + valorCompra.ToString());
            // string pIPI = RetornaIPI(pasta, pos);
            //mva é a porcentagem!!!!! consertar isso s2 
            //ver se o NCM possui MVA >0
            if (mva > 0)
            {
                //se mva > 0 então ele é substituto
                if (pIPI.IsNullOrEmpty()) //o produto não tem pIPI
                {
                    PrecoGoverno = valorCompra * (mva / 100);
                    MessageBox.Show("Preço do governo sem pIPI: "+PrecoGoverno.ToString());
                }
                else
                {
                    //o produto tem pIPI
                    PrecoGoverno = ((valorCompra + Convert.ToDecimal(pIPI)) * (mva / 100)) + Convert.ToDecimal(pIPI);
                    MessageBox.Show("Preço do governo com pIPI: " + PrecoGoverno.ToString());
                }
            }
            else //if (mva == 0)//se o mva for igual a zero é porque o produto é normal
            {
                if (pIPI == null) //o produto não tem pIPI
                {
                    PrecoGoverno = valorCompra;
                    MessageBox.Show("Preço do governo: " + PrecoGoverno.ToString());
                }
                else //o produto tem pIPI
                {
                    //truncar duas casas decimais após a virgula do preço do governo!!!!!!!!!!!
                    PrecoGoverno = valorCompra + Convert.ToDecimal(pIPI);
                    MessageBox.Show("Preço do governo: " + PrecoGoverno.ToString());
                }
            }
           
            return PrecoGoverno;
        }
        public decimal CalculaICMSAntecipado(decimal precoGoverno, decimal valorICMCompra)
        {
            //truncar duas casas decimais após a vírgula do ICMSAntecipado!!!!!!!!!!!
            ICMSAntecipado = ((precoGoverno * (17 / 100)) - valorICMCompra);
            MessageBox.Show("ICMS: " + ICMSAntecipado.ToString());

            return ICMSAntecipado;
        }
    }
}
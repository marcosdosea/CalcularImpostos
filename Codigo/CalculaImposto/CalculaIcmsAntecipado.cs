using System;
using System.Windows.Forms;
using System.Xml;

namespace CalculaImposto
{
    public class CalculaIcmsAntecipado
    {
        decimal PrecoGoverno;
        decimal ICMSAntecipado;

        public string RetornaIPI(string pasta, int pos)
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(pasta);
                XmlNodeList elemList = doc.GetElementsByTagName("IPI");
                var recuperaItem = elemList.Item(pos);
                string format = "";

                foreach (XmlNode node in elemList)
                {
                    XmlNodeList ali = doc.GetElementsByTagName("pIPI");
                    recuperaItem = ali.Item(pos);
                }
                if (recuperaItem == null)
                {
                    //se o produto não tiver o pIPI
                    return null;
                }
                else
                {
                    format = recuperaItem.OuterXml;
                    format = format.Replace("<pIPI xmlns=\"http://www.portalfiscal.inf.br/nfe\">", "");
                    format = format.Replace("</pIPI>", "");
                    return format;
                }
            }
            catch (Exception ex)
            {
                //  System.Windows.MessageBox.Show(string.Format("Não foi possível obter aliquota de origem. Erro: {0}", ex.Message), "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public decimal CalculaPrecoGoverno(decimal mva, string pIPI, int pos, decimal valorCompra)
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
                    PrecoGoverno = (valorCompra + Convert.ToDecimal(pIPI)) * (mva / 100) + Convert.ToDecimal(pIPI);
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
                System.Windows.MessageBox.Show("Preencha o MVA antes. MVA sem valor declarado. Caso o valor do MVA seja zero, digite 0.");
                return -1;
            }
            return PrecoGoverno;
        }
        public decimal CalculaICMSAntecipado(decimal precoGoverno, decimal valorICMCompra)
        {
            ICMSAntecipado = precoGoverno * (17 / 100) - valorICMCompra;

            return ICMSAntecipado;
        }
    }
}
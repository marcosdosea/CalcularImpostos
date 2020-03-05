using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace CalculaImposto
{
    public class GerenciadorNfe
    {
        private static GerenciadorNfe gNFe;

        public static GerenciadorNfe GetInstance()
        {
            if (gNFe == null)
            {
                gNFe = new GerenciadorNfe();
            }

            return gNFe;
        }

        private GerenciadorNfe()
        {
        }

        public TNfeProc LerNFE(string arquivo)
        {
            XmlDocument xmldocRetorno = new XmlDocument();
            xmldocRetorno.Load(arquivo);
            XmlNodeReader xmlReaderRetorno = new XmlNodeReader(xmldocRetorno.DocumentElement);
            XmlSerializer serializer = new XmlSerializer(typeof(TNfeProc));
            TNfeProc nfe = (TNfeProc)serializer.Deserialize(xmlReaderRetorno);

            String numeroNF = nfe.NFe.infNFe.ide.nNF;
        //    String fornecedor = nfe.NFe.infNFe.ide.;
            String dataEmissao = nfe.NFe.infNFe.ide.dhEmi;
            string ValorProdutos = nfe.NFe.infNFe.ide.verProc;
            string ValorFrete = nfe.NFe.infNFe.ide.cDV;
            string ValorTotal = nfe.NFe.infNFe.ide.cDV;

            return nfe;
        }
    }
}

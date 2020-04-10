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

        public GerenciadorNfe()
        {
        }

        public TNfeProc LerNFE(string arquivo)
        {
            XmlDocument xmldocRetorno = new XmlDocument();
            xmldocRetorno.Load(arquivo);
            XmlNodeReader xmlReaderRetorno = new XmlNodeReader(xmldocRetorno.DocumentElement);
            XmlSerializer serializer = new XmlSerializer(typeof(TNfeProc));
            TNfeProc nfe = (TNfeProc)serializer.Deserialize(xmlReaderRetorno);
           
            return nfe;
        }
    }
}

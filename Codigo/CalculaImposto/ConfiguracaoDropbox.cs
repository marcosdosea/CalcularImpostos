using System;
using System.Configuration;
using System.Xml;

namespace CalculaImposto
{
    public static class ConfiguracaoDropbox
    {
        public static void UpdateAppSettings(string chave, string valor)

        {

            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

         //   string value = System.Configuration.ConfigurationManager.AppSettings[chave];
           
            config.AppSettings.Settings.Remove(chave);

            config.AppSettings.Settings.Add(chave, valor);

            config.Save(ConfigurationSaveMode.Modified);

            config.Save(ConfigurationSaveMode.Full);

            ConfigurationManager.RefreshSection("appSettings");

        }
        public static void UpdateAppConfig(string tagName, string attributeName, string value)
        {
            var doc = new XmlDocument();
            doc.Load(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
           // doc.Load(@"C:\Users\barbi\source\repos\CalcularImpostos3\Codigo\CalculaImposto\app.config");
            var tags = doc.GetElementsByTagName(tagName);
            foreach (XmlNode item in tags)
            {
                var attribute = item.Attributes[attributeName];
                if (!ReferenceEquals(null, attribute))
                    attribute.Value = value;
            }
           // doc.Save(@"C:\Users\barbi\source\repos\CalcularImpostos3\Codigo\CalculaImposto\app.config");
            doc.Save(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
        }
    }
}

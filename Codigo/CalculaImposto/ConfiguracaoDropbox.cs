using System.Configuration;

namespace CalculaImposto
{
    public static class ConfiguracaoDropbox
    {
       /* public static ExeConfigurationFileMap configFileMap = new ExeConfigurationFileMap();

        configFileMap. = "importacao.config";

        Configuration config = ConfigurationManager.OpenMappedExeConfiguration(
        configFileMap, ConfigurationUserLevel.None);*/

        static Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
        public static string GetValue(string key)
        {
            return ConfigurationManager.AppSettings[key];
        }
        public static void UpdateAppSettings(string chave, string valor)

        {
            config.AppSettings.Settings.Remove(chave);

            config.AppSettings.Settings.Add(chave, valor);

            config.Save(ConfigurationSaveMode.Modified);

            config.Save(ConfigurationSaveMode.Full);

            ConfigurationManager.RefreshSection("appSettings");

        }
    }
}

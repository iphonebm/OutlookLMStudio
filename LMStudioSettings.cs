using System;

namespace OutlookLMStudio
{
    public class LMStudioSettings
    {
        private static readonly System.Configuration.Configuration Config = 
            System.Configuration.ConfigurationManager.OpenExeConfiguration(System.Configuration.ConfigurationUserLevel.None);

        public string ApiUrl { get; set; } = "http://localhost:1234";
        public double Temperature { get; set; } = 0.7;
        public int MaxTokens { get; set; } = 2000;
        public string[] StopSequences { get; set; } = new[] { "</response>" };
        public string ModelName { get; set; } = "default";
        public int TimeoutSeconds { get; set; } = 30;

        public static LMStudioSettings LoadFromConfig()
        {
            var settings = new LMStudioSettings();
            try
            {
                var appSettings = System.Configuration.ConfigurationManager.AppSettings;
                settings.ApiUrl = appSettings["LMStudioUrl"] ?? settings.ApiUrl;
                settings.Temperature = double.Parse(appSettings["Temperature"] ?? settings.Temperature.ToString());
                settings.MaxTokens = int.Parse(appSettings["MaxTokens"] ?? settings.MaxTokens.ToString());
                settings.StopSequences = (appSettings["StopSequences"] ?? string.Join("|", settings.StopSequences)).Split('|');
                settings.ModelName = appSettings["ModelName"] ?? settings.ModelName;
                settings.TimeoutSeconds = int.Parse(appSettings["Timeout"] ?? settings.TimeoutSeconds.ToString());
            }
            catch (Exception)
            {
                // En cas d'erreur, utiliser les valeurs par défaut
            }
            return settings;
        }

        public void SaveToConfig()
        {
            var config = System.Configuration.ConfigurationManager.OpenExeConfiguration(System.Configuration.ConfigurationUserLevel.None);
            
            UpdateOrAddSetting(config, "LMStudioUrl", ApiUrl);
            UpdateOrAddSetting(config, "Temperature", Temperature.ToString());
            UpdateOrAddSetting(config, "MaxTokens", MaxTokens.ToString());
            UpdateOrAddSetting(config, "StopSequences", string.Join("|", StopSequences));
            UpdateOrAddSetting(config, "ModelName", ModelName);
            UpdateOrAddSetting(config, "Timeout", TimeoutSeconds.ToString());

            config.Save(System.Configuration.ConfigurationSaveMode.Modified);
            System.Configuration.ConfigurationManager.RefreshSection("appSettings");
        }

        private void UpdateOrAddSetting(System.Configuration.Configuration config, string key, string value)
        {
            if (config.AppSettings.Settings[key] != null)
                config.AppSettings.Settings[key].Value = value;
            else
                config.AppSettings.Settings.Add(key, value);
        }

        public string ToJson()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this);
        }

        public static LMStudioSettings FromJson(string json)
        {
            return Newtonsoft.Json.JsonConvert.DeserializeObject<LMStudioSettings>(json);
        }
    }
}
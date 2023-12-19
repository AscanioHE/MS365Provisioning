﻿using Microsoft.Extensions.Logging;
using MS365Provisioning.Common;
using Newtonsoft.Json;

namespace MS365Provisioning.Common
{
    public class ExportSettings : IExportSettings
    {
        public object DtoFile { get; set; }
        public string FilePath { get; set; }
        public string FileName { get; set; }

        private readonly ILogger _logger;

        public ExportSettings(object dto, ILogger logger, string fileName, string file)
        {
            DtoFile = dto;
            FilePath = fileName;
            FileName = file;
            _logger = logger;
        }
        public string ConvertToJsonString()
        {
            string jsonString = JsonConvert.SerializeObject(DtoFile, Formatting.Indented);
            return jsonString;
        }

        public bool WriteJsonStringToFile()
        {
            try
            {
                
                FilePath += $"{FileName}";
                string json = JsonConvert.SerializeObject(DtoFile, Formatting.Indented);
                File.WriteAllText(FilePath, json + Environment.NewLine);
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogInformation($"Error Writing Json String to file : {ex.Message}");
                return false;
            }
        }
    }
}
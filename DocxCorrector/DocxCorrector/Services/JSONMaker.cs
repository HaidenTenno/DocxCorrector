using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;
using DocxCorrector.Models;

namespace DocxCorrector.Services
{
    public static class JSONMaker
    {
        // Создать JSON строку из списка ошибок mistakes
        public static string MakeMistakesJSON(List<ParagraphResult> results)
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true
            };

            string jsonString =  JsonSerializer.Serialize(results, options);

            return jsonString;
        }
    }
}

using System;
using System.Collections.Generic;
using DocxCorrectorCore.Models;
using Newtonsoft.Json;

namespace DocxCorrectorCore.Services
{
    public static class JSONWorker
    {
        // Создать JSON строку из объекта
        public static string MakeJSON<T>(T results)
        {
            return JsonConvert.SerializeObject(results, Formatting.Indented);
        }

        // Получить из JSON строки список классов
        public static List<ParagraphClass>? DeserializeParagraphsClasses(string jsonStr)
        {
            try
            {
                List<ParagraphClass> paragraphsClasses = JsonConvert.DeserializeObject<List<ParagraphClass>>(jsonStr);
                return paragraphsClasses;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }
    }
}

using System;
using System.Collections.Generic;
using DocxCorrector.Models;
using Newtonsoft.Json;

namespace DocxCorrector.Services
{
    public static class JSONMaker
    {
        // Создать JSON строку из объекта
        public static string MakeJSON<T>(List<T> results)
        {
            return JsonConvert.SerializeObject(results, Formatting.Indented);
        }

        public static string MakeJSON<T1,T2>(Dictionary<T1,T2> results)
        {
            return JsonConvert.SerializeObject(results, Formatting.Indented);
        }
    }
}

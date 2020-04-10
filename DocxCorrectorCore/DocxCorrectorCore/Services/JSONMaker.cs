using System.Collections.Generic;
using Newtonsoft.Json;

namespace DocxCorrectorCore.Services
{
    public static class JSONMaker
    {
        // Создать JSON строку из объекта
        public static string MakeJSON<T>(List<T> results)
        {
            return JsonConvert.SerializeObject(results, Formatting.Indented);
        }

        // Создать JSON строку из объекта типа словарь
        public static string MakeJSON<T1,T2>(Dictionary<T1, T2> results) where T1 : notnull
        {
            return JsonConvert.SerializeObject(results, Formatting.Indented);
        }
    }
}

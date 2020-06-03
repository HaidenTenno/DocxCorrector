using System;
using System.IO;
using Newtonsoft.Json;

namespace DocxCorrectorCore.Services.Utilities
{
    public static class JSONWorker
    {
        // Создать JSON строку из объекта
        public static string MakeJSON<T>(T results)
        {
            return JsonConvert.SerializeObject(results, Formatting.Indented);
        }

        // Получить из JSON строки объект
        public static T? DeserializeObject<T>(string jsonStr) where T: class
        {
            try
            {
                T result = JsonConvert.DeserializeObject<T>(jsonStr);
                return result;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        public static T? DeserializeObjectFromFile<T>(string filePath) where T : class
        {
            try
            {
                using StreamReader file = File.OpenText(filePath);
                JsonSerializer serializer = new JsonSerializer();
                T? result = (T?)serializer.Deserialize(file, typeof(T));
                return result;

                //T result = JsonConvert.DeserializeObject<T>(jsonStr);
                //return result;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }
    }
}

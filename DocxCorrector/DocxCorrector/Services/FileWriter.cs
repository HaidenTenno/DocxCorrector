using System;
using System.Globalization;
using System.Collections.Generic;
using System.IO;
using CsvHelper;

namespace DocxCorrector.Services
{
    public static class FileWriter
    {
        // Записать текст text в файл, расположенный в filePath
        public static void WriteToFile(string filePath, string text)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(filePath, false, System.Text.Encoding.UTF8))
                {
                    sw.WriteLine(text);
                }
            }
            catch (Exception e)
            {
#if DEBUG
                Console.WriteLine(e.Message);
#endif
            }
        }

        // Записать свойства параграфов paragraphsInfo в CSV файл filePath
        public static void FillPropertiesCSV<T>(string filePath, List<T> listData)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(filePath))
                using (CsvWriter csv = new CsvWriter(sw, CultureInfo.InvariantCulture))
                {
                    csv.Configuration.Delimiter = ";";
                    csv.WriteRecords(listData);
                }
            }
            catch (Exception e)
            {
#if DEBUG
                Console.WriteLine(e.Message);
#endif
            }
        }
    }
}

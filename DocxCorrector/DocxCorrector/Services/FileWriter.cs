using System;
using System.Globalization;
using System.Collections.Generic;
using System.IO;
using DocxCorrector.Models;
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
                using (StreamWriter sw = new StreamWriter(filePath, false, System.Text.Encoding.Default))
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

        // Записать свойства параграфов в CSV файл
        public static void FillPropertiesCSV(string filePath, List<ParagraphProperties> paragraphsInfo)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(filePath))
                using (CsvWriter csv = new CsvWriter(sw, CultureInfo.InvariantCulture))
                {
                    csv.Configuration.Delimiter = ";";
                    csv.WriteRecords(paragraphsInfo);
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

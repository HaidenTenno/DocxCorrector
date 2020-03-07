using System;
using System.Globalization;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

        // Записать в CSV файл filePath объекты из списка listData
        public static void FillCSV<T>(string filePath, List<T> listData)
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

        // Преобразование типов
        // Заполнить CSV файл для свойств параграфов
        public static void FillCSV(string filePath, List<ParagraphProperties> listData)
        {
            List<ParagraphPropertiesInterop> listDataInterop = listData.OfType<ParagraphPropertiesInterop>().ToList();
            if (listDataInterop != null)
            {
                FillCSV(filePath: filePath, listData: listDataInterop);
                return;
            }

            List<ParagraphPropertiesGemBox> listDataGemBox= listData.OfType<ParagraphPropertiesGemBox>().ToList();
            if (listDataGemBox != null)
            {
                FillCSV(filePath: filePath, listData: listDataGemBox);
                return;
            }

            FillCSV(filePath: filePath, listData: listData);
        }

        // TODO: - Описать аналогичные перегрузки для: NormalizedProperties, PageProperties
    }
}

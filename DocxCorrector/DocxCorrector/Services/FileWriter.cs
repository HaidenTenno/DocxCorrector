using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocxCorrector.Models;
using ServiceStack.Text;

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
            CsvConfig.ItemSeperatorString = ";";
            string csvString = CsvSerializer.SerializeToCsv(listData);

            WriteToFile(filePath, csvString);
        }

        // Заполнить CSV файл для свойств параграфов
        public static void FillCSV(string filePath, List<ParagraphProperties> listData)
        {
            List<ParagraphPropertiesInterop> listDataInterop = listData.OfType<ParagraphPropertiesInterop>().ToList();
            if (listDataInterop.Count != 0)
            {
                FillCSV(filePath: filePath, listData: listDataInterop);
                return;
            }

            List<ParagraphPropertiesGemBox> listDataGemBox = listData.OfType<ParagraphPropertiesGemBox>().ToList();
            if (listDataGemBox.Count != 0)
            {
                FillCSV(filePath: filePath, listData: listDataGemBox);
                return;
            }

            List<ParagraphPropertiesSpire> listDataSpire = listData.OfType<ParagraphPropertiesSpire>().ToList();
            if (listDataSpire.Count != 0)
            {
                FillCSV(filePath: filePath, listData: listDataSpire);
                return;
            }
        }
    }
}

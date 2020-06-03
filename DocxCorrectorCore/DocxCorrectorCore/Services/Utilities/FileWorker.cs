using System;
using System.Collections.Generic;
using System.IO;
using ServiceStack.Text;
using DocxCorrectorCore.Services.Helpers;
using Word = GemBox.Document;

namespace DocxCorrectorCore.Services.Utilities
{
    public static class FileWorker
    {
        // Записать текст text в файл, расположенный в filePath
        public static void WriteToFile(string filePath, string text)
        {
            try
            {
                using StreamWriter sw = new StreamWriter(filePath, false, System.Text.Encoding.UTF8);
                sw.WriteLine(text);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        // Записать в CSV файл filePath объекты из списка listData
        public static void FillCSV<T>(string filePath, List<T> listData)
        {
            CsvConfig.ItemSeperatorString = ";";
            string csvString = CsvSerializer.SerializeToCsv(listData);

            WriteToFile(filePath, csvString);
        }

        // TODO: REMOVE??
        //// Заполнить CSV файл для свойств параграфов
        //public static void FillCSV(string filePath, List<ParagraphProperties> listData)
        //{
        //    List<ParagraphPropertiesGemBox> listDataGemBox = listData.OfType<ParagraphPropertiesGemBox>().ToList();
        //    if (listDataGemBox.Count != 0)
        //    {
        //        FillCSV(filePath: filePath, listData: listDataGemBox);
        //        return;
        //    }
        //}

        // Сохранить документ filePath как pdf в директории resultDirPath
        public static void SaveDocumentAsPdf(string filePath, string resultDirPath)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return; }

            string resultFilePath = Path.Combine(resultDirPath, $"{Path.GetFileNameWithoutExtension(filePath)}.pdf");
            document.Save(resultFilePath);
        }

        // Сохранить страницы документа filePath как отдельные pdf в директории resultDirPath
        public static void SavePagesAsPdf(string filePath, string resultDirPath)
        {
            Word.DocumentModel? document = GemBoxHelper.OpenDocument(filePath: filePath);
            if (document == null) { return; }

            var pages = document.GetPaginator().Pages;

            int pageNumber = 1;
            foreach (var page in pages)
            {
                string resultFilePath = Path.Combine(resultDirPath, $"{pageNumber}.pdf");
                page.Save(resultFilePath);
                pageNumber++;
            }
        }

        // TODO: REMOVE??
        //// Получить все текстовое содержимое файла
        //public static string? GetContentOfTextFile(string filePath)
        //{
        //    try
        //    {
        //        return File.ReadAllText(filePath);
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine(e.Message);
        //        return null;
        //    }
        //}
    }
}

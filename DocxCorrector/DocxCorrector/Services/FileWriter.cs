using System;
using System.IO;

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
    }
}

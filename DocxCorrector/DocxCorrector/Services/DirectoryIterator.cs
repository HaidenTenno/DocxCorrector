using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;

namespace DocxCorrector.Services
{
    public static class DirectoryIterator
    {
        // Выполнить в каждой из поддиректории директории directoryPath функцию action, принимающую строку
        public static void IterateDir(string directoryPath, Action<string> action)
        {
            string[] subDirs = Directory.GetDirectories(directoryPath);

            foreach (string subDir in subDirs)
            {
#if DEBUG
                Console.WriteLine(subDir);
#endif
                action(subDir);
            }
        }

        // Выполнить для каждого docx файла в директории path функцию action, принимающую строку
        public static void IterateDocxFiles(string path, Action<string> action)
        {
            var files = Directory.EnumerateFiles(path, "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".docx") || s.EndsWith(".doc"));

            foreach (string file in files)
            {
#if DEBUG
                Console.WriteLine(file);
#endif
                action(file);
            }
        }
    }
}

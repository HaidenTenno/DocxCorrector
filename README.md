# DocxCorrector
## Консольное приложение для автоматизации нормоконтроля docx файлов и получения свойств элементов документа

Для запуска требуется:
* Visual Studio / Rider
* .NET Core 3.1

Настройка после клонирования:
* Поменять аргументы запуска в Visual studio (ПКМ по проекту &rarr; Отладка &rarr; Аргументы приложения) (Аргументы см. ниже)

======================================================
```
Usage:
  DocxCorrectorCore [options] [command]

Options:
   --version         Show version information
  -?, -h, --help    Show help and usage information

Commands:
  correct <file-to-correct> <GOST|ITMO> <paragraphs-classes> <result-dir>    Analyze the document for formatting errors using the selected rules and class list and save the result
  interactive                                                                Start the program interactively
```
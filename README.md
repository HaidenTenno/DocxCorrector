# DocxCorrector
## Консольное приложение для автоматизации нормоконтроля docx файлов и получения свойств элементов документа

Для запуска требуется:
* Visual Studio / Rider
* .NET Core 3.1

Настройка после клонирования:
* Поменять аргументы запуска, если запускаете из IDE (в Visual Studio ПКМ по проекту &rarr; Отладка &rarr; Аргументы приложения) (Аргументы см. ниже)

======================================================
```
Usage:
  DocxCorrectorCore [options] [command]

Options:
  --version         Show version information
  -?, -h, --help    Show help and usage information

Commands:
  correct <file-to-correct> <GOST|ITMO> <paragraphs-classes> <result-path>    Analyze the document for formatting
                                                                              errors using the selected rules and
                                                                              class list and save the result
  pull <file-to-analyse> <result-path>                                        Pull out the properties of document
                                                                              paragraphs and save them in csv
  interactive                                                                 Start the program interactively
```
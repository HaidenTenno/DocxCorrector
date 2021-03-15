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
  correct <file-to-correct> <GOST|ITMO>                       Analyze the document for formatting errors using the
  <paragraphs-classes> <result-path>                          selected rules and class list and save the result

  pull <file-to-analyse> <initial-paragraph-id>               Pull out the properties of document paragraphs and save
  <result-path1> <result-path2>                               them in csv

  interactive                                                 Start the program interactively

  pullWithPresets <file-to-analyse> <presets-file>            Pull out the properties of document paragraphs, try to
  <initial-paragraph-id> <result-path>                        classify the paragarps using presets, save result in csv

  correctParagraph <file-to-correct> <GOST|ITMO>              Analyze the paragraph of the document for formatting
  <paragraph-id>                                              errors using the selected rules and save the result
  <result-path>
  <b0|b1|b2|b3|b4|c0|c1|c2|c3|d0|d1|d2|d3|d4|d5|d6|e0|f0|f
  1|f3|f5|f6|g0|g1|g2|g3|h0|h1|h2|h3|h4|i0|i1|i2|j0|n0|NoC
  lass|r0>

  createModelFile <GOST|ITMO>                                 Get the file that contains the rules for paragraphs of a
  <result-path>                                               certain class for the selected requirements
  <b0|b1|b2|b3|b4|c0|c1|c2|c3|d0|d1|d2|d3|d4|d5|d6|e0|f0|f
  1|f3|f5|f6|g0|g1|g2|g3|h0|h1|h2|h3|h4|i0|i1|i2|j0|n0|NoC
  lass|r0>
```
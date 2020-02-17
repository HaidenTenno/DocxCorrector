# DocxCorrector
## Консольное приложение для получения списка ошибок оформления в docx файле

Для запуска требуется:
* ОС Windows
* VisualStudio
* Установленный Word (проверялось только с лицензией Office 365)

Настройка после клонирования:
* В файле Config.cs изменить пути для файлов на нужные
* Переподключить библитеку Microsoft.Office.Interop.Word в обозревателе решений
    * Зависимости -> Сборки -> Microsoft.Office.Interop.Word, ПКМ -> Удалить
    * Зависимости, ПКМ -> Добавить ссылку, Обзор
    * %Директория установки Visual Studio%\Shared\Visual Studio Tools for Office\PIA\Office15\Microsoft.Office.Interop.Word.dll -> ОК
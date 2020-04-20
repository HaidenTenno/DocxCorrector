using System;
using Pdf = GemBox.Pdf;

namespace DocxCorrectorCore.Services.Helpers
{
    public sealed class GemBoxPdfHelper
    {
        // Private
        // Ввод лицензионного ключа
        private void SetLicense()
        {
            Pdf.ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        }

        // Открыть документ filePath
        private Pdf.PdfDocument? OpenDocument(string filePath)
        {
            try
            {
                Pdf.PdfDocument document = Pdf.PdfDocument.Load(filePath);
                return document;
            }
            catch (Exception ex)
            {
#if DEBUG
                Console.WriteLine(ex.Message);
#endif
                Console.WriteLine("Can't open document");
                return null;
            }
        }

        // Public
        public GemBoxPdfHelper()
        {
            SetLicense();
        }

        // Напечатать содержимое документа filePath
        public void PrintContent(string filePath)
        {
            Pdf.PdfDocument? document = OpenDocument(filePath: filePath);
            if (document == null) { return; }

            foreach (var page in document.Pages)
            {
                Console.WriteLine(page.Content.ToString());
            }
        }
    }
}

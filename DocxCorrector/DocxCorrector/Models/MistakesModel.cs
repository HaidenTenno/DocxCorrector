
namespace DocxCorrector.Models
{
    // TODO: - Improve model
    public class Mistake
    {
        // ID параграфа
        public int ParagraphID { get; set; }
        // Сообщение об ошибке
        public string Message { get; set; }
    }
}
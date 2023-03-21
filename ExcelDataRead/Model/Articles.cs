namespace ExcelDataRead.Model
{
    public class Articles
    {
        public string ArticleTitle { get; set; } = string.Empty;
        public int Article_Order { get; set; }  
        public virtual List<Paragraph>? Paragraphs { get; set; }
    }
}

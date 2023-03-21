using ExcelDataRead.Model;
using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.Text.Json;


namespace ExcelDataRead.Controllers
{
    [Route("[controller]")]
    [ApiController]
    public class ExcelDataReadController : ControllerBase
    {


        [HttpGet]
        public async Task<IActionResult> GetData()
        {

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using var stream = System.IO.File.Open("C://Users//Kağan//Desktop//testdata2.xlsx", FileMode.Open, FileAccess.Read);
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                int article_counter = 1;
                List<string> articles = new List<string>();
                List<string> paragraphs = new List<string>();
                List<string> euTexts = new List<string>();
                List<string> crossReference = new List<string>();
                List<string> alignmentStatus = new List<string>();
                List<string> comments = new List<string>();
                List<string> proposal = new List<string>();
                List<string> timing = new List<string>();
                List<string> instituitons = new List<string>();
                List<string> budget = new List<string>();

                while (reader.Read())
                {

                    if (reader.GetString(0) == "Article" || reader.GetString(0) == null || reader.GetValue(1) == null)
                    {
                        _ = reader.Read() == false;
                    }

                    int counter = 1;
                    string[] rows = reader.GetString(0).Split(" \n");


                    articles.Add (rows[1].Trim());
                    paragraphs.Add (reader.GetValue(1).ToString());
                    euTexts.Add(reader.GetValue(2).ToString());
                    crossReference.Add(reader.GetValue(3).ToString());
                    alignmentStatus.Add(reader.GetValue(4).ToString());
                    comments.Add(reader.GetValue(5).ToString());
                    proposal.Add(reader.GetValue(6).ToString());
                    timing.Add(reader.GetValue(7).ToString());
                    instituitons.Add(reader.GetValue(8).ToString());
                    budget.Add(reader.GetValue(9).ToString());
                }

                foreach (var item in articles.Distinct())
                {
                    Articles article = new Articles();
                    article.Article_Order= article_counter;
                    article.ArticleTitle = item;

                    var firstIndex = articles.IndexOf(item);
                    var lastIndex = articles.LastIndexOf(item);
                    List<Paragraph> ps = new List<Paragraph>();
                    for (int i = firstIndex; i <= lastIndex; i++)
                    {
                        Paragraph paragraph = new Paragraph();
                        paragraph.ParagraphNumber = paragraphs[i];
                        paragraph.ParagraphText = euTexts[i];
                        paragraph.CrossReference = crossReference[i];
                        paragraph.AlignmentStatus = alignmentStatus[i];
                        paragraph.Comments = comments[i];
                        paragraph.Proposal = proposal[i];
                        paragraph.Timing = timing[i];
                        paragraph.Instituitons = instituitons[i];
                        paragraph.Budget = budget[i];
                        ps.Add(paragraph);
                        
                        
                    }
                    article.Paragraphs = ps;
                    var options = new JsonSerializerOptions { WriteIndented = true };
                    string jsonString = JsonSerializer.Serialize(article, options);
                    Console.WriteLine(jsonString);
                    article_counter++;
                }
                
                reader.Close();

            }

            //return Ok(valuesarray);
            return Ok();
        }

    }
}
using GenerateDocument.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Diagnostics;

namespace GenerateDocument.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        //Generate Document
        public async Task<string> GetDocument()
        {
            var Name = "Testing Pawan";
            var call = "720996187";
            var ABC = true;
            var date = DateTime.Now;
            string targetFilePath = Path.Combine(@"D:\GenerateDoc\GenerateDocument\GenerateDocument\wwwroot\DocTemp.docx");
            FileStream fileStreamPath = new FileStream(targetFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx);

            document.Open(fileStreamPath, FormatType.Docx);
            fileStreamPath.Dispose();
            document.Replace("<<TodayDate>>", date.ToString("d"), true, true);
            document.Replace("<<ContactNumber>>", call, true, true);
            document.Replace("<<Name>>", Name, true, true);

            string[] names = { "Apple", "Banana", "Orange", "Grape", "Mango" };
            int[] price = { 121, 443, 553, 889, 667 };


            string targetFilePath3 = Path.Combine(@"D:\GenerateDoc\GenerateDocument\GenerateDocument\wwwroot\DocTemp.docx");
            FileStream fileStreamPath3 = new FileStream(targetFilePath3, FileMode.OpenOrCreate, FileAccess.Read, FileShare.ReadWrite);
            WordDocument tmpDoc = new WordDocument();
            tmpDoc.Open(fileStreamPath3, FormatType.Docx);

            IWSection section = tmpDoc.AddSection();
            section.PageSetup.Margins.All = 72;
            section.PageSetup.PageSize = new Syncfusion.Drawing.SizeF(612, 792);

            var count = 1;
            if (ABC = true)
            {
                IWTable table = section.AddTable();
                table.TableFormat.Borders.LineWidth = 1f;
                WTableRow row = table.AddRow();
                WTableCell cell = row.AddCell();
                cell.Width = 270;
                cell.AddParagraph().AppendText("Item Description");
                cell = row.AddCell();
                cell.Width = 180;
                cell.AddParagraph().AppendText("Cost");
                row = table.AddRow(true, false);

                foreach (var name in names)
                {
                    cell = row.AddCell();
                    cell.Width = 270;
                    cell.AddParagraph().AppendText("Fruit Name" + name);
                }
                foreach (var cost in price)
                {
                    cell = row.AddCell();
                    cell.Width = 270;
                    cell.AddParagraph().AppendText("Price" + cost);
                }
                count++;
            }


            TextBodyPart replacePart = new TextBodyPart(tmpDoc);
            TextBodySelection textSel = new TextBodySelection(tmpDoc.LastSection.Body, 0, count - 1, 0, 1);

            //Copy the selected area into the TextBodyPart
            replacePart.Copy(textSel);

            //Replace Text with image and text.
            document.Replace("<<Proposal_Cost_Table>>", replacePart, false, true);

            document.Save(fileStreamPath, FormatType.Docx);

            fileStreamPath.Dispose();

            return null;
        }


    }
}
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace EverConsumerUtilities.Print
{
    class PrintTrnSalesInvoiceDetailPDF
    {
        public PrintTrnSalesInvoiceDetailPDF(Int32 SIId)
        {
            try
            {
                var httpWebRequest = (HttpWebRequest)WebRequest.Create("http://everconsumer.easyfis.com/api/everConsumerIntegration/salesInvoice/detail/" + SIId);
                httpWebRequest.Method = "GET";
                httpWebRequest.Accept = "application/json";

                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    var result = streamReader.ReadToEnd();
                    JavaScriptSerializer js = new JavaScriptSerializer();
                    Entities.TrnSalesInvoice deserializedSales = (Entities.TrnSalesInvoice)js.Deserialize(result, typeof(Entities.TrnSalesInvoice));

                    String fileName = "SalesInvoiceDetail_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";

                    Font fontTimesNewRoman7 = FontFactory.GetFont(BaseFont.TIMES_ROMAN, 7);
                    Font fontTimesNewRoman10 = FontFactory.GetFont(BaseFont.TIMES_ROMAN, 10);
                    Font fontTimesNewRoman10Italic = FontFactory.GetFont(BaseFont.TIMES_ROMAN, 10, Font.ITALIC);
                    Font fontTimesNewRoman11 = FontFactory.GetFont(BaseFont.TIMES_ROMAN, 11);
                    Font fontTimesNewRoman11Bold = FontFactory.GetFont(BaseFont.TIMES_ROMAN, 11, Font.BOLD);

                    Paragraph line = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(0.5F, 100.0F, BaseColor.DARK_GRAY, Element.ALIGN_MIDDLE, 10F)));
                    Rectangle pagesize = new Rectangle(504, 792);

                    Document document = new Document(pagesize);
                    document.SetMargins(50f, 50f, 50f, 0f);

                    PdfWriter pdfWriter = PdfWriter.GetInstance(document, new FileStream(fileName, FileMode.Create));

                    document.Open();

                    PdfPTable pdfTableSalesInvoiceHeader = new PdfPTable(3);
                    pdfTableSalesInvoiceHeader.WidthPercentage = 100;
                    pdfTableSalesInvoiceHeader.SetWidths(new float[] { 50f, 25f, 25f });

                    // invoice date
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.SIDate, fontTimesNewRoman7)) { Border = 0, Colspan = 3, HorizontalAlignment = 2, PaddingBottom = 12f });

                    // Term
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(" ", fontTimesNewRoman7)) { Border = 0 });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.Term, fontTimesNewRoman7)) { Border = 0, Colspan = 2, PaddingLeft = 50f, PaddingBottom = -8f });

                    // Customer, Address, Sales Person 
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.Customer + "\n" + deserializedSales.Address, fontTimesNewRoman7)) { Border = 0, Rowspan = 3 });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.SoldBy, fontTimesNewRoman7)) { Border = 0, Colspan = 2, PaddingLeft = 50f, PaddingBottom = 12f });

                    // SO Number Etc...
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.ManualSONumber, fontTimesNewRoman7)) { Border = 0, PaddingLeft = 50f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(" ", fontTimesNewRoman7)) { Border = 0, PaddingLeft = 50f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(" ", fontTimesNewRoman7)) { Border = 0, PaddingLeft = 50f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(" ", fontTimesNewRoman7)) { Border = 0, PaddingLeft = 50f });

                    // TIN and Business Style
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.TIN, fontTimesNewRoman7)) { Border = 0, PaddingLeft = 30f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.CustomerGroup, fontTimesNewRoman7)) { Border = 0, Colspan = 2, PaddingLeft = 70f });

                    // Notes
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.Remarks, fontTimesNewRoman7)) { Border = 0, PaddingLeft = 30f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(" ", fontTimesNewRoman7)) { Border = 0, Colspan = 2 });

                    document.Add(pdfTableSalesInvoiceHeader);

                    document.Close();

                    //Process.Start(fileName);

                    ProcessStartInfo info = new ProcessStartInfo(fileName)
                    {
                        Verb = "Print",
                        CreateNoWindow = true,
                        WindowStyle = ProcessWindowStyle.Hidden
                    };

                    Process printDwg = new Process
                    {
                        StartInfo = info
                    };

                    printDwg.Start();
                    printDwg.Close();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}

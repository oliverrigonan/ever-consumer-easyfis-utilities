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
                var httpWebRequest = (HttpWebRequest)WebRequest.Create("http://everconsumer.easyfis.com/api/everConsumerIntegration/salesInvoiceItem/list/" + SIId);
                httpWebRequest.Method = "GET";
                httpWebRequest.Accept = "application/json";

                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    var result = streamReader.ReadToEnd();
                    JavaScriptSerializer js = new JavaScriptSerializer();
                    List<Entities.TrnSalesInvoiceItem> deserializedSalesInvoiceItems = (List<Entities.TrnSalesInvoiceItem>)js.Deserialize(result, typeof(List<Entities.TrnSalesInvoiceItem>));

                    String fileName = "SalesInvoiceDetail_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";

                    Font fontVerdana9 = FontFactory.GetFont("Verdana", 9);

                    Paragraph line = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(0.5F, 100.0F, BaseColor.DARK_GRAY, Element.ALIGN_MIDDLE, 10F)));
                    Rectangle pagesize = new Rectangle(504, 792);

                    Document document = new Document(pagesize);
                    document.SetMargins(50f, 50f, 168f, 140f);

                    PdfWriter pdfWriter = PdfWriter.GetInstance(document, new FileStream(fileName, FileMode.Create));

                    document.Open();

                    Entities.TrnSalesInvoiceVATAnalysis VATAnalysis = new Entities.TrnSalesInvoiceVATAnalysis();
                    if (deserializedSalesInvoiceItems != null)
                    {
                        if (deserializedSalesInvoiceItems.Any())
                        {
                            VATAnalysis = new Entities.TrnSalesInvoiceVATAnalysis()
                            {
                                VATSales = 0,
                                VATZeroRatedSales = 0,
                                VATExemptSales = 0,
                                LessDiscount = 0,
                                TotalSales = 0,
                                VAT = 0,
                                TotalAmountDue = 0
                            };

                            foreach (var salesInvoiceItem in deserializedSalesInvoiceItems)
                            {
                                if (salesInvoiceItem.VAT == "VAT Output")
                                {
                                    VATAnalysis.VATSales += (salesInvoiceItem.Price * salesInvoiceItem.Quantity) - ((salesInvoiceItem.Price * salesInvoiceItem.Quantity) / (1 + (salesInvoiceItem.VATPercentage / 100)) * (salesInvoiceItem.VATPercentage / 100));
                                }

                                VATAnalysis.VATZeroRatedSales += salesInvoiceItem.Amount;

                                if (salesInvoiceItem.VAT == "VAT Exempt")
                                {
                                    VATAnalysis.VATExemptSales += (salesInvoiceItem.Price * salesInvoiceItem.Quantity) - ((salesInvoiceItem.Price * salesInvoiceItem.Quantity) / (1 + (salesInvoiceItem.VATPercentage / 100)) * (salesInvoiceItem.VATPercentage / 100));
                                }

                                VATAnalysis.LessDiscount += salesInvoiceItem.DiscountAmount;
                                VATAnalysis.TotalSales += salesInvoiceItem.Amount + salesInvoiceItem.DiscountAmount;
                                VATAnalysis.VAT += salesInvoiceItem.VATAmount;
                                VATAnalysis.TotalAmountDue += salesInvoiceItem.Amount;
                            }
                        }
                    }

                    pdfWriter.PageEvent = new SalesInvoiceHeaderFooter(SIId, VATAnalysis);

                    if (deserializedSalesInvoiceItems != null)
                    {
                        if (deserializedSalesInvoiceItems.Any())
                        {
                            // Sales Invoice Items
                            PdfPTable pdfTableSalesInvoiceItems = new PdfPTable(7);
                            pdfTableSalesInvoiceItems.WidthPercentage = 100;
                            pdfTableSalesInvoiceItems.SetWidths(new float[] { 40f, 15f, 5f, 7f, 10f, 10f, 13f });

                            foreach (var salesInvoiceItem in deserializedSalesInvoiceItems)
                            {
                                pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase(salesInvoiceItem.ItemDescription, fontVerdana9)) { Border = 0 });
                                pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase(salesInvoiceItem.BarcCode, fontVerdana9)) { Border = 0 });
                                pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase(" ", fontVerdana9)) { Border = 0 });
                                pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase(salesInvoiceItem.Quantity.ToString("#,##0"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                                pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase(salesInvoiceItem.Unit, fontVerdana9)) { Border = 0 });
                                pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase(salesInvoiceItem.Price.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                                pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase(salesInvoiceItem.Amount.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                            }

                            document.Add(pdfTableSalesInvoiceItems);
                        }
                        else
                        {
                            document.Add(line);
                        }
                    }
                    else
                    {
                        document.Add(line);
                    }

                    document.Close();

                    Process.Start(fileName);

                    //ProcessStartInfo info = new ProcessStartInfo(fileName)
                    //{
                    //    Verb = "Print",
                    //    CreateNoWindow = true,
                    //    WindowStyle = ProcessWindowStyle.Hidden
                    //};

                    //Process printDwg = new Process
                    //{
                    //    StartInfo = info
                    //};

                    //printDwg.Start();
                    //printDwg.Close();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }

    class SalesInvoiceHeaderFooter : PdfPageEventHelper
    {
        private Int32 SIId = 0;
        private Entities.TrnSalesInvoiceVATAnalysis VATAnalysis;

        public SalesInvoiceHeaderFooter(Int32 _SIId, Entities.TrnSalesInvoiceVATAnalysis _VATAnalysis)
        {
            SIId = _SIId;
            VATAnalysis = _VATAnalysis;
        }

        public override void OnEndPage(PdfWriter writer, Document document)
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

                    Font fontVerdana9 = FontFactory.GetFont("Verdana", 9);

                    PdfPTable pdfTableSalesInvoiceHeader = new PdfPTable(3);
                    pdfTableSalesInvoiceHeader.DefaultCell.Border = 0;
                    pdfTableSalesInvoiceHeader.TotalWidth = document.PageSize.Width - document.LeftMargin - document.RightMargin;
                    pdfTableSalesInvoiceHeader.LockedWidth = true;
                    pdfTableSalesInvoiceHeader.SetWidths(new float[] { 50f, 25f, 25f });

                    // SI Number
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.SINumber, fontVerdana9)) { Border = 0, Colspan = 3, HorizontalAlignment = 2 });

                    // Invoice date
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.SIDate, fontVerdana9)) { Border = 0, Colspan = 3, HorizontalAlignment = 2, PaddingBottom = 12f });

                    // Term
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(" ", fontVerdana9)) { Border = 0 });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.Term, fontVerdana9)) { Border = 0, Colspan = 2, PaddingLeft = 50f, PaddingBottom = -8f });

                    // Customer, Address, Sales Person 
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.Customer + "\n" + deserializedSales.Address, fontVerdana9)) { Border = 0, Rowspan = 3 });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.SoldBy, fontVerdana9)) { Border = 0, Colspan = 2, PaddingLeft = 50f, PaddingBottom = 12f });

                    // SO Number Etc...
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.ManualSONumber, fontVerdana9)) { Border = 0, PaddingLeft = 50f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(" ", fontVerdana9)) { Border = 0, PaddingLeft = 50f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(" ", fontVerdana9)) { Border = 0, PaddingLeft = 50f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(" ", fontVerdana9)) { Border = 0, PaddingLeft = 50f });

                    // TIN and Business Style
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.TIN, fontVerdana9)) { Border = 0, PaddingLeft = 30f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.CustomerGroup, fontVerdana9)) { Border = 0, Colspan = 2, PaddingLeft = 70f });

                    // Notes
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.Remarks, fontVerdana9)) { Border = 0, PaddingLeft = 30f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(" ", fontVerdana9)) { Border = 0, Colspan = 2 });

                    pdfTableSalesInvoiceHeader.WriteSelectedRows(0, -1, document.LeftMargin, writer.PageSize.GetTop(document.TopMargin - 132f), writer.DirectContent);

                    // Footer - VAT Analysis
                    PdfPTable pdfTableSalesInvoiceFooter = new PdfPTable(1);
                    pdfTableSalesInvoiceFooter.DefaultCell.Border = 0;
                    pdfTableSalesInvoiceFooter.TotalWidth = document.PageSize.Width - document.LeftMargin - document.RightMargin;
                    pdfTableSalesInvoiceFooter.LockedWidth = true;
                    pdfTableSalesInvoiceFooter.SetWidths(new float[] { 100f });
                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(VATAnalysis.VATSales.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(VATAnalysis.VATZeroRatedSales.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(VATAnalysis.VATExemptSales.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(VATAnalysis.LessDiscount.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(VATAnalysis.TotalSales.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(VATAnalysis.VAT.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(VATAnalysis.TotalAmountDue.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                    pdfTableSalesInvoiceFooter.WriteSelectedRows(0, -1, document.LeftMargin, writer.PageSize.GetBottom(document.BottomMargin - 20f), writer.DirectContent);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}

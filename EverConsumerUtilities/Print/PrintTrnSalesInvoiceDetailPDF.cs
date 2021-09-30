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
        public static String GetMoneyWord(String input)
        {
            String decimals = "";
            if (input.Contains("."))
            {
                decimals = input.Substring(input.IndexOf(".") + 1);
                input = input.Remove(input.IndexOf("."));
            }

            String strWords = GetMoreThanThousandNumberWords(input);
            if (decimals.Length > 0)
            {
                if (Convert.ToDecimal(decimals) > 0)
                {
                    String getFirstRoundedDecimals = new String(decimals.Take(2).ToArray());
                    strWords += " Pesos And " + GetMoreThanThousandNumberWords(getFirstRoundedDecimals) + " Cents Only";
                }
                else
                {
                    strWords += " Pesos Only";
                }
            }
            else
            {
                strWords += " Pesos Only";
            }

            return strWords;
        }

        private static String GetMoreThanThousandNumberWords(string input)
        {
            try
            {
                String[] seperators = { "", " Thousand ", " Million ", " Billion " };

                int i = 0;

                String strWords = "";

                while (input.Length > 0)
                {
                    String _3digits = input.Length < 3 ? input : input.Substring(input.Length - 3);
                    input = input.Length < 3 ? "" : input.Remove(input.Length - 3);

                    Int32 no = Int32.Parse(_3digits);
                    _3digits = GetHundredNumberWords(no);

                    _3digits += seperators[i];
                    strWords = _3digits + strWords;

                    i++;
                }

                return strWords;
            }
            catch
            {
                return "Invalid Amount";
            }
        }

        private static String GetHundredNumberWords(Int32 no)
        {
            String[] Ones =
            {
                "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven",
                "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Ninteen"
            };

            String[] Tens = { "Ten", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety" };
            String word = "";

            if (no > 99 && no < 1000)
            {
                Int32 i = no / 100;
                word = word + Ones[i - 1] + " Hundred ";
                no = no % 100;
            }

            if (no > 19 && no < 100)
            {
                Int32 i = no / 10;
                word = word + Tens[i - 1] + " ";
                no = no % 10;
            }

            if (no > 0 && no < 20)
            {
                word = word + Ones[no - 1];
            }

            return word;
        }

        public PrintTrnSalesInvoiceDetailPDF(Int32 SIId)
        {
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                // GET SALES
                var httpWebRequestSI = (HttpWebRequest)WebRequest.Create("http://mjr.liteclerk.com/api/everConsumerIntegration/salesInvoice/detail/" + SIId);
                httpWebRequestSI.Method = "GET";
                httpWebRequestSI.Accept = "application/json";

                var httpResponseSI = (HttpWebResponse)httpWebRequestSI.GetResponse();
                using (var streamReaderSI = new StreamReader(httpResponseSI.GetResponseStream()))
                {
                    var resultSI = streamReaderSI.ReadToEnd();
                    JavaScriptSerializer jsSI = new JavaScriptSerializer();
                    Entities.TrnSalesInvoice deserializedSales = (Entities.TrnSalesInvoice)jsSI.Deserialize(resultSI, typeof(Entities.TrnSalesInvoice));

                    // GET SALES ITEMS
                    var httpWebRequest = (HttpWebRequest)WebRequest.Create("http://mjr.liteclerk.com/api/everConsumerIntegration/salesInvoiceItem/list/" + SIId);
                    httpWebRequest.Method = "GET";
                    httpWebRequest.Accept = "application/json";

                    var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                    using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                    {
                        var result = streamReader.ReadToEnd();
                        JavaScriptSerializer js = new JavaScriptSerializer();
                        List<Entities.TrnSalesInvoiceItem> deserializedSalesInvoiceItems = (List<Entities.TrnSalesInvoiceItem>)js.Deserialize(result, typeof(List<Entities.TrnSalesInvoiceItem>));

                        String fileName = "SalesInvoiceDetail_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                        Font fontVerdana9 = FontFactory.GetFont("Times-Roman", 9);

                        Paragraph line = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(1F, 100.0F, BaseColor.DARK_GRAY, Element.ALIGN_MIDDLE, 10F)));

                        Document document = new Document(PageSize.LETTER);
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
                                PdfPTable pdfTableSalesInvoiceItems = new PdfPTable(8);
                                pdfTableSalesInvoiceItems.WidthPercentage = 100;
                                pdfTableSalesInvoiceItems.SetWidths(new float[] { 15f, 40f, 10f, 10f, 10f, 10f, 10f, 13f });

                                Decimal totalAmount = 0;
                                Decimal totalQuantity = 0;

                                foreach (var salesInvoiceItem in deserializedSalesInvoiceItems)
                                {
                                    pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase(salesInvoiceItem.BarcCode, fontVerdana9)) { Border = 0 });
                                    pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase(salesInvoiceItem.ItemDescription, fontVerdana9)) { Border = 0 });
                                    pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase(salesInvoiceItem.Quantity.ToString("#,##0"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                                    pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase(salesInvoiceItem.Unit, fontVerdana9)) { Border = 0 });
                                    pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase(salesInvoiceItem.Price.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                                    pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase(salesInvoiceItem.DiscountAmount.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                                    pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase("0.00", fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                                    pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase(salesInvoiceItem.Amount.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });

                                    totalAmount += salesInvoiceItem.Amount;
                                    totalQuantity += salesInvoiceItem.Quantity;
                                }

                                pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase("<<<< Nothing Follows >>>>", fontVerdana9)) { Border = 0, HorizontalAlignment = 1, Colspan = 8 });
                                pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase("", fontVerdana9)) { Border = 0, Colspan = 2 });
                                pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase(line)) { Border = 0, Colspan = 2 });
                                pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase("", fontVerdana9)) { Border = 0, Colspan = 3 });
                                pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase(line)) { Border = 0 });
                                pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase("", fontVerdana9)) { Border = 0, Colspan = 2, PaddingTop = -7f });
                                pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase(totalQuantity.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2, PaddingTop = -7f, Colspan = 2 });
                                pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase("", fontVerdana9)) { Border = 0, Colspan = 3, PaddingTop = -7f });
                                pdfTableSalesInvoiceItems.AddCell(new PdfPCell(new Phrase(totalAmount.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2, PaddingTop = -7f });
                                document.Add(pdfTableSalesInvoiceItems);

                                String amount = Convert.ToString(Math.Round(totalAmount * 100) / 100);
                                String amountString = GetMoneyWord(amount).ToUpper();

                                PdfPTable pdfTableSalesInvoiceAmountInWords = new PdfPTable(2);
                                pdfTableSalesInvoiceAmountInWords.DefaultCell.Border = 0;
                                pdfTableSalesInvoiceAmountInWords.TotalWidth = document.PageSize.Width - document.LeftMargin - document.RightMargin;
                                pdfTableSalesInvoiceAmountInWords.LockedWidth = true;
                                pdfTableSalesInvoiceAmountInWords.SetWidths(new float[] { 15f, 110f });
                                pdfTableSalesInvoiceAmountInWords.AddCell(new PdfPCell(new Phrase(" ", fontVerdana9)) { Border = 0, HorizontalAlignment = 0, PaddingTop = 10f });
                                pdfTableSalesInvoiceAmountInWords.AddCell(new PdfPCell(new Phrase(amountString, fontVerdana9)) { Border = 0, HorizontalAlignment = 0, PaddingTop = 10 });
                                pdfTableSalesInvoiceAmountInWords.AddCell(new PdfPCell(new Phrase("Remarks", fontVerdana9)) { Border = 0, HorizontalAlignment = 0, PaddingTop = 10 });
                                pdfTableSalesInvoiceAmountInWords.AddCell(new PdfPCell(new Phrase(deserializedSales.Remarks, fontVerdana9)) { Border = 0, HorizontalAlignment = 0, PaddingTop = 10 });
                                document.Add(pdfTableSalesInvoiceAmountInWords);
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
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                // GET SALES
                var httpWebRequest = (HttpWebRequest)WebRequest.Create("http://mjr.liteclerk.com/api/everConsumerIntegration/salesInvoice/detail/" + SIId);
                httpWebRequest.Method = "GET";
                httpWebRequest.Accept = "application/json";

                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    var result = streamReader.ReadToEnd();
                    JavaScriptSerializer js = new JavaScriptSerializer();
                    Entities.TrnSalesInvoice deserializedSales = (Entities.TrnSalesInvoice)js.Deserialize(result, typeof(Entities.TrnSalesInvoice));

                    Font fontVerdana9 = FontFactory.GetFont("Times-Roman", 9);

                    PdfPTable pdfTableSalesInvoiceHeader = new PdfPTable(4);

                    pdfTableSalesInvoiceHeader.DefaultCell.Border = 0;
                    pdfTableSalesInvoiceHeader.TotalWidth = document.PageSize.Width - document.LeftMargin - document.RightMargin;
                    pdfTableSalesInvoiceHeader.LockedWidth = true;
                    pdfTableSalesInvoiceHeader.SetWidths(new float[] { 15f, 70f, 15f, 25f });

                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(" ", fontVerdana9)) { Border = 0, HorizontalAlignment = 0, PaddingBottom = -2f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase("PARTICULARS HERE", fontVerdana9)) { Border = 0, HorizontalAlignment = 0, PaddingBottom = -2f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(" ", fontVerdana9)) { Border = 0, HorizontalAlignment = 0, PaddingBottom = -2f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.SIDate, fontVerdana9)) { Border = 0, HorizontalAlignment = 2, PaddingBottom = -2f });

                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(" ", fontVerdana9)) { Border = 0, HorizontalAlignment = 0, PaddingBottom = -2f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.Customer, fontVerdana9)) { Border = 0, HorizontalAlignment = 0, PaddingBottom = -2f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(" ", fontVerdana9)) { Border = 0, HorizontalAlignment = 0, PaddingBottom = -2f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.Term, fontVerdana9)) { Border = 0, HorizontalAlignment = 2, PaddingBottom = -2f });

                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(" ", fontVerdana9)) { Border = 0, HorizontalAlignment = 0, PaddingBottom = -2f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.Address, fontVerdana9)) { Border = 0, HorizontalAlignment = 0, PaddingBottom = -2f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase("Doc #: ", fontVerdana9)) { Border = 0, HorizontalAlignment = 0, PaddingBottom = -2f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.ManualSINumber, fontVerdana9)) { Border = 0, HorizontalAlignment = 2, PaddingBottom = -2f });

                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(" ", fontVerdana9)) { Border = 0, HorizontalAlignment = 0, PaddingBottom = -2f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(" ", fontVerdana9)) { Border = 0, HorizontalAlignment = 0, PaddingBottom = -2f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase("PO #: ", fontVerdana9)) { Border = 0, HorizontalAlignment = 0, PaddingBottom = -2f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.DocumentReference, fontVerdana9)) { Border = 0, HorizontalAlignment = 2, PaddingBottom = -2f });

                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase("Salesman", fontVerdana9)) { Border = 0, HorizontalAlignment = 0, PaddingBottom = -2f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.SoldBy, fontVerdana9)) { Border = 0, HorizontalAlignment = 0, PaddingBottom = -2f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase("TIN #: ", fontVerdana9)) { Border = 0, HorizontalAlignment = 0, PaddingBottom = -2f });
                    pdfTableSalesInvoiceHeader.AddCell(new PdfPCell(new Phrase(deserializedSales.TIN, fontVerdana9)) { Border = 0, HorizontalAlignment = 2, PaddingBottom = -2f });

                    pdfTableSalesInvoiceHeader.WriteSelectedRows(0, -1, document.LeftMargin, writer.PageSize.GetTop(document.TopMargin - 60f), writer.DirectContent);

                    // Footer - VAT Analysis
                    PdfPTable pdfTableSalesInvoiceFooter = new PdfPTable(3);
                    pdfTableSalesInvoiceFooter.DefaultCell.Border = 0;
                    pdfTableSalesInvoiceFooter.TotalWidth = document.PageSize.Width - document.LeftMargin - document.RightMargin;
                    pdfTableSalesInvoiceFooter.LockedWidth = true;
                    pdfTableSalesInvoiceFooter.SetWidths(new float[] { 70f, 40f, 10f });

                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(VATAnalysis.VATSales.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2, Rowspan = 2 });
                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(VATAnalysis.TotalSales.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(" ", fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(VATAnalysis.VAT.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(" ", fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });

                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(VATAnalysis.VATExemptSales.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(VATAnalysis.VATSales.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(" ", fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });

                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase("0.00", fontVerdana9)) { Border = 0, HorizontalAlignment = 2, Rowspan = 2 });
                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase("0.00", fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(" ", fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(VATAnalysis.VATSales.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(" ", fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });

                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(VATAnalysis.VAT.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(VATAnalysis.VAT.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(" ", fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });

                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase("", fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(VATAnalysis.TotalSales.ToString("#,##0.00"), fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });
                    pdfTableSalesInvoiceFooter.AddCell(new PdfPCell(new Phrase(" ", fontVerdana9)) { Border = 0, HorizontalAlignment = 2 });

                    pdfTableSalesInvoiceFooter.WriteSelectedRows(0, -1, document.LeftMargin, writer.PageSize.GetBottom(document.BottomMargin + 135f), writer.DirectContent);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}

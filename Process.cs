using Invoicer.Models;
using Invoicer.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using iTextSharp.text.html.simpleparser;
using System.IO;
using iTextSharp.text;

namespace Invoicer.GUI
{
    public class Process
    {
        //public void Go(DataTable dt,string BillNumber)
        //{
        //    //new InvoicerApi(SizeOption.A4, OrientationOption.Landscape, "RS.", BillNumber)
        //    //    .TextColor("#CC0000")
        //    //    .BackColor("#FFD6CC")
        //    //    .Image(@"..\..\image\sss_logo.png", 125, 27)
        //    //    .Company(Address.Make("FROM", new string[] { "SSS AGENCY", "MANGALAM", "PERAMBALUR", "PH : 99947 05803", "GSTIN : 33CXJPS0061H1ZJ" }, "GSTIN : 33CXJPS0061H1ZJ", ""))
        //    //    .Client(Address.Make("BILLING TO", GetCustomerName(dt)))
        //    //    .Items(GetItemRows(dt))
        //    //    .Totals(GetTotalRow(dt))
        //    //    .Details(new List<DetailRow> {
        //    //        DetailRow.Make("PAYMENT INFORMATION", "Make all cheques payable to SSS Agency.", "", "If you have any questions concerning this invoice.", "", "Thank you for your business.")
        //    //    })
        //    //    .Footer("SSS Agency")
        //    //    .Save();
        //}

        private Decimal GetTotalRow(DataTable dt)
        {
           // List<TotalRow> TotalRowList = new List<TotalRow>();
            Decimal subTotal = 0;
            //decimal Gst = 0;
            
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                subTotal += Convert.ToDecimal(dt.Rows[i]["F24"]);
            }
            //Gst = ((subTotal * 10) / 100);
            //TotalRow totalrow = new TotalRow();
            //totalrow.Name = "Sub Total";            
            //totalrow.Value = subTotal;
            //TotalRowList.Add(totalrow);

            //TotalRow totalrow1 = new TotalRow();            
            //totalrow1.Name = "GST 10%";
            //totalrow1.Value = Convert.ToDecimal(Gst);
            //TotalRowList.Add(totalrow1);

            //TotalRow totalrow2 = new TotalRow();
            //totalrow2.Name = "Total";
            //totalrow2.Value = subTotal + Convert.ToDecimal(Gst);
            //TotalRowList.Add(totalrow2);
            return subTotal;           
        }

        private string[] GetCustomerName(DataTable dt)
        {
            return new string[] { Convert.ToString(dt.Rows[0]["F6"]), "GSTIN/UIN :" + Convert.ToString(dt.Rows[0]["F7"]) };
        }

        private List<ItemRow> GetItemRows(DataTable dt)
        {
            List<ItemRow> ItemRowList = new List<ItemRow>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                ItemRow itemrow = new ItemRow();
                itemrow.Name = Convert.ToString(dt.Rows[i]["F26"]);
                itemrow.Description = string.Empty;//dt.Rows[i]["F26"].ToString();
                itemrow.Amount = Convert.ToDecimal(dt.Rows[i]["F9"]);
                itemrow.Price = Convert.ToDecimal(dt.Rows[i]["F9"]);
                itemrow.VAT = Convert.ToDecimal(dt.Rows[i]["F21"]);
                itemrow.Total = Convert.ToDecimal(dt.Rows[i]["F24"]);
                ItemRowList.Add(itemrow);
            }
            return ItemRowList;
        }

        public void Go(DataTable dt, string BillNumber)
        {
            HtmlToPdf(dt, BillNumber);
        }
        public void HtmlToPdf(DataTable dt, string BillNumber)
        {

            StringBuilder sb = new StringBuilder();
            sb.Append("<header>");
            sb.Append("<table border='1' style='font-weight:normal;font-size:10px;font-family:Times New Roman;'>");
            sb.Append("<tr>");
            sb.Append("<td colspan='2'><table border='0'>");
            sb.Append("<tr>");
            sb.Append("<td style='text-align:left;'>");
            sb.Append("GSTIN : 33CXJP0061H1ZJ");
            sb.Append("</td>");
            sb.Append("<td style='text-align:right;'>");
            sb.Append("PH : 99431 17876");
            sb.Append("</td>");
            sb.Append("</tr></table></td></tr>");
            sb.Append("<tr>");
            sb.Append("<td colspan='2' style='text-align:center;'>");
            sb.Append("<h3>SSS AENCY</h3></br> ");
            sb.Append("Mangalam <\br> ");
            sb.Append("<h6>CASH BILL</h6><\br> ");
            sb.Append("</td>");
            sb.Append("</tr>");
            sb.Append("<tr>");
            sb.Append("<td>");
            sb.Append("<table border='0'><tr><td>");
            sb.Append("To. " + Convert.ToString(dt.Rows[0]["F6"]));
            sb.Append("</td></tr>");           
            sb.Append("</table>");
            sb.Append("</td>");
            sb.Append("<td>");
            sb.Append("<table border='0'><tr><td>");
            sb.Append("Payment Terms : Cash ");
            sb.Append("</td></tr>");
            sb.Append("<tr><td>");
            sb.Append("Bill No : 41");
            sb.Append("</td></tr>");
            sb.Append("<tr><td>");
            sb.Append("Date : 19/09/2018");
            sb.Append("</td></tr></table>");
            sb.Append("</td>");
            sb.Append("</tr>");
            sb.Append("</table>");
            sb.Append("</header>");
            sb.Append("<main>");
            sb.Append("<table border='1' style='font-weight:normal;font-size:10px;font-family:Times New Roman;'>");
            sb.Append("<thead>");
            sb.Append("<tr>");
            sb.Append("<th>S.No</th>");
            sb.Append("<th>Item Name</th>");
            sb.Append("<th>HSN Code</th>");
            sb.Append("<th>Qty</th>");
            sb.Append("<th>Rate</th>");
            sb.Append("<th>Tax %</th>");
            sb.Append("<th>Total Amt</th>");
            sb.Append("</tr>");
            sb.Append("</thead>");
            sb.Append("<tbody>");
            sb.Append(GetItembody(dt));
            sb.Append("<tr>");
            sb.Append("<td colspan='7' style='text-align:right;'>Gross Total :" + GetTotalRow(dt) + "</td>");
            sb.Append("</tr>");
            sb.Append("</tbody>");
            sb.Append("</table>");
            sb.Append("</main>");
            //sb.Append("<footer>");
            //sb.Append("For SSS AGENCY.");
            //sb.Append("</footer>");
            StringReader sr = new StringReader(sb.ToString());
            string path = "C:\\Users\\user\\Desktop\\test.pdf";
            StringWriter sw = new StringWriter();
            var output = new FileStream(path, FileMode.Create);
            iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document(PageSize.A4, 10F, 10F, 100F, 0F);
            HTMLWorker htmlparser = new HTMLWorker(pdfDoc);
            iTextSharp.text.pdf.PdfWriter.GetInstance(pdfDoc, output);
            pdfDoc.Open();
            htmlparser.Parse(sr);
            pdfDoc.Close();
        }

        private string GetItembody(DataTable dt)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                sb.Append("<tr>");
                sb.Append("<td>"+i+"</td>");
                sb.Append("<td>"+Convert.ToString(dt.Rows[i]["F26"])+"</td>");
                sb.Append("<td>1511</td>");
                sb.Append("<td>"+Convert.ToDecimal(dt.Rows[i]["F9"])+"</td>");
                sb.Append("<td>"+Convert.ToDecimal(dt.Rows[i]["F9"])+"</td>");
                sb.Append("<td>5</td>");
                sb.Append("<td>"+Convert.ToDecimal(dt.Rows[i]["F24"])+"</td>");
                sb.Append("</tr>");                
            }            
            return sb.ToString();
        }
    }
}

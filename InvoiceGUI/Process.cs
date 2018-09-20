using Invoicer.Models;
using Invoicer.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace Invoicer.GUI
{
    public class Process
    {
        public void Go(DataTable dt,string BillNumber)
        {
            new InvoicerApi(SizeOption.A4, OrientationOption.Landscape, "RS.", BillNumber)
                .TextColor("#CC0000")
                .BackColor("#FFD6CC")
                .Image(@"..\..\image\sss_logo.png", 125, 27)
                .Company(Address.Make("FROM", new string[] { "SSS AGENCY", "MANGALAM", "PERAMBALUR", "PH : 99947 05803", "GSTIN : 33CXJPS0061H1ZJ" }, "GSTIN : 33CXJPS0061H1ZJ", ""))
                .Client(Address.Make("BILLING TO", GetCustomerName(dt)))
                .Items(GetItemRows(dt))
                .Totals(GetTotalRow(dt))
                .Details(new List<DetailRow> {
                    DetailRow.Make("PAYMENT INFORMATION", "Make all cheques payable to SSS Agency.", "", "If you have any questions concerning this invoice.", "", "Thank you for your business.")
                })
                .Footer("SSS Agency")
                .Save();
        }

        private List<TotalRow> GetTotalRow(DataTable dt)
        {
            List<TotalRow> TotalRowList = new List<TotalRow>();
            Decimal subTotal = 0;
            decimal Gst = 0;
            
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                subTotal += Convert.ToDecimal(dt.Rows[i]["F24"]);
            }
            Gst = ((subTotal * 10) / 100);
            TotalRow totalrow = new TotalRow();
            totalrow.Name = "Sub Total";            
            totalrow.Value = subTotal;
            TotalRowList.Add(totalrow);

            TotalRow totalrow1 = new TotalRow();            
            totalrow1.Name = "GST 10%";
            totalrow1.Value = Convert.ToDecimal(Gst);
            TotalRowList.Add(totalrow1);

            TotalRow totalrow2 = new TotalRow();
            totalrow2.Name = "Total";
            totalrow2.Value = subTotal + Convert.ToDecimal(Gst);
            TotalRowList.Add(totalrow2);
            return TotalRowList;           
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

    }
}

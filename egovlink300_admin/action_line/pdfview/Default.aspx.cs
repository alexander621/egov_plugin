using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using iTextSharp;
using iTextSharp.text;
using System.IO;


namespace PDFFormDemo
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {


        }

        protected void btnPrint_Click(object sender, EventArgs e)
        {


            string fullName = this.fullname.Text;
            bool rcvIncome = this.chkIncome.Checked;
            bool youIncome = this.chkYouIncome.Checked;
            bool corpRelation = this.chkCorpRelation.Checked;
            bool isUpdate = this.chkIsUpdate.Checked;
            string relationDesc = this.txtRelation.Text;
            string govtGuy = this.officer.Text;

            try
            {
                Document document = null;


             
                string path = Server.MapPath("~/CIQ.pdf");
                iTextSharp.text.pdf.PdfReader reader =new iTextSharp.text.pdf.PdfReader(path);
                MemoryStream ms = new MemoryStream();
                Rectangle psize = reader.GetPageSize(1);
                float width = psize.Width;
                float height = psize.Height;
                document = new Document(new Rectangle(width, height));
                iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, ms);
                document.Open();

                iTextSharp.text.pdf.PdfContentByte cb = writer.DirectContent;
                iTextSharp.text.pdf.PdfImportedPage page;
                page = writer.GetImportedPage(reader, 1);
                iTextSharp.text.pdf.BaseFont bf = iTextSharp.text.pdf.BaseFont.CreateFont(iTextSharp.text.pdf.BaseFont.HELVETICA, iTextSharp.text.pdf.BaseFont.CP1252, iTextSharp.text.pdf.BaseFont.NOT_EMBEDDED);

                cb.AddTemplate(page, 0, 0);
                iTextSharp.text.pdf.PdfTemplate tp = cb.CreateTemplate(width, height);

                tp.BeginText();
                tp.SetFontAndSize(bf, 10);
                tp.SetTextRenderingMode(iTextSharp.text.pdf.PdfContentByte.TEXT_RENDER_MODE_FILL);
                tp.MoveText(70, 540);
                tp.ShowText(fullName );

                tp.MoveText(-10, -40);
                if (isUpdate )
                {
                    tp.ShowText("X");
                }

                tp.MoveText(120,-80);
                tp.ShowText(govtGuy);

                tp.MoveText(-55, -120);
                if (rcvIncome)
                {
                    tp.ShowText("X");
                }
                tp.MoveText(0, -70);

                if(youIncome )
                {
                    tp.ShowText("X");
                }
                
                tp.MoveText(0, -60);
                if (corpRelation)
                    tp.ShowText("X");

                tp.MoveText(0, -40);

                if (relationDesc != "")
                {
                    tp.ShowText(relationDesc);
                }

                tp.MoveText(0,-65);
                tp.ShowText(fullName);
                tp.MoveText(260,0);
                tp.ShowText(DateTime.Now.ToShortDateString());

                cb.AddTemplate(tp, 0, 0);
                tp.EndText();
                cb.AddTemplate(tp, 0, 0);
                iTextSharp.text.pdf.Barcode39 code39 = new iTextSharp.text.pdf.Barcode39();
                code39.Code = "1992030495";
                code39.AltText = "";
                iTextSharp.text.pdf.PdfTemplate c39 = code39.CreateTemplateWithBarcode(cb, iTextSharp.text.BaseColor.BLACK, iTextSharp.text.BaseColor.BLACK);
                cb.AddTemplate(c39, 414, 275);
              
                document.Close();
                Response.Clear();
                Response.ClearContent();
                Response.ClearHeaders();
                Response.ContentType = "application/pdf";
                byte[] b = ms.ToArray();
                int len = b.Length;
                Response.OutputStream.Write(b, 0, len);
                Response.OutputStream.Flush();
                Response.OutputStream.Close();
                ms.Close();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message + ex.StackTrace);

            }



        }
    }
}
 
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Data;
using System.Data.SqlClient;

public partial class action_line_pdfview_Work_Order2 : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

        string path = Server.MapPath( "/public_documents300/eclink/unpublished_documents/PDFs/FarmingtonEmpApp.pdf" );
        //Response.Write( path );
        //Response.End( );

        PdfReader reader = new PdfReader( path );  // We crash here with "not found as file or resource." error
        MemoryStream ms = new MemoryStream( );
        PdfStamper stamper = new PdfStamper( reader, ms );
        AcroFields fields = stamper.AcroFields;

        fields.SetField( "userfname", "Steve" );

        stamper.Writer.CloseStream = false;
        stamper.FormFlattening = true;
        stamper.Close( );
        //reader.Close( );

        Response.Clear( );
        Response.ClearContent( );
        Response.ClearHeaders( );
        Response.ContentType = "application/pdf";
        byte[] b = ms.ToArray( );
        int len = b.Length;
        Response.OutputStream.Write( b, 0, len );
        Response.OutputStream.Flush( );
        Response.OutputStream.Close( );
        ms.Close( );
        reader.Close( );
    }
}
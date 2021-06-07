using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using iTextSharp;
using iTextSharp.text;
using System.IO;
using System.Data;
using System.Data.SqlClient;


public partial class recreation_display_waiver : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string waiverList = "0";
        string waiverText;
        if (Request["MASK"] != null)
        {
            // this is the list of waivers to display
            waiverList = Request["MASK"].ToString( );
        }

        // pull the wiavers here
        waiverText = getWaiverText( waiverList );

        if (waiverText != "")
        {
            // get the replacement values
            string amount = "";
            if (Request["amount"] != null)
                amount = Request["amount"].ToString( );
            waiverText = waiverText.Replace( "[*amount*]", amount );


            string checkInDate = "";
            if (Request["checkindate"] != null)
                checkInDate = Request["checkindate"].ToString( );
            waiverText = waiverText.Replace( "[*checkindate*]", checkInDate );

            string checkInTime = "";
            if (Request["checkintime"] != null)
                checkInTime = Request["checkintime"].ToString( );
            waiverText = waiverText.Replace( "[*checkintime*]", checkInTime );

            string checkOutDate = "";
            if (Request["checkoutdate"] != null)
                checkOutDate = Request["checkoutdate"].ToString( );
            waiverText = waiverText.Replace( "[*checkoutdate*]", checkOutDate );

            string checkOutTime = "";
            if (Request["checkouttime"] != null)
                checkOutTime = Request["checkouttime"].ToString( );
            waiverText = waiverText.Replace( "[*checkouttime*]", checkOutTime );

            string pointOfContact = "";
            if (Request["custom_poc"] != null)
                pointOfContact = Request["custom_poc"].ToString( );
            waiverText = waiverText.Replace( "[*pointofcontact*]", pointOfContact );

            string attending = "";
            if (Request["custom_attending"] != null)
                attending = Request["custom_attending"].ToString( );
            waiverText = waiverText.Replace( "[*attending*]", attending );

            string organization = "";
            if (Request["custom_org"] != null)
                organization = Request["custom_org"].ToString( );
            waiverText = waiverText.Replace( "[*organization*]", organization );

            int userId = 0;
            if (Request["iuserid"] != null)
            {
                userId = int.Parse( Request["iuserid"].ToString( ) );
                waiverText = getRenterInformation( userId, waiverText );
            }
            else
            {
                waiverText = waiverText.Replace( "[*firstname*]", "" );
                waiverText = waiverText.Replace( "[*middle*]", "" );
                waiverText = waiverText.Replace( "[*lastname*]", "" );
                waiverText = waiverText.Replace( "[*address1*]", "" );
                waiverText = waiverText.Replace( "[*address2*]", "" );
                waiverText = waiverText.Replace( "[*city*]", "" );
                waiverText = waiverText.Replace( "[*state*]", "" );
                waiverText = waiverText.Replace( "[*zip*]", "" );
                waiverText = waiverText.Replace( "[*email*]", "" );
            }

            // replace reservation details

            int startingXPosition = 20;
            int startingYPosition = 790;
            float defaultFontSize = 10;
            Document document = null;
            MemoryStream ms = new MemoryStream( );
            document = new Document( PageSize.A4 );
            iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance( document, ms );
            document.Open( );

            iTextSharp.text.pdf.PdfContentByte cb = writer.DirectContent;
            iTextSharp.text.pdf.BaseFont bf = iTextSharp.text.pdf.BaseFont.CreateFont( iTextSharp.text.pdf.BaseFont.HELVETICA, iTextSharp.text.pdf.BaseFont.CP1252, iTextSharp.text.pdf.BaseFont.NOT_EMBEDDED );
            iTextSharp.text.pdf.PdfTemplate tp = cb.CreateTemplate( 600f, 792f );

            tp.BeginText( );
            tp.SetFontAndSize( bf, defaultFontSize );
            tp.SetTextRenderingMode( iTextSharp.text.pdf.PdfContentByte.TEXT_RENDER_MODE_FILL );
            tp.MoveText( startingXPosition, startingYPosition ); // I am not sure that this does anything that ends up in the generated PDF.
            Font waiverFont = FontFactory.GetFont( "HELVETICA", 9f );

            // slpit the waiver text into an array
            string[] splitWaivers = waiverText.Split( new string[] { "[*NEWPAGE*]" }, StringSplitOptions.RemoveEmptyEntries );
            int pageCount = 1;
            Chunk waiverChunk;
            Phrase waiverPhrase;
            Paragraph waiverP;

            // loop on the array
            foreach (string waiver in splitWaivers)
            {
                if ( pageCount > 1 )
                    document.NewPage( );
                pageCount++;
                waiverChunk = new Chunk( waiver, waiverFont );
                waiverPhrase = new Phrase( waiverChunk );
                waiverP = new Paragraph( );
                waiverP.SetLeading( 0.0f, 1.0f );
                waiverP.Add( waiverPhrase );
                document.Add( waiverP );
            }

            // close up here
            cb.AddTemplate( tp, 0, 0 );
            tp.EndText( );
            document.Close( );

            // stream out the PDF here
            Response.Clear( );
            Response.ClearContent( );
            Response.ClearHeaders( );
            Response.ContentType = "application/pdf";
            Response.AppendHeader( "Content-Disposition", "filename=reservation_forms.pdf" );
            byte[] b = ms.ToArray( );
            int len = b.Length;
            Response.OutputStream.Write( b, 0, len );
            Response.OutputStream.Flush( );
            Response.OutputStream.Close( );
            ms.Close( );
        }
    }

    public string getWaiverText( string _WaiverList )
    {
        string waiverText = "";

        string sql = "SELECT body FROM egov_waivers where waiverid IN ( " + _WaiverList + " )";

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        while (myReader.Read( ))
        {
            waiverText += myReader["body"].ToString( );
        }

        myReader.Close( );
        myReader.Dispose( );
        sqlConn.Close( );
        sqlConn.Dispose( );

        return waiverText;
    }

    public string getRenterInformation( int _UserId, string _WaiverText )
    {
        // a lot of this is not used, but this is the query from the old asp script
        string sql = "SELECT ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, ISNULL(useraddress,'') AS useraddress, ";
	    sql += "ISNULL(usercity,'') AS usercity, ISNULL(userstate,'') AS userstate, ISNULL(userzip,'') AS userzip, ";
	    sql += "ISNULL(useremail,'') AS useremail, ISNULL(userhomephone,'') AS userhomephone, ISNULL(userbusinessname,'') AS userbusinessname, ";
	    sql += "ISNULL(userworkphone,'') AS userworkphone, ISNULL(userfax, '') AS userfax FROM egov_users WHERE userid = " + _UserId.ToString();

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            myReader.Read( );
            _WaiverText = _WaiverText.Replace( "[*firstname*]", myReader["userfname"].ToString( ) );
            _WaiverText = _WaiverText.Replace( "[*middle*]", "" );
            _WaiverText = _WaiverText.Replace( "[*lastname*]", myReader["userlname"].ToString( ) );
            _WaiverText = _WaiverText.Replace( "[*address1*]", myReader["useraddress"].ToString( ) );
            _WaiverText = _WaiverText.Replace( "[*address2*]", "" );
            _WaiverText = _WaiverText.Replace( "[*city*]", myReader["usercity"].ToString( ) );
            _WaiverText = _WaiverText.Replace( "[*state*]", myReader["userstate"].ToString( ) );
            _WaiverText = _WaiverText.Replace( "[*zip*]", myReader["userzip"].ToString( ) );
            _WaiverText = _WaiverText.Replace( "[*email*]", myReader["useremail"].ToString( ) );
        }

        myReader.Close( );
        myReader.Dispose( );
        sqlConn.Close( );
        sqlConn.Dispose( );

        return _WaiverText;
    }

}
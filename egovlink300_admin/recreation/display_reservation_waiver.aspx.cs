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
        int reservationId = 0;

        if (Request["MASK"] != null)
        {
            // this is the list of waivers to display
            waiverList = Request["MASK"].ToString( );
        }

        // pull the waivers here
        waiverText = getWaiverText( waiverList );

        reservationId = int.Parse( Request["reservationid"].ToString( ) );

        if (waiverText != "")
        {
            int userId = 0;

            // get the replacement values from the database tables
            waiverText = getReservationDetails( reservationId, waiverText, out userId );

            waiverText = getOtherStuff( reservationId, waiverText );

            waiverText = getRenterInformation( userId, waiverText );

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
        else
        {
            _WaiverText = _WaiverText.Replace( "[*firstname*]", "" );
            _WaiverText = _WaiverText.Replace( "[*middle*]", "" );
            _WaiverText = _WaiverText.Replace( "[*lastname*]", "" );
            _WaiverText = _WaiverText.Replace( "[*address1*]", "" );
            _WaiverText = _WaiverText.Replace( "[*address2*]", "" );
            _WaiverText = _WaiverText.Replace( "[*city*]", "" );
            _WaiverText = _WaiverText.Replace( "[*state*]", "" );
            _WaiverText = _WaiverText.Replace( "[*zip*]", "" );
            _WaiverText = _WaiverText.Replace( "[*email*]", "" );
        }

        myReader.Close( );
        myReader.Dispose( );
        sqlConn.Close( );
        sqlConn.Dispose( );

        return _WaiverText;
    }

    public string getReservationDetails( int _ReservationId, string _WaiverText, out int _LesseeId )
    {
        _LesseeId = 0;

        string sql = "SELECT ISNULL(amount,0.00) AS amount, checkindate, checkintime, checkoutdate, checkouttime, lesseeid FROM egov_facilityschedule ";
        sql += "INNER JOIN egov_facility ON egov_facilityschedule.facilityid = egov_facility.facilityid WHERE facilityscheduleid = " + _ReservationId.ToString( );

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            myReader.Read( );
            _WaiverText = _WaiverText.Replace( "[*amount*]", myReader["amount"].ToString( ) );
            _WaiverText = _WaiverText.Replace( "[*checkindate*]", myReader["checkindate"].ToString( ) );
            _WaiverText = _WaiverText.Replace( "[*checkintime*]", myReader["checkintime"].ToString( ) );
            _WaiverText = _WaiverText.Replace( "[*checkoutdate*]", myReader["checkoutdate"].ToString( ) );
            _WaiverText = _WaiverText.Replace( "[*checkouttime*]", myReader["checkouttime"].ToString( ) );
            _LesseeId = int.Parse( myReader["lesseeid"].ToString( ) );
        }
        else
        {
            _WaiverText = _WaiverText.Replace( "[*amount*]", "" );
            _WaiverText = _WaiverText.Replace( "[*checkindate*]", "" );
            _WaiverText = _WaiverText.Replace( "[*checkintime*]", "" );
            _WaiverText = _WaiverText.Replace( "[*checkoutdate*]", "" );
            _WaiverText = _WaiverText.Replace( "[*checkouttime*]", "" );
        }


        myReader.Close( );
        myReader.Dispose( );
        sqlConn.Close( );
        sqlConn.Dispose( );

        return _WaiverText;
    }

    public string getOtherStuff( int _ReservationId, string _WaiverText )
    {
        string sql = "SELECT ISNULL(F.fieldname,'') AS fieldname, ISNULL(V.fieldvalue,'') AS fieldvalue FROM egov_facility_field_values V ";
        sql += "INNER JOIN egov_facility_fields F ON V.fieldid = F.fieldid WHERE V.paymentid = " + _ReservationId.ToString( ) + " ORDER BY V.paymentid, V.fieldid";

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            while (myReader.Read( ))
            {
                switch (myReader["fieldname"].ToString( ))
                {
                    case "poc":
                        _WaiverText = _WaiverText.Replace( "[*pointofcontact*]", myReader["fieldvalue"].ToString( ) );
                        break;
                    case "organization":
                        _WaiverText = _WaiverText.Replace( "[*organization*]", myReader["fieldvalue"].ToString( ) );
                        break;
                    case "purpose":
                        _WaiverText = _WaiverText.Replace( "[*purpose*]", myReader["fieldvalue"].ToString( ) );
                        break;
                    case "attending":
                        _WaiverText = _WaiverText.Replace( "[*attending*]", myReader["fieldvalue"].ToString( ) );
                        break;
                }
            }
        }
        else
        {
            _WaiverText = _WaiverText.Replace( "[*pointofcontact*]", "" );
            _WaiverText = _WaiverText.Replace( "[*organization*]", "" );
            _WaiverText = _WaiverText.Replace( "[*purpose*]", "" );
            _WaiverText = _WaiverText.Replace( "[*attending*]", "" );
        }

        myReader.Close( );
        myReader.Dispose( );
        sqlConn.Close( );
        sqlConn.Dispose( );

        return _WaiverText;
    }



}
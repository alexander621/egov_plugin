using System;
using System.Collections.Generic;
using System.Configuration;
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

public partial class action_line_pdfview_work_orderh : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        int requestId = 0;
        string trackingNumber = "";
        string submittedDateTime;
        string debugline;
        string createdBy;
        string formTitle;
        string firstName;
        string lastName;
        string businessName;
        string citizenEmail;
        string dayPhone;
        string fax;
        string citizenAddress;
        string citizenCity;
        string citizenState;
        string citizenZip;
        string citizenCountry;
        string contactInformation = "";
        string issueLocation = "";
        Paragraph bodyText;
        string questionsAndAnswers;
        string hideActivityLog = "N";
        string activityLog;
        Boolean isCondensed = false;
        string signatureLines = "";
	string IssueLocationOn = "";

	try
	{
		string uid = Request.Cookies["User"]["UserID"];
	}
	catch
	{
               Response.Redirect( "../../login.asp" );    
	}


        if (Request["iRequestID"] != null)
        {
            if ( common.IsNumeric( Request[ "iRequestID" ].ToString( ) ) == true )
                requestId = int.Parse( Request["iRequestID"].ToString( ) );
            else
                Response.Redirect( "../../login.asp" );    
        }

        if (Request["hideActLog"] != null)
            hideActivityLog = Request["hideActLog"].ToString( ).ToUpper( );

        if (Request["pdfaction"] != null)
        {
            if (Request["pdfaction"].ToString( ) == "WORKORDER_CONDENSED")
                isCondensed = true;
        }

        // pull request information here
        trackingNumber = getGeneralRequestInfo( requestId.ToString( ), out submittedDateTime, out createdBy, out formTitle, out firstName, out lastName, out businessName, out citizenEmail,
                        out dayPhone, out fax, out citizenAddress, out citizenCity, out citizenState, out citizenZip, out citizenCountry, out IssueLocationOn );
//        Response.Write( "Tracking Number: " + trackingNumber );
//        Response.End( );

        MemoryStream ms = new MemoryStream( );
        Document document = new Document( PageSize.LETTER );
        PdfWriter writer = PdfWriter.GetInstance( document, ms );

        // this is the repeating header on each page
        MyPageEventHandlerWOH eh = new MyPageEventHandlerWOH( );
        eh.formTitle = formTitle;
        eh.trackingNumber = trackingNumber;
        eh.receivedDate = submittedDateTime;
        eh.createdBy = createdBy;
        writer.PageEvent = eh;

        document.SetMargins( document.LeftMargin, document.RightMargin, document.TopMargin + 95f, document.BottomMargin );
        document.Open( );

        Font arial = FontFactory.GetFont( "arial", 10f, Font.NORMAL );

        // display the contact information
        document.Add( createSubTitle( "Contact Information", document ) );
        if (isCondensed)
        {
            if ( firstName != "" || lastName != "" )
                contactInformation = firstName + " " + lastName + "\n";
            if (businessName != "")
                contactInformation += businessName + "\n";
            if (citizenEmail != "")
                contactInformation += citizenEmail + "\n";
            if (dayPhone != "")
                contactInformation += dayPhone + "\n";
            if (fax != "")
                contactInformation += fax + "\n";
            if (citizenAddress != "")
                contactInformation += citizenAddress + "\n";
            if (citizenCity != "")
                contactInformation += citizenCity + "/";
            if (citizenState != "")
                contactInformation += citizenState + "/";
            if (citizenZip != "")
                contactInformation += citizenZip;
        }
        else
        {
            contactInformation = "First Name: " + firstName;
            contactInformation += "\nLast Name: " + lastName;
            contactInformation += "\nBusiness Name: " + businessName;
            contactInformation += "\nEmail: " + citizenEmail;
            contactInformation += "\nDaytime Phone: " + dayPhone;
            contactInformation += "\nFax: " + fax;
            contactInformation += "\nAddress: " + citizenAddress;
            contactInformation += "\nCity: " + citizenCity;
            contactInformation += "\nState: " + citizenState;
            contactInformation += "\nZip: " + citizenZip;
            contactInformation += "\nCountry: " + citizenCountry;
        }
        bodyText = new Paragraph( contactInformation, arial );
        bodyText.SpacingBefore = 1f;
        bodyText.SpacingAfter = 1f;
        document.Add( bodyText );

        // display the issue location information
        if (IssueLocationOn != "False")
        {
		var subTitle = createSubTitle( "Issue Location", document );
		document.Add( subTitle );
	}
        issueLocation = getIssueLocationDetails( requestId.ToString( ), isCondensed );
        bodyText = new Paragraph( issueLocation, arial );
        bodyText.SpacingBefore = 1f;
        bodyText.SpacingAfter = 1f;
        if (IssueLocationOn != "False")
        {
        	document.Add( bodyText );
	}

        // display the request details (questions and answers)
        document.Add( createSubTitle( "Request Details", document ) );
        questionsAndAnswers = getQuestionsAndAnswers( requestId.ToString( ) );
        bodyText = new Paragraph( questionsAndAnswers, arial );
        bodyText.SpacingBefore = 1f;
        bodyText.SpacingAfter = 1f;
        document.Add( bodyText );

        // display the Request Activity
        if (hideActivityLog == "N")
        {
            document.Add( createSubTitle( "Request Activity", document ) );
            activityLog = getActivityLog( requestId.ToString( ), submittedDateTime, createdBy );
            bodyText = new Paragraph( activityLog, arial );
            bodyText.SpacingBefore = 1f;
            bodyText.SpacingAfter = 1f;
            document.Add( bodyText );
        }


        // here we want to add a page break and underscore lines for the regular work orders
        if (isCondensed == false)
        {
            document.NewPage( );
            string commentLine = new string( '_', 97 );
            string commentBlock = "Comments: ";

            for (int x = 0; x < 15; x++)
            {
                commentBlock += "\n" + commentLine;
            }

            bodyText = new Paragraph( commentBlock, arial );
            bodyText.SpacingBefore = 18f;
            bodyText.SpacingAfter = 10f;
            document.Add( bodyText );
        }

        // here we add the date and signature lines
        signatureLines = "__________________                                                                                    ____________________________________";
        signatureLines += "\n                            Date                                                                                                                           Authorized Signature";
        bodyText = new Paragraph( signatureLines, arial );
        bodyText.SpacingBefore = 18f;
        bodyText.SpacingAfter = 1f;
        document.Add( bodyText );

        // clean up and stream the PDF to the browser
        document.Close( );
        Response.Clear( );
        Response.ClearContent( );
        Response.ClearHeaders( );
        Response.ContentType = "application/pdf";
        Response.AppendHeader( "Content-Disposition", "filename=workorder.pdf" );
        byte[] b = ms.ToArray( );
        int len = b.Length;
        Response.OutputStream.Write( b, 0, len );
        Response.OutputStream.Flush( );
        Response.OutputStream.Close( );
        ms.Close( );
    }


    public string getActivityLog( string _RequestId, string _SubmittedDateTime, string _CreatedBy )
    {
        string activityLog = "";
        DateTime editDate;

        string sql = "SELECT action_editdate, ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, ISNULL(action_status,'') AS action_status, ";
        sql += "ISNULL(action_citizen,'') AS action_citizen, ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, ";
        sql += "ISNULL(action_internalcomment,'') AS action_internalcomment, ISNULL(action_externalcomment,'') AS action_externalcomment ";
        sql += "FROM egov_action_responses ";
        sql += "LEFT OUTER JOIN egov_users ON egov_action_responses.action_userid = egov_users.userid ";
        sql += "LEFT OUTER JOIN users ON egov_action_responses.action_userid = users.userid ";
        sql += "WHERE action_autoid = " + _RequestId;
        sql += " ORDER BY action_editdate DESC";

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            while (myReader.Read( ))
            {
                editDate = DateTime.Parse( myReader["action_editdate"].ToString( ) );
                if (activityLog != "")
                    activityLog += "\n";

                activityLog += editDate.ToString( "M/dd/yyyy h:mm tt" ) + " -- " + myReader["firstname"].ToString( ) + " " + myReader["lastname"].ToString( ) + " - " + myReader["action_status"].ToString( ).ToUpper( );

                if (myReader["action_externalcomment"].ToString( ) != "")
                    activityLog += "\n-----Note to Citizen: " + cleanForPDF( myReader["action_externalcomment"].ToString( ) );

                if (myReader["action_citizen"].ToString( ) != "")
                    activityLog += "\n" + editDate.ToString( "M/dd/yyyy h:mm tt" ) + " -- " + myReader["userfname"].ToString( ) + " " + myReader["userlname"].ToString( ) + " - " + myReader["action_citizen"].ToString( );

                if (myReader["action_internalcomment"].ToString( ) != "")
                    activityLog += "\n-----Internal Note: " + cleanForPDF( myReader["action_internalcomment"].ToString( ) );

            }
        }
        else
            activityLog += "\n" + _SubmittedDateTime + " -- No activity Reported.";

        activityLog += "\n" + _SubmittedDateTime + " -- " + _CreatedBy + " - SUBMITTED";

        myReader.Close( );
        myReader.Dispose( );
        sqlConn.Close( );
        sqlConn.Dispose( );

        return activityLog;

    }


    public string getQuestionsAndAnswers( string _RequestId )
    {
        string questionsAndAnswers = "";
        string question;
        string answer;

        string sql = "SELECT F.submitted_request_field_prompt, ISNULL(submitted_request_field_response,'') AS submitted_request_field_response ";
        sql += "FROM egov_submitted_request_fields F ";
        sql += "LEFT OUTER JOIN egov_submitted_request_field_responses R ON F.submitted_request_field_id = R.submitted_request_field_id ";
        sql += "WHERE F.submitted_request_id = " + _RequestId;
        sql += " ORDER BY F.submitted_request_field_sequence";

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        while (myReader.Read( ))
        {
            if (questionsAndAnswers != "")
                questionsAndAnswers += "\n\n";

            question = cleanForPDF( myReader["submitted_request_field_prompt"].ToString( ) );
            if (question != "")
                questionsAndAnswers += question + "\n";

            answer = cleanForPDF( myReader["submitted_request_field_response"].ToString( ) );
            if (answer != "")
                questionsAndAnswers += answer;
        }

        myReader.Close( );

	sql = "SELECT mobileappdescription FROM egov_actionline_requests WHERE action_autoid = " + _RequestId;
        myCommand = new SqlCommand( sql, sqlConn );
        myReader = myCommand.ExecuteReader( );
        while (myReader.Read( ))
        {
            if (questionsAndAnswers != "")
                questionsAndAnswers += "\n\n";

	    //questionsAndAnswers += "Request Description: " + cleanForPDF( myReader["mobileappdescription"].ToString() );;
	    questionsAndAnswers += cleanForPDF( myReader["mobileappdescription"].ToString() );;
        }

        myReader.Close( );

        myReader.Dispose( );
        sqlConn.Close( );
        sqlConn.Dispose( );

        return questionsAndAnswers;

    }


    public string cleanForPDF( string _Value )
    {
        string cleanValue = "";

        cleanValue = _Value;
        cleanValue = cleanValue.Replace( "<p>", "" ).Replace( "</p>", "" ).Replace( "<P>", "" ).Replace( "</P>", "" ).Replace( "</br>", "" ).Replace( "<br>", "" );
        cleanValue = cleanValue.Replace( "<br/>", " " ).Replace( "&lt;br/>", " " ).Replace( "&lt;BR>", " " ).Replace( "</BR>", "" ).Replace( "<BR>", "" ).Replace( "</b>", "" );
        cleanValue = cleanValue.Replace( "<b>", "" ).Replace( "</B>", "" ).Replace( "<B>", "" ).Replace( "<u>", "" ).Replace( "</u>", "" ).Replace( "<U>", "" ).Replace( "</U>", "" );
        cleanValue = cleanValue.Replace( "&quot;", "\"" ).Replace( "\n", "" ).Replace( "default_novalue", "" );

        return cleanValue;
    }


    public string buildStreetAddress( string _StreetNumber, string _Prefix, string _StreetName, string _Suffix, string _Direction )
    {
        string streetAddress = "";

        if ( _StreetNumber != "")
            streetAddress = _StreetNumber;
        if (_Prefix != "")
        {
            if (streetAddress != "")
                streetAddress += " ";
            streetAddress += _Prefix;
        }
        if (_StreetName != "")
        {
            if (streetAddress != "")
                streetAddress += " ";
            streetAddress += _StreetName;
        }
        if (_Suffix != "")
        {
            if (streetAddress != "")
                streetAddress += " ";
            streetAddress += _Suffix;
        }
        if (_Direction != "")
        {
            if (streetAddress != "")
                streetAddress += " ";
            streetAddress += _Direction;
        }

        return streetAddress;
    }


    public string getIssueLocationDetails( string _RequestId, Boolean _IsCondensed )
    {
        string issueLocationDetails = "";
        string street = "";

        string sql = "SELECT ISNULL(streetnumber,'') AS streetnumber, ISNULL(streetprefix,'') AS streetprefix, ISNULL(streetaddress,'') AS streetaddress, ";
        sql += "ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection, ISNULL(streetunit,'') AS streetunit, ";
        sql += "ISNULL(city,'') AS city, ISNULL(state,'') AS state, ISNULL(zip,'') AS zip, ISNULL(comments ,'') AS comments ";
        sql += "FROM egov_action_response_issue_location WHERE actionrequestresponseid = " + _RequestId;

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            myReader.Read( );
            street = buildStreetAddress( myReader["streetnumber"].ToString( ), myReader["streetprefix"].ToString( ), myReader["streetaddress"].ToString( ), myReader["streetsuffix"].ToString( ), myReader["streetdirection"].ToString( ) );
            if (_IsCondensed)
            {
                if (street != "")
                    issueLocationDetails = street + "\n";
                if (myReader["streetunit"].ToString( ) != "")
                    issueLocationDetails += myReader["streetunit"].ToString( ) + "\n";
                if (myReader["city"].ToString( ) != "")
                    issueLocationDetails += myReader["city"].ToString( ) + "/";
                if (myReader["state"].ToString( ) != "")
                    issueLocationDetails += myReader["state"].ToString( ) + "/";
                if (myReader["zip"].ToString( ) != "")
                    issueLocationDetails += myReader["zip"].ToString( ) + "\n";
                if (myReader["comments"].ToString( ) != "")
                    issueLocationDetails += myReader["comments"].ToString( );
            }
            else
            {
                issueLocationDetails = "Street: " + street;
                issueLocationDetails += "\nUnit: " + myReader["streetunit"].ToString( );
                issueLocationDetails += "\nCity: " + myReader["city"].ToString( );
                issueLocationDetails += "\nState: " + myReader["state"].ToString( );
                issueLocationDetails += "\nZip: " + myReader["zip"].ToString( );
                issueLocationDetails += "\nComments: " + myReader["comments"].ToString( );
            }
        }
        else
        {
            issueLocationDetails = "Street:\nUnit:\nCity:\nState:\nZip:\nComments:";
        }

        myReader.Close( );
        myReader.Dispose( );
        sqlConn.Close( );
        sqlConn.Dispose( );

        return issueLocationDetails;

    }


    public string getGeneralRequestInfo( string _RequestId, out string _SubmittedDateTime, out string _CreatedBy, out string _FormTitle, out string _FirstName, out string _LastName, out string _BusinessName, 
            out string _CitizenEmail, out string _DayPhone, out string _Fax, out string _CitizenAddress, out string _CitizenCity, out string _CitizenState, out string _CitizenZip, out string _CitizenCountry, out string _IssueLocationOn  )
    {
        string trackingNumber = "";
        DateTime submitDate;
        _SubmittedDateTime = "";
        _CreatedBy = "";
        _FormTitle = "";
        _FirstName = "";
        _LastName = "";
        _BusinessName = "";
        _CitizenEmail = "";
        _DayPhone = "";
        _Fax = "";
        _CitizenAddress = "";
        _CitizenCity = "";
        _CitizenState = "";
        _CitizenZip = "";
        _CitizenCountry = "";
	_IssueLocationOn = "";

        string sql = "SELECT R.category_title, R.status, R.submit_date, R.comment, R.userid, R.assignedemployeeid, ISNULL(R.contactmethodid,0) AS contactmethodid, ";
        sql += "ISNULL(R.employeesubmitid,0) AS employeesubmitid, (ISNULL(U.FirstName,'') + ' ' + ISNULL(U.LastName,'')) AS EmployeeSubmitName, ";
	    sql += "ISNULL(EU.userfname,'') AS CitizenFirstName, ISNULL(EU.userlname,'') AS CitizenLastName, ISNULL(EU.userbusinessname,'') AS userbusinessname, ";
	    sql += "ISNULL(EU.useremail,'') AS useremail, ISNULL(EU.userhomephone,'') AS userhomephone, ISNULL(EU.userfax,'') AS userfax, ";
	    sql += "ISNULL(EU.useraddress,'') AS useraddress, ISNULL(EU.usercity,'') AS usercity, ISNULL(EU.userstate,'') AS userstate, ";
	    sql += "ISNULL(EU.userzip,'') AS userzip, ISNULL(EU.usercountry,'') AS usercountry, F.DeptID, F.action_form_display_issue ";
	    sql += "FROM egov_actionline_requests R ";
	    sql += "LEFT OUTER JOIN users U ON R.employeesubmitid = U.userid ";
	    sql += "LEFT OUTER JOIN egov_users EU ON R.userid = EU.userid ";
	    sql += "LEFT OUTER JOIN egov_action_request_forms AS F ON R.category_id = F.action_form_id ";
        sql += "WHERE R.action_autoid = " + _RequestId;

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            myReader.Read( );
            submitDate = DateTime.Parse( myReader["submit_date"].ToString( ) );
            trackingNumber = _RequestId + submitDate.ToString( "HHmm" );
            _SubmittedDateTime = submitDate.ToString( "M/dd/yyyy h:mmtt" );
            if (myReader["EmployeeSubmitName"].ToString( ).Trim( ) != "")
                _CreatedBy = myReader["EmployeeSubmitName"].ToString( ) + " (Admin Employee)";
            else
                _CreatedBy = myReader["CitizenFirstName"].ToString( ) + " " + myReader["CitizenLastName"].ToString( ) + " (Citizen)";
            _FormTitle = myReader["category_title"].ToString( );
            _FirstName = myReader["CitizenFirstName"].ToString( );
            _LastName = myReader["CitizenLastName"].ToString( );
            _BusinessName = myReader["userbusinessname"].ToString( );
            _CitizenEmail = myReader["useremail"].ToString( );
            _DayPhone = common.formatPhoneNumber( myReader["userhomephone"].ToString( ) );
            _Fax = common.formatPhoneNumber( myReader["userfax"].ToString( ) );
            _CitizenAddress = myReader["useraddress"].ToString( );
            _CitizenCity = myReader["usercity"].ToString( );
            _CitizenState = myReader["userstate"].ToString( );
            _CitizenZip = myReader["userzip"].ToString( );
            _CitizenCountry = myReader["usercountry"].ToString( );
	    _IssueLocationOn = myReader["action_form_display_issue"].ToString();
        }

        myReader.Close( );
        myReader.Dispose( );
        sqlConn.Close( );
        sqlConn.Dispose( );

        return trackingNumber;
    }


    public Paragraph createSubTitle( string _TitleText, Document _Document )
    {
        Font arialSubTitle = FontFactory.GetFont( "arial", 14f, Font.NORMAL, new BaseColor( 255, 255, 255 ) );
        PdfPTable titleTable = new PdfPTable( 1 );
        titleTable.TotalWidth = _Document.PageSize.Width - (_Document.LeftMargin + _Document.RightMargin);
        titleTable.LockedWidth = true;
        Chunk titleChunk = new Chunk( _TitleText, arialSubTitle );
        Phrase titlePhrase = new Phrase( );
        titlePhrase.Add( titleChunk );
        PdfPCell titleCell = new PdfPCell( titlePhrase );
        titleCell.Border = PdfPCell.NO_BORDER;
        titleCell.VerticalAlignment = Element.ALIGN_MIDDLE;
        titleCell.HorizontalAlignment = Element.ALIGN_LEFT;
        titleCell.BackgroundColor = new BaseColor( 60, 60, 60 );
        titleTable.AddCell( titleCell );
        Paragraph titleParagraph = new Paragraph( );
        titleParagraph.Add( titleTable );
        titleParagraph.SpacingBefore = 4f;

        return titleParagraph;
    }
}

public class MyPageEventHandlerWOH : PdfPageEventHelper
{
    public string formTitle { get; set; }
    public string trackingNumber {get; set;}
    public string receivedDate { get; set; }
    public string createdBy { get; set; }

    public override void OnEndPage( PdfWriter writer, Document document )
    {
        float cellHeight = document.TopMargin;
        Rectangle page = document.PageSize;
        PdfPCell c;

        // the header is a table
        PdfPTable head = new PdfPTable( 1 );
        head.TotalWidth = document.PageSize.Width - (document.LeftMargin + document.RightMargin);
        head.SpacingAfter = 24f;
        head.SpacingBefore = 12f;
     
        // the fonts we are using in the different parts os the header
        Font arialTitle = FontFactory.GetFont( "arial", 16f, Font.NORMAL );
        // this is white text to go on the dark background
        Font arialSubTitle = FontFactory.GetFont( "arial", 14f, Font.NORMAL, new BaseColor( 255, 255, 255 ) );
        Font arialHeader = FontFactory.GetFont( "arial", 10f );

	//Logo
	iTextSharp.text.Image addLogo = default(iTextSharp.text.Image);
 	addLogo = iTextSharp.text.Image.GetInstance(System.Web.HttpContext.Current.Server.MapPath("../../../custom/images/harlingen/citylogo.jpg"));
	addLogo.ScaleToFit(50, 50);

	c = new PdfPCell(addLogo);
	c.Colspan = 2; // either 1 if you need to insert one cell
	c.Border = 0;
	c.HorizontalAlignment = Element.ALIGN_LEFT;
        head.AddCell( c );

        // the main title
        c = new PdfPCell( new Phrase( "Work Order Form", arialTitle ) );
        c.Border = PdfPCell.BOTTOM_BORDER;
        c.VerticalAlignment = Element.ALIGN_BOTTOM;
        c.HorizontalAlignment = Element.ALIGN_CENTER;
        //c.FixedHeight = cellHeight;
        head.AddCell( c );

        // the form name
        c = new PdfPCell( new Phrase( formTitle, arialSubTitle ) );
        c.Border = PdfPCell.NO_BORDER;
        c.VerticalAlignment = Element.ALIGN_TOP;
        c.HorizontalAlignment = Element.ALIGN_LEFT;
        //c.PaddingLeft = 30f;
        c.BackgroundColor = new BaseColor( 60, 60, 60 );
        head.AddCell( c );

        // Basic request info
        string headerText = "Tracking Number: " + trackingNumber + "\nDate Time Received: " + receivedDate + "\nCreated By: " + createdBy + "";
        c = new PdfPCell( new Phrase( headerText, arialHeader ) );
        c.VerticalAlignment = Element.ALIGN_TOP;
        c.HorizontalAlignment = Element.ALIGN_LEFT;
        //c.PaddingLeft = document.LeftMargin;
        c.BackgroundColor = new BaseColor( 255, 255, 255 );
        head.AddCell( c );

        float TableBottom;
        TableBottom = page.Height - cellHeight + head.TotalHeight +5f;

        // since the table header is implemented using a PdfPTable, we call
        // WriteSelectedRows(), which requires absolute positions!
        head.WriteSelectedRows(
          0, -1,  // first/last row; -1 flags all write all rows
          document.LeftMargin,      // left offset
            // ** bottom** yPos of the table
          TableBottom,
          writer.DirectContent
        );

    }
}

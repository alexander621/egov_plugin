<%@ Application Language="C#" %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Configuration" %>
<%@ Import Namespace="System.Diagnostics" %>
<%@ Import Namespace="System.Net" %>
<%@ Import Namespace="System.Net.Mail" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Net.Security" %>
<%@ Import Namespace="System.Web.SessionState" %>

<script runat="server">

    void Application_Start(object sender, EventArgs e) 
    {
        
        // Code that runs on application startup
        System.Net.ServicePointManager.ServerCertificateValidationCallback = new System.Net.Security.RemoteCertificateValidationCallback(IgnoreCertificateErrorHandler);

        Application["environment"] = "PROD";

    }

    private bool IgnoreCertificateErrorHandler(object sender, System.Security.Cryptography.X509Certificates.X509Certificate certificate, System.Security.Cryptography.X509Certificates.X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors)
    {

        return true;
    }

    
    void Application_End(object sender, EventArgs e) 
    {
        //  Code that runs on application shutdown

    }
        
    void Application_Error(object sender, EventArgs e) 
    {
		// Code that runs when an unhandled error occurs
		Exception currentException;
		currentException = Server.GetLastError( ).GetBaseException( );
		//Response.Write("<br />Error Message: " + currentException.Message.ToString());
		Session["LastException"] = currentException;
		Session["ErrorURL"] = Request.Url.ToString( );
		Server.ClearError( );
		string iErrorId = "1";
		iErrorId = LogErrorToDB( currentException, Request.Url.ToString( ) );
		if ( Request.ServerVariables["REMOTE_ADDR"] != "24.106.89.6" && Request.ServerVariables["REMOTE_ADDR"].Substring( 0, 7 ) != "10.0.8." )
		{
			sendErrorMessage( iErrorId, currentException.Message, currentException.StackTrace, Request.Url.ToString( ) );
		}
		Server.Transfer( "errormsg.aspx?errorid=" + iErrorId );
    }

	void sendErrorMessage( string sErrorId, string Description, string sStackTrace, string sFile )
	{
		string sBody = "----------------------------------------------------------------------------------------------------";
		sBody += "\r\nEC LINK ASP.NET SCRIPT ERROR";
		sBody += "\r\n----------------------------------------------------------------------------------------------------";
		sBody += "\r\n\r\nDescription: " + Description;
		sBody += "\r\nFile: " + sFile;
		sBody += "\r\nStackTrace: " + sStackTrace + "\r\nREMOTE_ADDR:" + Request.ServerVariables["REMOTE_ADDR"];
		sBody += "\r\n\r\nFor more information regarding this error message click the link below or copy/paste link into your web browser:";
		sBody += "\r\nhttp://" + Request.ServerVariables["HTTP_HOST"] + "/eclink/admin/errormsg.aspx?errorid=" + sErrorId;
		sBody += "\r\n\r\n---------------------------------------------------------------------------------------------------";
		sBody += "\r\nEC LINK ASP.NET SCRIPT ERROR SUMMARY";
		sBody += "\r\n---------------------------------------------------------------------------------------------------";
		sBody += "\r\n\r\nThis automated message was sent by the web server because an ASP.NET script error was encountered. Do not reply to this message. Contact mailto://development@eclink.com for inquiries regarding this email.";

		MailMessage message = new MailMessage( );
		message.From = new MailAddress( "noreply@eclink.com", "ECLINK WEB SERVER" );
		message.To.Add( "egovsupport@eclink.com" );
		message.Subject = "EC LINK ASP.NET SCRIPT ERROR";
		//message.IsBodyHtml = false;
		message.Body = sBody;
		message.Priority = MailPriority.High;

        SmtpClient smtp = new SmtpClient( ConfigurationManager.AppSettings["SESMailServer"], 587 );
        NetworkCredential basicCredentials = new NetworkCredential(ConfigurationManager.AppSettings["SES_UserName"], ConfigurationManager.AppSettings["SES_Password"]);
        smtp.Credentials = basicCredentials;
        smtp.EnableSsl = true;

		smtp.Send( message );
	}

	string LogErrorToDB( Exception sNewExecption, string sNewUrl )
	{
		string sSql;
		string newErrorId = "1";
		sSql = "INSERT INTO errorlog (description, [file], webappid, category, source, sessioncollection, browserinformation, ";
		sSql += "applicationcollection, cookiescollection, servervariablescollection, requestformcollection, ";
		sSql += "requestquerystringcollection, remoteaddress, httphost ) VALUES ( '";
		sSql += dbsafeString( sNewExecption.Message ) + "', '" + dbsafeString( sNewUrl ) + "', 5, 'ASP.NET Runtime', '";
		sSql += dbsafeString( sNewExecption.StackTrace.Replace( "\r\n", "<br />" ) ) + "', '";
		sSql += dbsafeString( getSessionCollection( ) ) + "', '";
		sSql += dbsafeString( Request.ServerVariables["HTTP_USER_AGENT"] ) + "', '";
		sSql += dbsafeString( getApplicationCollection( ) ) + "', '";
		sSql += dbsafeString( getCookiesCollection( ) ) + "', '";
		sSql += dbsafeString( Request.ServerVariables["ALL_RAW"].Replace( "\r\n", "<br />" ) ) + "<br />" + dbsafeString( Request.ServerVariables["ALL_HTTP"].Replace( "\r\n", "<br />" ) ) + "', '";
		sSql += dbsafeString( getFormCollection( ) ) + "', '";
		sSql += dbsafeString( getRequestQueryStringCollection( ) ) + "', '";
		sSql += dbsafeString( Request.ServerVariables["REMOTE_ADDR"] ) + "', '";
		sSql += dbsafeString( Request.ServerVariables["HTTP_HOST"] ) + "' )";

		SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["errConn"].ConnectionString );
		SqlCommand sqlCommander = new SqlCommand( );
		sqlCommander.Connection = sqlConn;

		sqlConn.Open( );

		sqlCommander.CommandText = sSql + ";SELECT @@IDENTITY";
		newErrorId = Convert.ToString( sqlCommander.ExecuteScalar( ) );
		sqlConn.Close( );
		sqlCommander.Dispose( );

		return newErrorId;
	}

	string dbsafeString( string sValue )
	{
		string sNewString;
		sNewString = sValue.Replace( "'", "''" );
		//sNewString = sNewString.Replace("<", "&lt;");
		return sNewString;
	}

	string getSessionCollection( )
	{
		string sValue = "";
		foreach ( string sessionkey in Session.Keys )
		{
			sValue += "<br /><b>" + sessionkey + "</b>: " + Session[sessionkey];
		}
		return sValue;
	}

	string getApplicationCollection( )
	{
		string sValue = "";
		foreach ( string applicationkey in Application.Keys )
		{
			sValue += "<br /><b>" + applicationkey + "</b>: " + Application[applicationkey];
		}
		return sValue;
	}

	string getCookiesCollection( )
	{
		string sValue = "";
		foreach ( string cookiekey in Request.Cookies.Keys )
		{
			sValue += "<br /><b>" + cookiekey + "</b>: " + Request.Cookies[cookiekey].Value;
		}
		return sValue;
	}

	string getServerVariablesCollection( )
	{
		string sValue = "";
		foreach ( string key in Request.ServerVariables.Keys )
		{
			sValue += "<br /><b>" + key + "</b>: " + Request.ServerVariables[key];
		}
		return sValue;
	}

	string getFormCollection( )
	{
		string sValue = "";
		foreach ( string formkey in Request.Form )
		{
			sValue += "<br /><b>" + formkey + "</b>: " + Request.Form[formkey];
		}
		return sValue;
	}

	string getRequestQueryStringCollection( )
	{
		string sValue = "";
		foreach ( string querystringkey in Request.QueryString )
		{
			sValue += "<br /><b>" + querystringkey + "</b>: " + Request.QueryString[querystringkey];
		}
		return sValue;
	}

    void Session_Start(object sender, EventArgs e) 
    {
        // Code that runs when a new session is started

    }

    void Session_End(object sender, EventArgs e) 
    {
        // Code that runs when a session ends. 
        // Note: The Session_End event is raised only when the sessionstate mode
        // is set to InProc in the Web.config file. If session mode is set to StateServer 
        // or SQLServer, the event is not raised.
        //string sSessionID = "testing blah";
        //string sSessionID = HttpContext.Current.Session.SessionID;
        string sSessionID = Session.SessionID;

        removeAllItemsFromCart(sSessionID);

    }

    public void removeAllItemsFromCart(string iSessionID)
    {
        string sSQL = "";
        string sSessionID = "";
        
        if (iSessionID != "")
        {
            sSessionID = iSessionID;
            sSessionID = common.dbSafe(sSessionID);
        }

        sSessionID = "'" + sSessionID + "'";

        sSQL  = "SELECT cartid, ";
        sSQL += " classtimeid, ";
        sSQL += " buyorwait ";
        sSQL += " FROM egov_class_cart ";
        sSQL += " WHERE sessionid_csharp = " + sSessionID;

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                classes.removeItemFromCart(Convert.ToInt32(myReader["cartid"]),
                                           Convert.ToInt32(myReader["classtimeid"]),
                                           Convert.ToString(myReader["buyorwait"]));
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
    }
/*
    public void removeItemFromCart(Int32 iCartID,
                                   Int32 iTimeID,
                                   string iBuyOrWait)
    {
        Int32 sCartQty = 0;
        Int32 sClassID = 0;

        string sSQL = "";
        string sSQL2 = "";
        dtbSendEmail("4a");
        sSQL  = "SELECT classid, ";
        sSQL += " quantity ";
        sSQL += " FROM egov_class_cart ";
        sSQL += " WHERE cartid = " + iCartID.ToString();
        dtbSendEmail("4b - " + sSQL);
        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            sCartQty = Convert.ToInt32(myReader["quantity"]);
            sClassID = Convert.ToInt32(myReader["classid"]);
        }
        dtbSendEmail("4c");
        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        classes.updateClassTime(iTimeID,
                                sCartQty,
                                iBuyOrWait);
        dtbSendEmail("4d");
        sSQL2 = "DELETE FROM egov_class_cart_price_dtb WHERE cartid = " + iCartID;
        common.RunSQLStatement(sSQL2);
        dtbSendEmail("4e");
        sSQL2 = "DELETE FROM egov_class_cart_dtb WHERE cartid = " + iCartID;
        common.RunSQLStatement(sSQL2);
        dtbSendEmail("4f");
        classes.updateClassTimeSeriesChildren(sClassID,
                                              sCartQty,
                                              iBuyOrWait);
        dtbSendEmail("4g");
    }
*/
    public void dtbSendEmail(string iInputValue)
    {
        string sInputValue = "";
        
        if(iInputValue != "")
        {
            sInputValue = iInputValue;
            sInputValue = common.dbSafe(sInputValue);
        }

        sInputValue = "'" + sInputValue + "'";
        
        common.dtb_debug(iInputValue + " - a");

        
        common.sendMessage(37,
                           "dboyer@eclink.com",
                           "dboyer28",
                           "dboyer@eclink.com",
                           "testing global.asax",
                           Convert.ToString(DateTime.Now) + ": (" + sInputValue + ")",
                           "high",
                           true,
                           "",
                           "");
        common.dtb_debug(iInputValue + " - b");
    }       
</script>

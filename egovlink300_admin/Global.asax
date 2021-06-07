<%@ Application Language="C#" %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Configuration" %>
<%@ Import Namespace="System.Diagnostics" %>
<%@ Import Namespace="System.Net" %>
<%@ Import Namespace="System.Net.Mail" %>
<%@ Import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>


<script runat="server">

    void Application_Start(object sender, EventArgs e) 
    {
        // Code that runs on application startup

    }
    
    void Application_End(object sender, EventArgs e) 
    {
        //  Code that runs on application shutdown
        

    }
        
    void Application_Error(object sender, EventArgs e) 
    { 
        // Code that runs when an unhandled error occurs
        Exception currentException;
        currentException = Server.GetLastError().GetBaseException();
        //Response.Write("<br />Error Message: " + currentException.Message.ToString());
        Session["LastException"] = currentException;
        Session["ErrorURL"] = Request.Url.ToString();
        Server.ClearError();
        string iErrorId = "1";
        iErrorId = LogErrorToDB(currentException, Request.Url.ToString());
        if (Request.ServerVariables["REMOTE_ADDR"] != "24.106.89.6" && Request.ServerVariables["REMOTE_ADDR"].Substring(0,7) != "10.0.8.")
        {
            sendErrorMessage(iErrorId, currentException.Message, currentException.StackTrace, Request.Url.ToString());
        }
        Server.Transfer("errormsg.aspx?errorid=" + iErrorId);

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

        MailMessage message = new MailMessage();
        message.From = new MailAddress("noreplies@eclink.com", "ECLINK WEB SERVER");
        message.To.Add("devsupport@eclink.com");
        message.Subject = "EC LINK ASP.NET SCRIPT ERROR";
        //message.IsBodyHtml = false;
        message.Body = sBody;
        message.Priority = MailPriority.High;

        SmtpClient smtp = new SmtpClient( ConfigurationManager.AppSettings["SESMailServer"], 587 );
        NetworkCredential basicCredentials = new NetworkCredential(ConfigurationManager.AppSettings["SES_UserName"], ConfigurationManager.AppSettings["SES_Password"]);
        smtp.Credentials = basicCredentials;
        smtp.EnableSsl = true;

        smtp.Send(message);
    }        

    string LogErrorToDB(Exception sNewExecption, string sNewUrl)
    {
        string sSql;
        string newErrorId = "1";
        sSql = "INSERT INTO errorlog (description, [file], webappid, category, source, sessioncollection, browserinformation, ";
        sSql += "applicationcollection, cookiescollection, servervariablescollection, requestformcollection, ";
        sSql += "requestquerystringcollection, remoteaddress, httphost ) VALUES ( '";
        sSql += dbsafeString(sNewExecption.Message) + "', '" + dbsafeString(sNewUrl) + "', 5, 'ASP.NET Runtime', '";
        sSql += dbsafeString(sNewExecption.StackTrace.Replace("\r\n", "<br />")) + "', '";
        sSql += dbsafeString(getSessionCollection()) + "', '";
        sSql += dbsafeString(Request.ServerVariables["HTTP_USER_AGENT"]) + "', '";
        sSql += dbsafeString(getApplicationCollection()) + "', '";
        sSql += dbsafeString(getCookiesCollection()) + "', '";
        sSql += dbsafeString(Request.ServerVariables["ALL_RAW"].Replace("\r\n", "<br />")) + "<br />" + dbsafeString(Request.ServerVariables["ALL_HTTP"].Replace("\r\n", "<br />")) + "', '";
        sSql += dbsafeString(getFormCollection()) + "', '";
        sSql += dbsafeString(getRequestQueryStringCollection()) + "', '";
        sSql += dbsafeString(Request.ServerVariables["REMOTE_ADDR"]) + "', '";
        sSql += dbsafeString(Request.ServerVariables["HTTP_HOST"]) + "' )";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["errConn"].ConnectionString);
        SqlCommand sqlCommander = new SqlCommand();
        sqlCommander.Connection = sqlConn;

        sqlConn.Open();

        sqlCommander.CommandText = sSql + ";SELECT @@IDENTITY";
        newErrorId = Convert.ToString(sqlCommander.ExecuteScalar());
        sqlConn.Close();
        sqlCommander.Dispose();

        return newErrorId;
    }

    string dbsafeString(string sValue)
    {
        string sNewString;
        sNewString = sValue.Replace("'", "''");
        //sNewString = sNewString.Replace("<", "&lt;");
        return sNewString;
    }

    string getSessionCollection()
    {
        string sValue = "";
        foreach (string sessionkey in Session.Keys)
        {
            sValue += "<br /><b>" + sessionkey + "</b>: " + Session[sessionkey];
        }
        return sValue;
    }

    string getApplicationCollection()
    {
        string sValue = "";
        foreach (string applicationkey in Application.Keys)
        {
            sValue += "<br /><b>" + applicationkey + "</b>: " + Application[applicationkey];
        }
        return sValue;
    }

    string getCookiesCollection()
    {
        string sValue = "";
        foreach (string cookiekey in Request.Cookies.Keys)
        {
            sValue += "<br /><b>" + cookiekey + "</b>: " + Request.Cookies[cookiekey].Value;
        }
        return sValue;
    }

    string getServerVariablesCollection()
    {
        string sValue = "";
        foreach (string key in Request.ServerVariables.Keys)
        {
            sValue += "<br /><b>" + key + "</b>: " + Request.ServerVariables[key];
        }
        return sValue;
    }

    string getFormCollection()
    {
        string sValue = "";
        foreach (string formkey in Request.Form)
        {
            sValue += "<br /><b>" + formkey + "</b>: " + Request.Form[formkey];
        }
        return sValue;
    }

    string getRequestQueryStringCollection()
    {
        string sValue = "";
        foreach (string querystringkey in Request.QueryString)
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

    }
       
</script>

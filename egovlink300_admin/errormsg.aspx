<%@ Page Language="C#" %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Net" %>
<%@ Import Namespace="System.Net.Mail" %>
<%@ Import Namespace="System.Configuration" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    
    string sErrorNo;
    string sErrorPageTitle;
    
    protected void Page_Load(object sender, EventArgs e)
    {
        string ErrMsg = "";
        string sErrorId;

        if (Request.ServerVariables["REMOTE_ADDR"] == "24.106.89.6" || Request.ServerVariables["REMOTE_ADDR"] == "184.180.44.105" || Request.ServerVariables["REMOTE_ADDR"] == "74.87.250.138" || Request.ServerVariables["REMOTE_ADDR"].Substring( 0, 7 ) == "10.0.8." || Request.ServerVariables["REMOTE_ADDR"].Substring( 0, 8 ) == "10.0.48." || Request.ServerVariables["REMOTE_ADDR"].Substring( 0, 8 ) == "10.0.12.")
        {
            //ErrMsg = "ERROR DETAILS - ( " + Request["errorid"] + " )";

            if (string.IsNullOrEmpty(Request["errorid"]))
                sErrorId = "0";
            else
                sErrorId = Request["errorid"];
            
            SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
            sqlConn.Open();

            Int32 iErrorNo = Int32.Parse(sErrorId);
            sErrorNo = iErrorNo.ToString();
            sErrorPageTitle = "ERROR DETAILS - ( " + sErrorNo + " )";
            
            string sSql = "SELECT errordatetime, category, description, [file] AS errorfile, source, sessioncollection, ";
            sSql += "cookiescollection, browserinformation, applicationcollection, servervariablescollection, ";
            sSql += "requestformcollection, requestquerystringcollection, remoteaddress, httphost ";
            sSql += "FROM errorlog WHERE rowid = " + iErrorNo.ToString();
            //ErrMsg += "<br />" + sSql;
            SqlCommand myCommand = new SqlCommand(sSql, sqlConn);
            SqlDataReader myReader;
            myReader = myCommand.ExecuteReader();
            
            while ( myReader.Read() ) 
            {

                ErrMsg += "<div class=\"errorbox\"><strong>Error Date:</strong> " + fixHTML(myReader["errordatetime"]);
                ErrMsg += "<br /><strong>Category:</strong> " + fixHTML(myReader["category"]) + "</div>";

                ErrMsg += "<strong>Error Object Information</strong><br />";
                ErrMsg += "<div class=\"errorbox\"><strong>Description:</strong> " + fixHTML(myReader["description"]);
                ErrMsg += "<br /><strong>File:</strong> " + fixHTML(myReader["errorfile"]);
                ErrMsg += "<br /><strong>Host:</strong> " + fixHTML(myReader["httphost"]);
                ErrMsg += "<br /><strong>Remote Address:</strong> " + fixHTML(myReader["remoteaddress"]) + "</div>";

                ErrMsg += "<strong>Stacktrace</strong><br />";
                ErrMsg += "<div class=\"errorbox\">" + fixHTML(myReader["source"]) + "</div>";

                ErrMsg += "<strong>Browser Information</strong><br />";
                ErrMsg += "<div class=\"errorbox\">" + fixHTML(myReader["browserinformation"]) + "</div>";

                ErrMsg += "<strong>Application Collection</strong><br />";
                ErrMsg += "<div class=\"errorbox\">" + fixHTML(myReader["applicationcollection"]) + "</div>";

                ErrMsg += "<strong>Request Form Collection</strong><br />";
                ErrMsg += "<div class=\"errorbox\">" + fixHTML(myReader["requestformcollection"]) + "</div>";

                ErrMsg += "<strong>Querystring Collection</strong><br />";
                ErrMsg += "<div class=\"errorbox\">" + fixHTML(myReader["requestquerystringcollection"]) + "</div>";

                ErrMsg += "<strong>Session Collection</strong><br />";
                ErrMsg += "<div class=\"errorbox\">" + fixHTML(myReader["sessioncollection"]) + "</div>";

                ErrMsg += "<strong>Cookies Collection</strong><br />";
                ErrMsg += "<div class=\"errorbox\">" + fixHTML(myReader["cookiescollection"]) + "</div>";

                ErrMsg += "<strong>Server Variable Collection</strong><br />";
                ErrMsg += "<div class=\"errorbox\">" + fixHTML(myReader["servervariablescollection"]) + "</div>";

            }
            Label1.Text = ErrMsg;
            
            myReader.Close();
            sqlConn.Close();
        }
        else
        {
            sErrorPageTitle = "ERROR";
            Label1.Text = "<br /><font style=\"font-family: arial,tahoma;font-size: 12px;\" ><div id=\"errorpublic\">We're sorry, but the server encountered an error processing your request. An email notification has been sent to our technical department with the details of this error.</div><br /><br /><br /><hr size=\"1\" color=\"#000000\" ><center>Developed by <i>electronic commerce</i> link, inc. dba. <i>ec</i> link.</font></center>";
        }
    }

    public string fixHTML(object input)
    {
	    return input.ToString().Replace("<", "&lt;").Replace("&lt;br","<br").Replace("&lt;b>","<b>").Replace("&lt;/b","</b");
    }

    
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>APPLICATION ERROR</title>
    
    <META NAME="ROBOTS" CONTENT="NOINDEX">
    
    <link rel="stylesheet" type="text/css" href="global.css" />
    
    <style>
		a:link {font:8pt/11pt verdana; color:#FF0000;}
		a:visited {font:8pt/11pt verdana; color:#4e4e4e;}
		
		div.errorbox {
		    text-align:left; 
		    width:90%; 
		    border: 1px solid #000000;
		    font-family: arial,tahoma; 
		    font-size: 12px; 
		    color:#000000;
		    padding:5px;
		    margin-bottom: 1.5em;
		    background-color:#e0e0e0;
		    }
		    
        div#errorpublic {
            text-align: left; 
            border: 1px solid #c0c0c0;
            font-family: arial,tahoma; 
            font-size: 14px; 
            color: yellow;
            padding: 5px;
            background-color: #FF0000;
            }
	</style>
	
</head>
<body>
    <form id="form1" runat="server">
    <!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong><%=sErrorPageTitle%></strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

            <div>
                <asp:Label ID="Label1" runat="server" Text="Error Messsage"></asp:Label><br /><br />
            </div>
   		</div>
	</div>
    <!--END PAGE CONTENT-->
    </form>
</body>
</html>

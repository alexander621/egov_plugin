using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;

/// <summary>
/// This is the common methods for egov on the public side. You do not need to instantiate this class.
/// Try to keep these in alphabetical order, please.
/// </summary>
public static class common
{

	public static void AddToPaymentLog( string iOrgId, string sFeature, string iPaymentControlNumber, string sLogEntry )
	{
		string sSql = "INSERT INTO paymentlog ( paymentcontrolnumber, orgid, applicationside, feature, logentry ) VALUES ( ";
		sSql += iPaymentControlNumber + ", " + iOrgId + ", 'public', '" + sFeature + "', '" + dbready_string( sLogEntry, 450 ) + "' )";
		common.RunSQLStatement( sSql );
	}


	public static string CreatePaymentControlRow( string iOrgId, string sFeature, string sLogEntry )
	{
		string sSql = "INSERT INTO paymentlog ( orgid, applicationside, feature, logentry ) VALUES ( ";
		sSql += iOrgId + ", 'public', '" + sFeature + "', '" + sLogEntry + "' )";

		string iPaymentControlNumber = RunInsertStatement( sSql );

		return iPaymentControlNumber;
	}


	public static string dbready_string( string sValue, int iMaxLength )
	{
		string sReturn = "";

		if ( sValue != "" && iMaxLength > 0 )
		{
			sReturn = sValue.Trim( );
			sReturn = sReturn.Replace( "<", "&lt;" );
			sReturn = sReturn.Replace( ">", "&gt;" );
			if ( sReturn.Length > iMaxLength )
				sReturn = sReturn.Substring( 0, iMaxLength );

			sReturn = sReturn.Replace( "'", "''" );
		}
		return sReturn;
	}


    public static string dbSafe(string sValue)
    {
        string sNewString = "";
        
        if ( ! String.IsNullOrEmpty(sValue) )
        {
            sNewString = sValue.Trim( );
            sNewString = sNewString.Replace("'", "''");
            sNewString = sNewString.Replace("<", "&lt;");
        }
        return sNewString;
    }

    public static string getOrgId( )
    {
        string sOrgId    = "0"; 
        string sProtocol = "http://";
        string sOrgURL   = "";
        string sSQL      = "";

        //Build the current URL
        if (HttpContext.Current.Request.ServerVariables["HTTPS"].ToUpper() == "ON")
        {
            sProtocol = "https://";
        }

        sOrgURL  = sProtocol;
        sOrgURL += HttpContext.Current.Request.ServerVariables["server_name"].ToLower();
        sOrgURL += "/";
        sOrgURL += common.GetVirtualDirectyName(HttpContext.Current.Request.ServerVariables["URL"].ToLower());

        //NOTE: Since the URL is stored in the Organizations table as "http://dev4.egovlink.com",
        //we have to replace "https://secure.egovlink.com" with the proper URL if/when we are
        //in the secure site so that we can get the orgid.
        sOrgURL = sOrgURL.Replace("https://secure.egovlink.com", ConfigurationManager.AppSettings["baseURL"]);

		if (sOrgURL.IndexOf("//egovlink.com") > 0)
		{
			sOrgURL = sOrgURL.Replace("//egovlink.com","//www.egovlink.com");
		}

        sSQL = "SELECT * FROM Organizations WHERE OrgEgovWebsiteURL = '" + sOrgURL.Replace("https:","http:") + "'";
        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();
        
        while ( myReader.Read() ) 
        {
            sOrgId = myReader["orgid"].ToString();
        }

        myReader.Close();
        sqlConn.Close();

        return sOrgId;
    }

    public static string getBaseURL(string iCurrentURL)
    {
        string lcl_return  = "";
        string sCurrentURL = ConfigurationManager.AppSettings["baseURL"];

        if (iCurrentURL != "")
        {
            sCurrentURL = iCurrentURL.ToLower();
        }

        if (HttpContext.Current.Request.ServerVariables["HTTPS"].ToUpper() == "ON")
        {
            sCurrentURL = sCurrentURL.Replace(ConfigurationManager.AppSettings["baseURL"], "https://secure.egovlink.com");
        }

        lcl_return = sCurrentURL;

        return lcl_return;
    }

    public static string getOrgFullSite(string _OrgId)
    {
        string FullSite = "";
        string Sql = "SELECT OrgEgovWebsiteURL FROM organizations WHERE orgid = " + _OrgId;

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(Sql, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();
            FullSite = myReader["OrgEgovWebsiteURL"].ToString();
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return FullSite.Replace("http:","https:");
    }

    public static string getOrgName(string _OrgId)
    {
        string OrgName = "";
        string Sql = "SELECT orgname FROM organizations WHERE orgid = " + _OrgId;

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(Sql, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();
            OrgName = myReader["orgname"].ToString();
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return OrgName;
    }

    public static string getOrgInfo(string iOrgID, 
                                    string iColumnName)
    {
        Int32 sOrgID = 0;
        
        string sSQL        = "";
        string sColumnName = "";
        string lcl_return  = "";

        if (iOrgID != null)
        {
            try
            {
                sOrgID = Convert.ToInt32(iOrgID);
            }
            catch
            {
                lcl_return = "";
            }
        }

        if(iColumnName != null) {
            sColumnName = iColumnName;
        }

        if (sOrgID != null && sColumnName != null)
        {
            sSQL = "SELECT o." + sColumnName + " as dbcolumn";
            sSQL += " FROM organizations o ";
            sSQL += " INNER JOIN TimeZones t ON o.orgTimeZoneID = t.timeZoneID ";
            sSQL += " WHERE orgid = " + sOrgID;

            SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
            sqlConn.Open();

            SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
            SqlDataReader myReader;
            myReader = myCommand.ExecuteReader();

            if (myReader.HasRows)
            {
                myReader.Read();
                lcl_return = myReader["dbcolumn"].ToString();
            }

            myReader.Close();
            sqlConn.Close();
            myReader.Dispose();
            sqlConn.Dispose();
        }

        return lcl_return;
    }

    public static Boolean checkMenuOptionEnabled(string iOrgID, string iField)
    {
        Boolean lcl_return = false;
        Int32 sOrgID       = 0;
        string sSQL        = "";

        if (iOrgID != null)
        {
            try
            {
                sOrgID = Convert.ToInt32(iOrgID);
            }
            catch
            {
                lcl_return = false;
            }
        }

        sSQL  = "SELECT public_menuopt_" + iField + "home_enabled AS MenuOpt_Enabled ";
        sSQL += " FROM organizations ";
        sSQL += " WHERE orgid = " + sOrgID;

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();
            lcl_return = Convert.ToBoolean(myReader["MenuOpt_Enabled"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
        
        return lcl_return;
    }

    public static string getMenuOptionLabel(string iOrgID, string iField)
    {
        string lcl_return = "";
        Int32  sOrgID     = 0;
        string sSQL       = "";

        if (iOrgID != null)
        {
            try
            {
                sOrgID = Convert.ToInt32(iOrgID);
            }
            catch
            {
                lcl_return = "";
            }
        }

        sSQL = "SELECT public_menuopt_" + iField + "home_label AS MenuOpt_Label ";
        sSQL += " FROM organizations ";
        sSQL += " WHERE orgid = " + sOrgID;

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();
            lcl_return = myReader["MenuOpt_Label"].ToString();
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string GetVirtualDirectyName( string sURL )
    {
        string[] sNewUrl = sURL.Split('/');
        return sNewUrl[1];
    }

    public static string Left( string _Value, int _Length )
    {
        if (_Length < 0)
            throw new ArgumentOutOfRangeException( "length", _Length, "length must be > 0" );
        else if (_Length == 0 || _Value.Length == 0)
            return "";
        else if (_Value.Length <= _Length)
            return _Value;
        else
            return _Value.Substring( 0, _Length );
    }

    public static string Right( string _Value, int _Length )
    {
        if (_Length < 0)
            throw new ArgumentOutOfRangeException( "length", _Length, "length must be > 0" );
        else if (_Length == 0 || _Value.Length == 0)
            return "";
        else if (_Value.Length <= _Length)
            return _Value;
        else
            return _Value.Substring( _Value.Length - _Length );
    }

    public static string RunInsertStatement( string sSql )
    {
        string iNewId = "0";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        SqlCommand sqlCommander = new SqlCommand();
        sqlCommander.Connection = sqlConn;

        sqlConn.Open();

        sqlCommander.CommandText = sSql + ";SELECT @@IDENTITY";
        iNewId = Convert.ToString(sqlCommander.ExecuteScalar());
        sqlConn.Close();
        sqlCommander.Dispose();

        return iNewId;
    }

    public static void RunSQLStatement( string sSql )
    {
        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);

        SqlCommand sqlCommander = new SqlCommand();
        sqlCommander.Connection = sqlConn;

        sqlConn.Open();

        sqlCommander.CommandText = sSql;

        sqlCommander.ExecuteNonQuery();

        sqlConn.Close();
        sqlCommander.Dispose();
    }

    public static void setOrganizationSessionVariables( )
    {
        // Initialize the session variables. THis prevents problems later
        HttpContext.Current.Session["orgid"] = "0";
        HttpContext.Current.Session["egovclientwebsiteurl"] = "";

        string sProtocol = "http://";
        if (HttpContext.Current.Request.ServerVariables["HTTPS"].ToUpper() == "ON")
            sProtocol = "https://";

        string sOrgURL = sProtocol + HttpContext.Current.Request.ServerVariables["server_name"].ToLower() + "/" + common.GetVirtualDirectyName(HttpContext.Current.Request.ServerVariables["URL"].ToLower());

        string sSql = "SELECT * FROM Organizations INNER JOIN TimeZones ON Organizations.OrgTimeZoneID = TimeZones.TimeZoneID WHERE OrgEgovWebsiteURL = '" + sOrgURL + "'";
        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSql, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();
        
        while ( myReader.Read() ) 
        {
            //add session variables to the list below as needed
            HttpContext.Current.Session["orgid"] = myReader["orgid"];
            HttpContext.Current.Session["egovclientwebsiteurl"] = myReader["OrgEgovWebsiteURL"];
        }

        myReader.Close();
        sqlConn.Close();
    }

    public static Boolean orgHasFeature(string iOrgID, string iFeature)
    {
        Boolean lcl_return = false;
        int sOrgID = 0;
        string sFeature = "";
        string sSQL = "";

        if (iOrgID != null)
        {
            try
            {
                sOrgID = Convert.ToInt32(iOrgID);
            }
            catch
            {
                lcl_return = false;
            }
        }

        if (iFeature != null)
        {
            sFeature = iFeature.Trim();
            sFeature = dbSafe(sFeature);
            sFeature = "'" + sFeature + "'";
        }

        sSQL = "SELECT COUNT(fo.featureid) as feature_count ";
        sSQL += " FROM egov_organizations_to_features fo, ";
        sSQL += " egov_organization_features f ";
        sSQL += " WHERE fo.featureid = f.featureid ";
        sSQL += " AND f.feature = " + sFeature;
        sSQL += " AND fo.orgid = " + sOrgID;

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (Convert.ToInt32(myReader["feature_count"]) > 0)
            {
                lcl_return = true;
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Boolean orgHasDisplay(string iOrgID, string iDisplay)
    {
        Boolean lcl_return = false;
        Int32 sOrgID       = 0;
        string sDisplay    = "";
        string sSQL        = "";

        if (iOrgID != null)
        {
            try
            {
                sOrgID = Convert.ToInt32(iOrgID);
            }
            catch
            {
                lcl_return = false;
            }
        }

        if (iDisplay != null)
        {
            sDisplay = iDisplay.Trim();
            sDisplay = dbSafe(sDisplay);
            sDisplay = "'" + sDisplay + "'";
        }

        sSQL = "SELECT COUNT(od.displayid) as display_count ";
        sSQL += " FROM egov_organizations_to_displays od, ";
        sSQL += " egov_organization_displays d ";
        sSQL += " WHERE od.displayid = d.displayid ";
        sSQL += " AND od.orgid = " + sOrgID;
        sSQL += " AND d.display = " + sDisplay;
        
        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (Convert.ToInt32(myReader["display_count"]) > 0)
            {
                lcl_return = true;
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getOrgDisplay(string iOrgID, string iDisplay)
    {
        string lcl_return = "";
        string sSQL       = "";

        sSQL  = "SELECT isnull(od.displaydescription, d.displaydescription) AS displaydescription ";
        sSQL += " FROM egov_organizations_to_displays od, ";
        sSQL +=      " egov_organization_displays d ";
        sSQL += " WHERE od.displayid = d.displayid ";
        sSQL += " AND od.orgid = " + iOrgID;
        sSQL += " AND d.display = '" + iDisplay + "'";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = myReader["displaydescription"].ToString();
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Boolean publicCanViewFeature(string iOrgID, string iFeature)
    {
        Boolean lcl_return = false;
        int sOrgID         = 0;
        string sFeature    = "";
        string sSQL        = "";

        if (iOrgID != null)
        {
            try
            {
                sOrgID = Convert.ToInt32(iOrgID);
            }
            catch
            {
                lcl_return = false;
            }
        }

        if (iFeature != null)
        {
            sFeature = iFeature.Trim();
            sFeature = dbSafe(sFeature);
            sFeature = "'" + sFeature + "'";
        }

        sSQL = "SELECT isnull(fo.publiccanview,0) as publiccanview ";
        sSQL += " FROM egov_organizations_to_features fo, ";
        sSQL +=      " egov_organization_features f ";
        sSQL += " WHERE fo.featureid = f.featureid ";
        sSQL += " AND f.feature = " + sFeature;
        sSQL += " AND fo.orgid = " + sOrgID;

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if(myReader["publiccanview"] != null)
            {
                lcl_return = Convert.ToBoolean(myReader["publiccanview"]);
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Boolean checkAccessToList(string iUserID, string iOrgID, string iListType)
    {
        Boolean lcl_return = false;
        string sSQL        = "";
        Int32 sOrgID       = 0;
        Int32 sUserID      = 0;

        if (iOrgID != null)
        {
            try
            {
                sOrgID = Convert.ToInt32(iOrgID);
            }
            catch
            {
                lcl_return = false;
            }
        }

        if (iUserID != null)
        {
            try
            {
                sUserID = Convert.ToInt32(iUserID);
            }
            catch
            {
                lcl_return = false;
            }
        }

        if (iListType != null)
        {
            sSQL = "SELECT isDoNotKnockVendor_" + iListType + " AS 'ListAccess' ";
            sSQL += " FROM egov_users ";
            sSQL += " WHERE orgid = " + sOrgID;
            sSQL += " AND userid = " + sUserID;

            SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
            sqlConn.Open();

            SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
            SqlDataReader myReader;
            myReader = myCommand.ExecuteReader();

            if (myReader.HasRows)
            {
                myReader.Read();

                lcl_return = Convert.ToBoolean(myReader["ListAccess"]);
            }

            myReader.Close();
            sqlConn.Close();
            myReader.Dispose();
            sqlConn.Dispose();
        }

        return lcl_return;
    }

    public static void logThePageVisit(double iStartCounter, string iPageURL, string iApplicationSide)
    {
        //Log the page visit
        double endCounter  = DateTime.Now.TimeOfDay.TotalSeconds;
        double loadSeconds = endCounter - iStartCounter;

        DateTime localNow          = DateTime.Now;
        DateTimeOffset localOffset = new DateTimeOffset(localNow);
        string sLogDate            = string.Format("{0:M/d/yyyy HH:mm:ss}",localOffset);

        string sApplicationSide = "'" + common.dbSafe(iApplicationSide) + "'";
        string sPageURL         = "'" + common.dbSafe(iPageURL)         + "'";
        string sessionValues    = "";
        string cookieValues     = "";
        string formValues       = "";
        string sSQL             = "";
        string userAgentGroup   = "";
        string sUserName        = "";

        //sUserName = common.getUserName(sUserID);
        
        foreach (string sessionkey in HttpContext.Current.Session.Keys)
        {
            sessionValues += "<br /><strong>" + sessionkey + "</strong>: " + HttpContext.Current.Session[sessionkey];
        }

        foreach (string cookiekey in HttpContext.Current.Request.Cookies.Keys)
        {
            cookieValues += "<br /></strong>" + cookiekey + "</strong>: " + HttpContext.Current.Request.Cookies[cookiekey].Value;
        }

        foreach (string formkey in HttpContext.Current.Request.Form)
        {
            formValues += "<br /><strong>" + formkey + "</strong>: " + HttpContext.Current.Request.Form[formkey];
        }
        
        if (HttpContext.Current.Request.ServerVariables["HTTP_USER_AGENT"].ToString().Length > 0)
        {
            userAgentGroup = common.getUserAgentGroup(HttpContext.Current.Request.ServerVariables["HTTP_USER_AGENT"].ToString().ToLower());
        } else {
            userAgentGroup = common.getUntrackedUserAgentGroup();
        }

        sSQL  = "INSERT INTO egov_pagelog (";
        sSQL += "logdate,";
        sSQL += "virtualdirectory,";
        sSQL += "applicationside,";
        sSQL += "page,";
        sSQL += "loadtime,";
        sSQL += "scriptname,";
        sSQL += "querystring,";
        sSQL += "servername,";
        sSQL += "remoteaddress,";
        sSQL += "requestmethod,";
        sSQL += "orgid,";
        sSQL += "userid,";
        sSQL += "username,";
        sSQL += "reportloadtime,";
        sSQL += "sectionid,";
        sSQL += "documenttitle,";
        sSQL += "useragent,";
        sSQL += "useragentgroup,";
        sSQL += "sessioncollection,";
        sSQL += "cookiescollection,";
        sSQL += "requestformcollection,";
        sSQL += "sessionid";
        sSQL += ") VALUES (";
        sSQL += "'" + sLogDate + "', ";
        sSQL += "'" + HttpContext.Current.Request.ApplicationPath + "', ";
        sSQL += sApplicationSide + ", ";
        sSQL += sPageURL         + ", ";
        sSQL += loadSeconds.ToString("0.000") + ", ";
        sSQL += "'" + HttpContext.Current.Request.ServerVariables["URL"].ToLower() + "', ";
        sSQL += "'" + common.dbSafe(common.Left(HttpContext.Current.Request.ServerVariables["QUERY_STRING"].ToLower(), 500)) + "', ";
        sSQL += "'" + common.dbSafe(HttpContext.Current.Request.ServerVariables["SERVER_NAME"].ToLower())     + "', ";
        sSQL += "'" + common.dbSafe(HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"].ToLower())     + "', ";
        sSQL += "'" + common.dbSafe(HttpContext.Current.Request.ServerVariables["REQUEST_METHOD"].ToString()) + "', ";
        sSQL += common.getOrgId() + ", ";
        sSQL += "'" + HttpContext.Current.Session["userid"] + "', ";
        sSQL += "'" + sUserName + "', ";
        sSQL += "NULL, ";
        sSQL += "NULL, ";
        sSQL += "NULL, ";
        sSQL += "'" + common.dbSafe(HttpContext.Current.Request.ServerVariables["HTTP_USER_AGENT"].ToString()) + "', ";
        sSQL += "'" + userAgentGroup               + "', ";
        sSQL += "'" + common.dbSafe(sessionValues) + "', ";
        sSQL += "'" + common.dbSafe(cookieValues)  + "', ";
        sSQL += "'" + common.dbSafe(formValues)    + "', ";
        sSQL += "'" + common.dbSafe(HttpContext.Current.Session.SessionID) + "'";
        sSQL += ")";
        
        //common.dtb_debug(sSQL);
        common.RunSQLStatement(sSQL);
    }

    public static string getUntrackedUserAgentGroup()
    {
        string userAgentGroup = "";

        string sSQL  = "SELECT useragentgroup ";
               sSQL += " FROM UserAgent_Group ";
               sSQL += " WHERE isuntracked = 1";
        
        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        //myReader = myCommand.ExecuteReader();  //<<<<---- THE PROBLEM IS HERE!!!
        /*
        if (myReader.HasRows)
        {
            myReader.Read();
            userAgentGroup = myReader["useragentgroup"].ToString();
        }
        */

        //myReader.Close();
        sqlConn.Close();
        //myReader.Dispose();
        sqlConn.Dispose();

        return userAgentGroup;
    }

    public static string changeBGColor(string iValue, string iColor1, string iColor2)
    {
        string lcl_return = "#ffffff";
        string sColor1 = "#eeeeee";
        string sColor2 = "#ffffff";

        if (iColor1 != "")
        {
            sColor1 = iColor1;
        }

        if (iColor2 != "")
        {
            sColor2 = iColor2;
        }

        if (iValue == sColor1)
        {
            lcl_return = sColor2;
        } else {
            lcl_return = sColor1;
        }

        return lcl_return;
    }

    public static string getUserAgentGroup(string iUserAgent)
    {
        string userAgentGroup = "";
        string sSQL           = "";

        //set the default in case we do not find a match
        userAgentGroup = common.getUntrackedUserAgentGroup();
        /*
        sSQL  = "SELECT useragentgroup ";
        sSQL += " FROM UserAgent_Groups ";
        sSQL += " WHERE isuntracked = 0 ";
        sSQL += " AND isactive = 1 ";
        sSQL += " ORDER BY checkorder ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                if (iUserAgent.IndexOf(myReader["useragentgroup"].ToString()) != -1)
                {
                    userAgentGroup = myReader["useragentgroup"].ToString();
                    break;
                }
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
        */

        return userAgentGroup;
    }

    public static string getFeatureName(string iOrgID, string iFeature)
    {
        Int32 sOrgID = 0;

        string lcl_return = "";
        string sSQL       = "";
        string sFeature   = "";

        if (iOrgID != null)
        {
            try
            {
                sOrgID = Convert.ToInt32(iOrgID);
            }
            catch
            {
                sOrgID = 0;
            }
        }

        sFeature = iFeature.Replace("'","''");
        sFeature = "'" + sFeature + "'";

        sSQL  = "SELECT ISNULL(fo.featurename, f.featurename) as featurename ";
        sSQL += " FROM egov_organizations_to_features fo, ";
        sSQL +=      " egov_organization_features f ";
        sSQL += " WHERE fo.featureid = f.featureid ";
        sSQL += " AND fo.orgid = " + sOrgID.ToString();
        sSQL += " AND feature = " + sFeature;

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = myReader["featurename"].ToString();
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string formatPhoneNumber(string iPhoneNumber)
    {
        string lcl_return   = iPhoneNumber;
        string sPhoneLeft   = "";
        string sPhoneMiddle = "";
        string sPhoneRight  = "";

        if (iPhoneNumber.Length == 10)
        {
            sPhoneLeft   = iPhoneNumber.Substring(0, 3);
            sPhoneMiddle = iPhoneNumber.Substring(3, 3);
            sPhoneRight  = iPhoneNumber.Substring(6);

            lcl_return = "(" + sPhoneLeft + ") " + sPhoneMiddle + "-" + sPhoneRight;
        }

        return lcl_return;
    }

    public static Boolean UserIsMissingKeyData(Int32 iUserID)
    {
        Boolean lcl_return = false;
        
        string sSQL = "";

        sSQL  = "SELECT isnull(userfname,'') as userfname, ";
        sSQL += " isnull(userlname,'') as userlname, ";
        sSQL += " isnull(userhomephone,'') as userhomephone ";
        sSQL += " FROM egov_users ";
        sSQL += " WHERE userid = " + iUserID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (myReader["userfname"].ToString() == "" ||
                myReader["userlname"].ToString() == "" ||
                myReader["userhomephone"].ToString() == "")
            {
                lcl_return = true;
            }
        }
        else
        {
            lcl_return = true;
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Int32 checkUserSessionExists(Int32 iOrgID,
                                               string iSessionID)
    {
        Int32 lcl_return = 0;

        string sSQL       = "";
        string sSessionID = "";

        if(iSessionID != "")
        {
            sSessionID = common.dbSafe(iSessionID);
            sSessionID = "'" + sSessionID + "'";
        }

        sSQL  = "SELECT usersessionid ";
        sSQL += " FROM egov_aspnet_to_asp_usersessions ";
        sSQL += " WHERE orgid = " + iOrgID.ToString();
        sSQL += " AND sessionid = " + sSessionID;

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToInt32(myReader["usersessionid"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getEGovDefaultPage(Int32 iOrgID)
    {
        string lcl_return = "default.asp";
        string sSQL       = "";

        sSQL  = "SELECT orgEGovWebsiteURL ";
        sSQL += " FROM organizations ";
        sSQL += " WHERE orgid = " + iOrgID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToString(myReader["orgEGovWebsiteURL"]);
            lcl_return = lcl_return + "/";
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static void sendMessage(Int32 iOrgID,
                                   string iFromEmail,
                                   string iFromName,
                                   string iToEmail,
                                   string iSubject,
                                   string iBody,
                                   string iPriority,
                                   Boolean iIsHTMLFormat,
                                   string iCcEmail,
                                   string iBccEmail)
    {
        MailMessage message = new MailMessage();
        string sEGovDefaultPage = common.getEGovDefaultPage(iOrgID);

        // From - The name can be optional
        //if (iFromName.Trim() != "")
            //message.From = new MailAddress(iFromEmail, iFromName);
        //else
        //{
            //if (!iFromEmail.ToLower().StartsWith("noreply") && iFromEmail != "")
            //{
                //message.From = new MailAddress(iFromEmail);
            //}
            //else
            //{
                //Find the current domain
                string currentDomain = HttpContext.Current.Request.ServerVariables["SERVER_NAME"].ToLower();

                if (currentDomain.IndexOf(".") <= 0 || currentDomain.IndexOf(".") == currentDomain.Length)
                {
                    /*
                    string orgId = "0";

                    if (HttpContext.Current.Request.Cookies["orgid"] != null)
                    {
                        if (HttpContext.Current.Request.Cookies["orgid"].Value != "" && HttpContext.Current.Request.Cookies["orgid"].Value != null)
                            orgId = HttpContext.Current.Request.Cookies["orgid"].Value;
                    }
                    //Lookup user's org url
                    string sql = "SELECT ccmlinkurl"
                                    + " FROM organizations "
                                    + " WHERE orgid = " + orgId;

                    SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
                    sqlConn.Open();

                    SqlCommand myCommand = new SqlCommand(sql, sqlConn);
                    SqlDataReader myReader;
                    myReader = myCommand.ExecuteReader();

                    if (myReader.HasRows)
                    {
                        myReader.Read();
                        currentDomain = myReader["ccmlinkurl"].ToString();
                    }

                    myReader.Close();
                    sqlConn.Close();
                    myReader.Dispose();
                    sqlConn.Dispose();
                    */

                    //Uri orgUri = new Uri(currentDomain);sEGovDefaultPage
                    Uri orgUri = new Uri(sEGovDefaultPage);
                    currentDomain = orgUri.Host;
                }
                currentDomain = currentDomain.Substring(currentDomain.IndexOf(".") + 1);

                message.From = new MailAddress("noreplies@egovlink.com", currentDomain);
            //}
        //}

        // To - This can be a list of email addresses separated by commas
        message.To.Add(iToEmail);

        // CC - This is the cc list. It can be a comma separated list
        if (iCcEmail != "")
            message.CC.Add(iCcEmail);

        // BCC - This is the bcc list. It can be a comma separated list
        if (iBccEmail != "")
            message.Bcc.Add(iBccEmail);

        // Subject
        if (iSubject != "")
        {
            string messageSubject = iSubject;
            messageSubject = messageSubject.Replace("\n", " ").Trim();
            message.Subject = messageSubject;
        }

        // Body
        if (iBody != "")
            message.Body = iBody;

        // Flag indicating HTML or Text Body format
        message.IsBodyHtml = iIsHTMLFormat;

        // Message Priority
        if (iPriority.ToLower() == "high")
            message.Priority = MailPriority.High;
        else
            message.Priority = MailPriority.Normal;

        // Point to the email server
        SmtpClient smtp = new SmtpClient( ConfigurationManager.AppSettings["SESMailServer"], 587 );
        NetworkCredential basicCredentials = new NetworkCredential(ConfigurationManager.AppSettings["SES_UserName"], ConfigurationManager.AppSettings["SES_Password"]);
        smtp.Credentials = basicCredentials;
        smtp.EnableSsl = true;

        // Send the message
	if( !isSuppressed(iToEmail, iCcEmail, iBccEmail) )
	{
        	smtp.Send(message);
	}

        message.Dispose();
    }

    public static bool isSuppressed(string ToEmails, string CCEmails, string BCCEmails)
    {
	    bool retVal = false;

	    //Check To Emails
	    retVal = checkEmailList(ToEmails);

	    //Check CCs
	    if (!String.IsNullOrEmpty(CCEmails) && !retVal)
	    {
		retVal = checkEmailList(CCEmails);
	    }

	    //Check BCCEmails
	    if (!String.IsNullOrEmpty(BCCEmails) && !retVal)
	    {
		retVal = checkEmailList(BCCEmails);
	    }
	    
	    return retVal;
    }

    public static bool checkEmailList(string emailList)
    {
	    bool retVal = false;
	    List<string> emails = emailList.Split(',').ToList();
	    foreach (string email in emails)
	    {
		if (checkSuppressionList(email.Trim()) && !retVal)
		{
			retVal = true;
		}
	    }
	    return retVal;
    }

    public static bool checkSuppressionList(string email)
    {
	bool retVal = false;

	string sSQL = "SELECT TOP 1 emailsuppressionid FROM emailsuppressionlist WHERE emailaddress = '" + dbSafe(email) + "'";
        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
	    retVal = true;
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

	    return retVal;
    }

    public static string displayGenderPicks(string iElement,
                                            string iGenderMatch)
    {
        string lcl_return      = "";
        string sSQL            = "";
        string sSelectedGender = "";

        sSQL  = "SELECT gender, ";
        sSQL += " genderdescription ";
        sSQL += " FROM egov_user_genders ";
        sSQL += " ORDER BY displayorder ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            lcl_return = "<select name=\"" + iElement + "\" id=\"" + iElement + "\">";
            lcl_return += "  <option value=\"N\">Select a gender...</option>";

            while (myReader.Read())
            {
                if (iGenderMatch == Convert.ToString(myReader["gender"]))
                {
                    sSelectedGender = " selected=\"selected\"";
                }

                lcl_return += "  <option value=\"" + Convert.ToString(myReader["gender"]) + "\"" + sSelectedGender + ">" + Convert.ToString(myReader["genderdescription"]) + "</option>";
            }

            lcl_return += "</select>";
        }
        else
        {
            lcl_return = "<input type=\"text\" name=\"" + iElement + "\" id=\"" + iElement + "\" value=\"N\" />";
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;

    }

    public static Int32 getDisplayID(string iDisplay)
    {
        Int32 lcl_return = 0;

        string sSQL     = "";
        string sDisplay = "";

        if(iDisplay != "")
        {
            sDisplay = iDisplay.ToString();
            sDisplay = dbSafe(sDisplay);
            sDisplay = "'" + sDisplay + "'";
        }

        sSQL  = "SELECT displayid ";
        sSQL += " FROM egov_organization_displays ";
        sSQL += " WHERE display = " + sDisplay;

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToInt32(myReader["displayid"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getOrgDisplayWithID(Int32 iOrgID,
                                             Int32 iDisplayID,
                                             Boolean iUsesDisplayName)
    {
        string lcl_return   = "";
        string sSQL         = "";
        string sSearchField = "displaydescription";

        if(iUsesDisplayName)
        {
            sSearchField = "displayname";
        }

        sSQL  = "SELECT isnull(od." + sSearchField + ", d." + sSearchField + ") as displayfield ";
        sSQL += " FROM egov_organizations_to_displays od, ";
        sSQL +=      " egov_organization_displays d ";
        sSQL += " WHERE od.displayid = d.displayid ";
        sSQL += " AND orgid = " + iOrgID.ToString();
        sSQL += " AND d.displayid = " + iDisplayID;

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToString(myReader["displayfield"]);
        }
        else
        {
            lcl_return = common.getDisplayName(iDisplayID);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getDisplayName(Int32 iDisplayID)
    {
        string lcl_return = "";
        string sSQL       = "";

        sSQL  = "SELECT displayname ";
        sSQL += " FROM egov_organization_displays ";
        sSQL += " WHERE displayid = " + iDisplayID;

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToString(myReader["displayname"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Boolean orgHasNeighborhoods(Int32 iOrgID)
    {
        Boolean lcl_return = false;

        string sSQL = "";

        sSQL  = "SELECT count(neighborhoodid) as hits ";
        sSQL += " FROM egov_neighborhoods ";
        sSQL += " WHERE orgid = " + iOrgID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (Convert.ToInt32(myReader["hits"]) > 0)
            {
                lcl_return = true;
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Boolean featureIsTurnedOnForPublic(Int32 iOrgID,
                                                     string iFeature)
    {
        Boolean lcl_return = false;

        string sSQL = "";
        string sFeature = "";

        if (iFeature != "")
        {
            sFeature = common.dbSafe(iFeature);
            sFeature = "'" + sFeature + "'";
        }

        sSQL = "SELECT fo.publiccanview ";
        sSQL += " FROM egov_organizations_to_features fo, ";
        sSQL += " egov_organization_features f ";
        sSQL += " WHERE fo.featureid = f.featureid ";
        sSQL += " AND orgid = " + Convert.ToString(iOrgID);
        sSQL += " AND f.feature = " + sFeature;

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (Convert.ToBoolean(myReader["publiccanview"]))
            {
                lcl_return = true;
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Boolean cartHasMerchandise(string iSessionID)
    {
        Boolean lcl_return = false;

        Int32 sItemTypeID = getItemTypeID("merchadise");

        string sSQL = "";
        string sSessionID = "0";

        if (iSessionID != "")
        {
            sSessionID = common.dbSafe(iSessionID);
        }

        sSessionID = "'" + sSessionID + "'";

        sSQL  = "SELECT count(cartid) as hits ";
        sSQL += " FROM egov_class_cart ";
        sSQL += " WHERE isnull(sessionid_csharp,CAST(sessionid AS VARCHAR)) = " + sSessionID;
        sSQL += " AND itemtypeid = " + Convert.ToString(sItemTypeID);

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (Convert.ToInt32(myReader["hits"]) > 0)
            {
                lcl_return = true;
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
        
        return lcl_return;
    }

    public static Int32 getItemTypeID(string iItemType)
    {
        Int32 lcl_return = 0;

        string sItemType = "";
        string sSQL      = "";

        if (iItemType != "")
        {
            sItemType = common.dbSafe(iItemType);
            sItemType = "'" + sItemType + "'";

            sSQL  = "SELECT itemtypeid ";
            sSQL += " FROM egov_item_types ";
            sSQL += " WHERE itemtype = " + sItemType;

            SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
            sqlConn.Open();

            SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
            SqlDataReader myReader;
            myReader = myCommand.ExecuteReader();

            if (myReader.HasRows)
            {
                myReader.Read();

                lcl_return = Convert.ToInt32(myReader["itemtypeid"]);
            }

            myReader.Close();
            sqlConn.Close();
            myReader.Dispose();
            sqlConn.Dispose();
        }

        return lcl_return;
    }

    public static string getItemType(Int32 iItemTypeID)
    {
        string lcl_return = "";
        string sSQL = "";

        sSQL  = "SELECT itemtype ";
        sSQL += " FROM egov_item_types ";
        sSQL += " WHERE itemtypeid = " + iItemTypeID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToString(myReader["itemtype"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getPaymentProcessingRoute( string _OrgId )
    {
        string paymentProcessingRoute = "";
        string paymentGatewayID = getOrgInfo( _OrgId, "OrgPaymentGateway" );

        string sql = "SELECT ISNULL(processingroute,'') AS ProcessingRoute FROM egov_payment_gateways WHERE paymentgatewayid = " + paymentGatewayID;

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            myReader.Read( );
            paymentProcessingRoute = myReader["ProcessingRoute"].ToString( );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );

        return paymentProcessingRoute;
    }
    public static double getPNPFee( string _OrgId, double purchaseAmount, out string ErrorMsg, out bool sPNPFee)
    {
	double retVal = 0;
	ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;

	HttpWebRequest r = ( HttpWebRequest)WebRequest.Create(getOrgFullSite(_OrgId) + "/payment_processors/pnpfeecheck.aspx?chkamount=" + purchaseAmount.ToString());
        r.Method = "Get";
        HttpWebResponse res = (HttpWebResponse)r.GetResponse();
        Stream sr=  res.GetResponseStream();
        StreamReader sre = new StreamReader(sr);

        string s= sre.ReadToEnd();
	

	string status = getPNPResponseValue(s,"status");
	ErrorMsg = getPNPResponseValue(s,"errors");

    	sPNPFee = true;
	if (status != "success")
	{
		sPNPFee = false;
	}
	else
	{
		retVal = Double.Parse(getPNPResponseValue(s,"fee"));
		//retVal = getPNPResponseValue(s,"fee");
	}



	return retVal;
    }
    public static string getPNPResponseValue( string response, string paramName)
    {
	    string retVal = response.Substring(response.IndexOf(paramName)+paramName.Length+1);
	    return retVal.Substring(0,retVal.IndexOf("&") > 0 ? retVal.IndexOf("&") : retVal.Length);
    }

    public static double getCitizenAccountBalance( string _UserId )
    {
        double citizenAccountBalance = 0;

        string sql = "SELECT ISNULL(accountbalance,0) AS CitizenAccountBalance FROM egov_users WHERE userid = " + _UserId;

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            myReader.Read( );
            double.TryParse( myReader["CitizenAccountBalance"].ToString( ), out citizenAccountBalance );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );

        return citizenAccountBalance;
    }

    public static int createInitialPaymentLogEntry( string _OrgId, string _ApplicationSide, string _Feature, string _LogEntry )
    {
        int paymentControlNumber = 0;
        string returnedControllNumber = "0";

        string sql = "INSERT INTO paymentlog ( orgid, applicationside, feature, logentry ) VALUES ( ";
        sql += _OrgId + ", '" + _ApplicationSide + "', '" + _Feature + "', '" + dbSafe( _LogEntry ) + "' )";

        returnedControllNumber = RunInsertStatement( sql );

        int.TryParse( returnedControllNumber, out paymentControlNumber );

        return paymentControlNumber;
    }

    public static void makePaymentLogEntry( ref int _PaymentControlNumber, string _OrgId, string _ApplicationSide, string _Feature, string _LogEntry )
    {
        string sql = "";

        if (_PaymentControlNumber < 1)
        {
            _PaymentControlNumber = createInitialPaymentLogEntry( _OrgId, _ApplicationSide, _Feature, _LogEntry );
            sql = "UPDATE paymentlog SET paymentcontrolnumber = " + _PaymentControlNumber.ToString( ) + " WHERE paymentlogid = " + _PaymentControlNumber.ToString( );
            RunSQLStatement( sql );
        }
        else
        {
            sql = "INSERT INTO paymentlog ( paymentcontrolnumber, orgid, applicationside, feature, logentry ) VALUES ( ";
            sql += _PaymentControlNumber.ToString( ) + ", " + _OrgId + ", '" + _ApplicationSide + "', '" + _Feature + "', '" + dbSafe( _LogEntry ) + "' )";
            RunSQLStatement( sql );
        }
    }

    public static string cleanAndSizeForPayFlowPro( string _Parameter )
    {
        _Parameter = _Parameter.Replace( "\"", "" );
        _Parameter = _Parameter.Replace( "\n", "" );
        _Parameter = _Parameter.Replace( "'", "" );
        _Parameter = _Parameter.Replace( "&", "and" );
        _Parameter = _Parameter.Replace( "=", "is" );
        _Parameter = _Parameter.Replace( "</br>", ", " );
        _Parameter = _Parameter.Replace( "<br />", ", " );
        _Parameter = _Parameter.Replace( "<br>", ", " );
        _Parameter = _Parameter.Replace( ", ,", ", " );
        _Parameter = _Parameter.Trim( );
        _Parameter = Left( _Parameter, 128 );
        return _Parameter;
    }

    public static string cleanAndSizeNotesForPointAndPay( string _Notes )
    {
        _Notes = _Notes.Replace( "\"", "" );
        _Notes = _Notes.Replace( "\n", "" );
        _Notes = _Notes.Replace( "'", "" );
        _Notes = _Notes.Replace( "&", "and" );
        _Notes = _Notes.Replace( "=", "is" );
        _Notes = _Notes.Replace( "</br>", ", " );
        _Notes = _Notes.Replace( "<br />", ", " );
        _Notes = _Notes.Replace( "<br>", ", " );
        _Notes = _Notes.Replace( ", ,", ", " );
        _Notes = _Notes.Trim( );
        _Notes = Left( _Notes, 255 );
        return _Notes;
    }

    public static string getPaymentProcessorResponseValue( string _Results, string _ResponseField )
    {
        string responseValue = "";
        string namePortion = "";
        string valuePortion = "";
        string workingString = "";

        while (_Results.Length > 0)
        {
            if (_Results.Contains( "&" ))
            {
                // try to pull our a name value pair
                workingString = common.Left( _Results, _Results.IndexOf( "&" ) );
            }
            else
            {
                // this case is when there is only one name value pair
                workingString = _Results;
            }

            // the name value pairs are seperated by an '=' sign
            namePortion = common.Left( workingString, workingString.IndexOf( "=" ) );
            valuePortion = workingString.Substring( workingString.IndexOf( "=" ) + 1 );

            if (namePortion.ToUpper( ) == _ResponseField.ToUpper( ))
            {
                // if the field name matches the wanted field then we have a match and are done
                responseValue = valuePortion;
                break;
            }

            if (_Results.Length > workingString.Length)
            {
                // remove the name value pair we just examined from the results
                _Results = _Results.Substring( workingString.Length + 1 );
            }
            else
            {
                // since the results are the same as the last name value pair, then we have searched the entier results with no match
                _Results = "";
            }
        }

        return responseValue;
    }

    public static string savePaymentProcessingError( string _OrgID, string _PaymentProcessor, string _Feature, string _Action, string _ErrorMsg, string _Amount )
    {
        string gatewayErrorId = "0";

        if (_ErrorMsg == "")
            _ErrorMsg = "Unknown Payment Processor Error";

        string amount = _Amount.Replace( ",", "" ).Replace( "$", "" );

        string sql = "INSERT INTO egov_paymentgatewayerrors ( ";
        sql += "orgid, ";
        sql += "paymentgateway, ";
        sql += "action, ";
        sql += "feature, ";
        sql += "errormessage, ";
        sql += "amount";
        sql += " ) VALUES ( ";
        sql += _OrgID + ", ";
        sql += "'" + _PaymentProcessor + "', ";
        sql += "'" + _Action + "', ";
        sql += "'" + _Feature + "', ";
        sql += "'" + dbready_string( _ErrorMsg, 1000 ) + "', ";
        sql += dbready_string( amount, 50 );
        sql += " )";

        gatewayErrorId = RunInsertStatement( sql );

        return gatewayErrorId;
    }

    public static int getPaymentLocationId( )
    {
        int paymentLocationId = 0;

        string sql = "SELECT paymentlocationid FROM egov_paymentlocations WHERE ispublicmethod = 1";

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            myReader.Read( );
            int.TryParse( myReader["paymentlocationid"].ToString( ), out paymentLocationId );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );

        return paymentLocationId;
    }

    public static int getPaymentTypeId( string _OrgId, string _PaymentTypeName )
    {
        int paymentTypeId = 0;

        string sql = "SELECT T.paymenttypeid FROM egov_paymenttypes T, egov_organizations_to_paymenttypes O ";
        sql += "WHERE T.paymenttypename = '" + _PaymentTypeName + "' AND T.paymenttypeid = O.paymenttypeid AND O.orgid = " + _OrgId;

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            myReader.Read( );
            int.TryParse( myReader["paymenttypeid"].ToString( ), out paymentTypeId );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );

        return paymentTypeId;
    }

    public static int getPaymentAccountId( string _OrgId, string _PaymentTypeId )
    {
        int paymentAccountId = 0;

        string sql = "SELECT ISNULL(accountid,0) AS accountid FROM egov_organizations_to_paymenttypes ";
        sql += "WHERE orgid = " + _OrgId + " AND paymenttypeid = " + _PaymentTypeId;

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            myReader.Read( );
            int.TryParse( myReader["accountid"].ToString( ), out paymentAccountId );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );

        return paymentAccountId;
    }

    public static int getJournalEntryTypeId( string _JournalEntryType )
    {
        int journalEntryTypeId = 0;

        string sql = "SELECT journalentrytypeid FROM egov_journal_entry_types WHERE journalentrytype = '" + _JournalEntryType + "'";

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            myReader.Read( );
            int.TryParse( myReader["journalentrytypeid"].ToString( ), out journalEntryTypeId );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );

        return journalEntryTypeId;
    }

    public static string makeTheJournalEntry( string _OrgId, string _PaymentLocationId, string _BuyerUserId, double _TotalPurchaseAmount, string _JournalEntryTypeId, string _Notes, string _IsForRentals, string _ReservationId )
    {
        string journalEntryId = "0";

        string sql = "INSERT INTO egov_class_payment ( paymentdate, paymentlocationid, orgid, adminlocationid, userid, adminuserid, paymenttotal, journalentrytypeid, notes, isforrentals, reservationid ) VALUES ( ";
        sql += "dbo.GetLocalDate(" + _OrgId + ",GetDate()), " + _PaymentLocationId + ", " + _OrgId + ", 0, " + _BuyerUserId + ", 0, " + _TotalPurchaseAmount.ToString( "F2" );
        sql += ", " + _JournalEntryTypeId + ", '" + dbSafe( _Notes ) + "', " + _IsForRentals + ", " + _ReservationId + " )";

        journalEntryId = RunInsertStatement( sql );

        return journalEntryId;
    }

    public static Int32 getCartUserID(string iSessionID)
    {
        Int32 lcl_return = 0;

        string sSQL = "";
        string sSessionID = "";

        if (iSessionID != "")
        {
            sSessionID = common.dbSafe(iSessionID);
        }

        sSessionID = "'" + sSessionID + "'";

        sSQL  = "SELECT TOP 1 userid ";
        sSQL += " FROM egov_class_cart ";
        sSQL += " WHERE isnull(sessionid_csharp, CAST(sessionid AS VARCHAR)) = " + sSessionID;

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToInt32(myReader["userid"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string saveTransactionDetails( string _PaymentId, string _LedgerEntryId, string _PaymentTypeId, double _Amount, string _Status, string _CheckNo, string _AccountId, string _AuthCode, string _PNRef, string _RespMsg, string _OrderNumber, string _SVA, string _FeeAmount )
    {
        string paymentInformationId = "0";

        if (_AuthCode.ToUpper( ) != "NULL")
            _AuthCode = "'" + _AuthCode + "'";

        if (_PNRef.ToUpper( ) != "NULL")
            _PNRef = "'" + _PNRef + "'";

        if (_RespMsg.ToUpper( ) != "NULL")
            _RespMsg = "'" + _RespMsg + "'";

        if (_SVA.ToUpper( ) != "NULL")
            _SVA = "'" + _SVA + "'";

        if (_CheckNo.ToUpper( ) != "NULL")
            _CheckNo = "'" + _CheckNo + "'";

        //if (_OrderNumber.ToUpper( ) != "NULL")
        //    _OrderNumber = "'" + _OrderNumber + "'"; // this is an int in the table so '' is probably not needed

        string sql = "INSERT INTO egov_verisign_payment_information ( paymentid, ledgerid, paymenttypeid, amount, paymentstatus, checkno, citizenuserid, ";
        sql += "authorizationcode, paymentreferenceid, paymentmessage, sva, processingfee, ordernumber ) VALUES ( ";
        sql += _PaymentId + ", " + _LedgerEntryId + ", " + _PaymentTypeId + ", " + _Amount.ToString( "F2" ) + ", '" + _Status + "', " + _CheckNo + ", ";
        sql += _AccountId + ", " + _AuthCode + ", " + _PNRef + ", " + _RespMsg + ", " + _SVA + ", " + _FeeAmount + ", " + _OrderNumber + " )";

        paymentInformationId = common.RunInsertStatement( sql );

        return paymentInformationId;
    }

    public static void setCitizenAccountBalance( string _UserId, string _NewAccountBalance )
    {
        string sql = "UPDATE egov_users SET accountbalance = " + _NewAccountBalance + " WHERE userid = " + _UserId;

        RunSQLStatement( sql );
    }

    public static string getCitizenEmailAddress( string _UserId )
    {
        string emailAddress = "";

        string sql = "SELECT ISNULL(useremail,'') AS useremail FROM egov_users WHERE userid = " + _UserId;

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            myReader.Read( );
            emailAddress = myReader["useremail"].ToString( );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );

        return emailAddress;
    }

    public static string getFamilyMemberName( string _FamilyMemberId )
    {
        string name = "";

        string sql = "SELECT ISNULL(U.userfname,'') AS userfname, ISNULL(U.userlname,'') AS userlname ";
        sql += "FROM egov_users U, egov_familymembers F WHERE U.userid = F.userid AND F.familymemberid = " + _FamilyMemberId;

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            myReader.Read( );
            name = myReader["userfname"].ToString( ) + " " + myReader["userlname"].ToString( );
            name = name.Trim( );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );

        return name;
    }

    public static string getUserName(string iUserID)
    {
        string lcl_return = "";
        string sSQL       = "";
        string sUserID    = "";

        if(iUserID != "")
        {
            sUserID = dbSafe(iUserID);
        }

        sSQL  = "SELECT userfname, ";
        sSQL += " userlname ";
        sSQL += " FROM egov_users ";
        sSQL += " WHERE userid = " + sUserID;

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (Convert.ToString(myReader["userfname"]).Trim() != "" || Convert.ToString(myReader["userlname"]).Trim() != "")
            {
                if (Convert.ToString(myReader["userfname"]).Trim() != "")
                {
                    lcl_return = Convert.ToString(myReader["userfname"]).Trim();
                }

                if (Convert.ToString(myReader["userlname"]).Trim() != "")
                {
                    if (lcl_return != "")
                    {
                        lcl_return += " " + Convert.ToString(myReader["userlname"]).Trim();
                    }
                }

            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static double getCitizenAccountAmount(Int32 iUserID)
    {
        double lcl_return = 0.0000;

        string sSQL = "";

        sSQL  = "SELECT isnull(accountbalance, 0.0000) as accountbalance ";
        sSQL += " FROM egov_users ";
        sSQL += " WHERE userid = " + Convert.ToString(iUserID);

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToDouble(myReader["accountbalance"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static void dtb_debug(string iValue)
    {
        string sSQL = "";
        string sValue = "";

        if (iValue != "")
        {
            sValue = iValue;
            sValue = common.dbSafe(sValue);
        }

        sValue = "'" + sValue + "'";

        sSQL = "INSERT INTO my_table_dtb(notes) VALUES (" + sValue + ")";

        common.RunSQLStatement(sSQL);
    }

    public static Boolean paymentGatewayRequiresFeeCheck(Int32 iOrgID)
    {
        Boolean lcl_return = false;

        string sSQL = "";

        sSQL  = "SELECT requiresfeecheck ";
        sSQL += " FROM egov_payment_gateways g, ";
        sSQL += " organizations o ";
        sSQL += " WHERE g.paymentgatewayid = o.orgpaymentgateway ";
        sSQL += " AND o.orgid = " + iOrgID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (Convert.ToBoolean(myReader["requiresfeecheck"]))
            {
                lcl_return = true;
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Boolean citizenPaysFee(Int32 iOrgID)
    {
        Boolean lcl_return = false;

        string sSQL = "";

        sSQL  = "SELECT citizenpaysfee ";
        sSQL += " FROM organizations ";
        sSQL += " WHERE orgid = " + iOrgID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (Convert.ToBoolean(myReader["citizenpaysfee"]))
            {
                lcl_return = true;
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Boolean getPNPFee(double iPurchaseAmount,
                                    double iFeeAmount,
                                    out double sFeeAmount,
                                    out string sErrorMsg)
    {
        Boolean lcl_return = false;

        string sParmList       = "";
        string sPurchaseAmount = "";

        sFeeAmount = 0.00;
        sErrorMsg  = "";

        if (iPurchaseAmount != null)
        {
            sPurchaseAmount = string.Format("0:0,0.00}",iPurchaseAmount);
        }

        if (iFeeAmount != null)
        {
            sFeeAmount = iFeeAmount;
        }

        sParmList = "chkamount=" + sPurchaseAmount;

        //check in processpayment.aspx for PostDataToProcessor function

        return lcl_return;
    }

    public static string buildStateDropDownOptions(string iStateOrg,
                                         string iStateCitizen)
    {
        string lcl_return     = "";
        string sSQL           = "";
        string sStateOrg      = "";
        string sStateCitizen  = "";
        string sStateValue    = "";
        string sStateSelected = "";
        string sStateCode     = "";

        if (iStateCitizen != "")
        {
            sStateValue = iStateCitizen;
        }
        else
        {
            sStateValue = iStateOrg;
        }

        sStateValue = sStateValue.ToUpper();
        sStateValue = common.dbSafe(sStateValue);

        sSQL  = "SELECT statecode ";
        sSQL += " FROM states ";
        sSQL += " ORDER BY statecode";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while(myReader.Read())
            {
                sStateCode     = Convert.ToString(myReader["statecode"]).ToUpper();
                sStateSelected = "";

                if(sStateValue == sStateCode)
                {
                    sStateSelected = " selected=\"selected\"";
                }

                lcl_return += "<option value=\"" + sStateCode + "\"" + sStateSelected + ">";
                lcl_return +=    sStateCode;
                lcl_return += "</option>";
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string buildCreditCardOptions(Int32 iOrgID)
    {
        string lcl_return = "";
        string sSQL       = "";
        string sCardType  = "";

        sSQL  = "SELECT c.creditcard ";
        sSQL += " FROM creditcards c, ";
        sSQL += " egov_organizations_to_creditcards o ";
        sSQL += " WHERE o.creditcardid = c.creditcardid ";
        sSQL += " AND orgid = " + iOrgID.ToString();
        sSQL += " ORDER BY creditcard";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                sCardType = Convert.ToString(myReader["creditcard"]);

                lcl_return += "<option value=\"" + sCardType.ToLower() + "\">" + sCardType + "</option>";
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string buildMonthOptions()
    {
        Int32 i = 1;

        string lcl_return = "";
        string sMonth = "";

        while (i < 13)
        {
            sMonth = i.ToString();

            if (i < 10)
            {
                sMonth = "0" + sMonth;
            }

            lcl_return += "<option value=\"" + sMonth + "\">" + sMonth + "</option>";

            i = i + 1;
        }

        return lcl_return;
    }

    public static string buildYearOptions()
    {
        Int32 i = 0;
        
        string lcl_return    = "";
        string sCurrentYear  = string.Format("{0:YY}", Convert.ToString(DateTime.Now.Year));
        string sYearOption   = "";
        string sYearValue = "";
        string sSelectedYear = "";

        while (i < 11)
        {
            sYearOption = string.Format("{0:YY}", Convert.ToString(DateTime.Now.Year + i));
            sYearValue = Right( sYearOption, 2 );
            sSelectedYear = "";

            if (sCurrentYear == sYearOption)
            {
                sSelectedYear = " selected=\"selected\"";
            }

            lcl_return += "<option value=\"" + sYearValue + "\"" + sSelectedYear + ">";
            lcl_return +=    sYearOption;
            lcl_return += "</option>";

            i = i + 1;
        }

        return lcl_return;
    }

    public static Int32 saveProcessingPaymentErrorMsg(string iProcessingPath,
                                                      Int32 iOrgID,
                                                      Int32 iUserID,
                                                      Int32 iGatewayErrorID,
                                                      string iErrorMsg,
                                                      string iPNREF,
                                                      string iResult,
                                                      string iRESPMSG,
                                                      double iAmount,
                                                      string iOrderNumber,
                                                      string iSVA,
                                                      string iAUTHCODE,
                                                      string iSessionID,
                                                      string iPaymentName,
                                                      string iCardNumber,
                                                      string iSJName,
                                                      string iStreetAddress,
                                                      string iCity,
                                                      string iState,
                                                      string iZip)
    {
        Int32 lcl_return       = 0;
        Int32 sOrgID           = iOrgID;
        Int32 sUserID          = iUserID;
        Int32 sGatewayErrorID  = iGatewayErrorID;

        string sSQL            = "";
        string sErrorMsg       = "";
        string sPNREF          = "";
        string sResult         = "";
        string sRESPMSG        = "";
        string sAmount         = "";
        string sOrderNumber    = "";
        string sProcessingPath = "PAYPAL";
        string sSVA            = "";
        string sAUTHCODE       = "";
        string sSessionID      = "";
        string sPaymentName    = "";
        string sCardNumber     = "";
        string sSJName         = "";
        string sStreetAddress  = "";
        string sCity           = "";
        string sState          = "";
        string sZip            = "";

        try
        {
            sOrgID = Convert.ToInt32(iOrgID);
        }
        catch
        {
            sOrgID = 0;
        }

        try
        {
            sGatewayErrorID = Convert.ToInt32(iGatewayErrorID);
        }
        catch
        {
            sGatewayErrorID = 0;
        }

        try
        {
            sUserID = Convert.ToInt32(iUserID);
        }
        catch
        {
            sUserID = 0;
        }

        if (iErrorMsg != "")
        {
            sErrorMsg = common.dbSafe(iErrorMsg);
        }

        if (iPNREF != "")
        {
            sPNREF = common.dbSafe(iPNREF);
        }

        if (iResult != "")
        {
            sResult = common.dbSafe(iResult);
        }

        if (iRESPMSG != "")
        {
            sRESPMSG = common.dbSafe(iRESPMSG);
        }

        if (Convert.ToString(iAmount) != "")
        {
            sAmount = Convert.ToString(iAmount);
            sAmount = common.dbSafe(sAmount);
        }

        if (iOrderNumber != "")
        {
            sOrderNumber = common.dbSafe(iOrderNumber);
        }

        if (iProcessingPath != "")
        {
            sProcessingPath = iProcessingPath.ToUpper();
            sProcessingPath = common.dbSafe(sProcessingPath);
        }

        if (iSVA != "")
        {
            sSVA = common.dbSafe(iSVA);
        }

        if (iAUTHCODE != "")
        {
            sAUTHCODE = common.dbSafe(iAUTHCODE);
        }

        if (iSessionID != "")
        {
            sSessionID = common.dbSafe(iSessionID);
        }

        if (iPaymentName != "")
        {
            sPaymentName = common.dbSafe(iPaymentName);
        }

        if (iCardNumber != "")
        {
            sCardNumber = common.dbSafe(iCardNumber);
        }

        if (iSJName != "")
        {
            sSJName = common.dbSafe(iSJName);
        }

        if (iStreetAddress != "")
        {
            sStreetAddress = common.dbSafe(iStreetAddress);
        }

        if (iCity != "")
        {
            sCity = common.dbSafe(iCity);
        }

        if (iState != "")
        {
            sState = common.dbSafe(iState);
        }

        if (iZip != "")
        {
            sZip = common.dbSafe(iZip);
        }

        sErrorMsg       = "'" + sErrorMsg       + "'";
        sPNREF          = "'" + sPNREF          + "'";
        sResult         = "'" + sResult         + "'";
        sRESPMSG        = "'" + sRESPMSG        + "'";
        sAmount         = "'" + sAmount         + "'";
        sOrderNumber    = "'" + sOrderNumber    + "'";
        sProcessingPath = "'" + sProcessingPath + "'";
        sSVA            = "'" + sSVA            + "'";
        sAUTHCODE       = "'" + sAUTHCODE       + "'";
        sSessionID      = "'" + sSessionID      + "'";
        sPaymentName    = "'" + sPaymentName    + "'";
        sCardNumber     = "'" + sCardNumber     + "'";
        sSJName         = "'" + sSJName         + "'";
        sStreetAddress  = "'" + sStreetAddress  + "'";
        sCity           = "'" + sCity           + "'";
        sState          = "'" + sState          + "'";
        sZip            = "'" + sZip            + "'";

        sSQL = "INSERT INTO egov_paymentprocessingerrors (";
        sSQL += "orgid,";
        sSQL += "userid, ";
        sSQL += "gatewayerrorid,";
        sSQL += "errorMsg, ";
        sSQL += "processingpath, ";
        sSQL += "PNREF, ";
        sSQL += "Result, ";
        sSQL += "RESPMSG, ";
        sSQL += "amount, ";
        sSQL += "ordernumber, ";
        sSQL += "SVA, ";
        sSQL += "AUTHCODE, ";
        sSQL += "sessionid, ";
        sSQL += "paymentname, ";
        sSQL += "cardnumber, ";
        sSQL += "sjname, ";
        sSQL += "streetaddress, ";
        sSQL += "city, ";
        sSQL += "state, ";
        sSQL += "zip ";
        sSQL += ") VALUES (";
        sSQL += sOrgID.ToString()          + ", ";
        sSQL += sUserID.ToString()         + ", ";
        sSQL += sGatewayErrorID.ToString() + ", ";
        sSQL += sErrorMsg                  + ", ";
        sSQL += sProcessingPath            + ", ";
        sSQL += sPNREF                     + ", ";
        sSQL += sResult                    + ", ";
        sSQL += sRESPMSG                   + ", ";
        sSQL += sAmount                    + ", ";
        sSQL += sOrderNumber               + ", ";
        sSQL += sSVA                       + ", ";
        sSQL += sAUTHCODE                  + ", ";
        sSQL += sSessionID                 + ", ";
        sSQL += sPaymentName               + ", ";
        sSQL += sCardNumber                + ", ";
        sSQL += sSJName                    + ", ";
        sSQL += sStreetAddress             + ", ";
        sSQL += sCity                      + ", ";
        sSQL += sState                     + ", ";
        sSQL += sZip;
        sSQL += ")";

        lcl_return = Convert.ToInt32(common.RunInsertStatement(sSQL));

        return lcl_return;
    }

    public static string getGatewayErrorMsg(Int32 iOrgID,
                                            Int32 iGatewayErrorID)
    {
        Int32 sOrgID = 0;
        Int32 sGatewayErrorID = 0;

        string lcl_return = "";
        string sSQL = "";

        try
        {
            sOrgID = Convert.ToInt32(iOrgID);
        }
        catch
        {
            sOrgID = 0;
        }

        try
        {
            sGatewayErrorID = Convert.ToInt32(iGatewayErrorID);
        }
        catch
        {
            sGatewayErrorID = 0;
        }

        sSQL = "SELECT isnull(errormessage, '') as errormessage ";
        sSQL += " FROM egov_paymentgatewayerrors ";
        sSQL += " WHERE orgid = " + sOrgID.ToString();
        sSQL += " AND gatewayerrorid = " + sGatewayErrorID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();
            {
                lcl_return = Convert.ToString(myReader["errormessage"]);
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getProcessingPaymentErrorMsg(Int32 iOrgID,
                                                      Int32 iProcessingErrorID)
    {
        Int32 sGatewayErrorID = 0;

        string lcl_return = "";
        string sSQL = "";
        string sProcessErrorMsg = "";
        string sGatewayErrorMsg = "";

        sSQL = "SELECT gatewayerrorid, ";
        sSQL += " isnull(errorMsg, '') as errorMsg ";
        sSQL += " FROM egov_paymentprocessingerrors ";
        sSQL += " WHERE orgid = " + iOrgID.ToString();
        sSQL += " AND processingerrorid = " + iProcessingErrorID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();
            {
                sGatewayErrorID  = Convert.ToInt32(myReader["gatewayerrorid"]);
                sProcessErrorMsg = Convert.ToString(myReader["errorMsg"]);
                sGatewayErrorMsg = common.getGatewayErrorMsg(iOrgID, sGatewayErrorID);

                if (sProcessErrorMsg.ToLower() == "processdeclinedtransaction")
                {
                    sProcessErrorMsg = common.buildProcessDeclinedMsg(iOrgID,
                                                                      iProcessingErrorID);
                }

                if (sProcessErrorMsg != "")
                {
                    lcl_return = sProcessErrorMsg;
                }

                if (sGatewayErrorMsg != "")
                {
                    if (lcl_return != "")
                    {
                        lcl_return += "<br /><br />";
                        lcl_return += sGatewayErrorMsg;
                    }
                }
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string buildProcessDeclinedMsg(Int32 iOrgID,
                                                 Int32 iProcessingErrorID)
    {
        Int32 sPaymentGatewayID = Convert.ToInt32(getOrgInfo(Convert.ToString(iOrgID),
                                                             "OrgPaymentGateway"));
        Int32 sUserID         = 0;
        Int32 sGatewayErrorID = 0;

        string lcl_return      = "";
        string sSQL            = "";
        string sRelPath        = "../";
        string sErrorMsg       = "";
        string sPNREF          = "";
        string sResult         = "";
        string sRESPMSG        = "";
        string sAmount         = "";
        string sOrderNumber    = "";
        string sDisplayPNREF   = "";
        string sProcessingPath = "";
        string sSVA            = "";
        string sAUTHCODE       = "";
        string sSessionID      = "";
        string sPaymentName    = "";
        string sCardNumber     = "";
        string sSJName         = "";
        string sStreetAddress  = "";
        string sCity           = "";
        string sState          = "";
        string sZip            = "";
        string sShowClassNames = "";
        string sPaymentGatewayName = getPaymentGatewayName(sPaymentGatewayID);
        string sPaymentIMG         = getPaymentImage(sPaymentGatewayID,
                                                     sRelPath);

        sSQL  = "SELECT userid, ";
        sSQL += " gatewayerrorid, ";
        sSQL += " errorMsg, ";
        sSQL += " PNREF, ";
        sSQL += " result, ";
        sSQL += " RESPMSG, ";
        sSQL += " amount, ";
        sSQL += " ordernumber, ";
        sSQL += " processingPath, ";
        sSQL += " SVA, ";
        sSQL += " AUTHCODE, ";
        sSQL += " sessionid, ";
        sSQL += " paymentname, ";
        sSQL += " cardnumber, ";
        sSQL += " sjname, ";
        sSQL += " streetaddress, ";
        sSQL += " city, ";
        sSQL += " state, ";
        sSQL += " zip ";
        sSQL += " FROM egov_paymentprocessingerrors ";
        sSQL += " WHERE processingerrorid = " + iProcessingErrorID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();
            {
                sUserID         = Convert.ToInt32(myReader["userid"]);
                sGatewayErrorID = Convert.ToInt32(myReader["gatewayerrorid"]);
                sErrorMsg       = Convert.ToString(myReader["errorMsg"]);
                sProcessingPath = Convert.ToString(myReader["processingPath"]).ToUpper();
                sPNREF          = Convert.ToString(myReader["PNREF"]);
                sResult         = Convert.ToString(myReader["result"]);
                sRESPMSG        = Convert.ToString(myReader["RESPMSG"]);
                sAmount         = Convert.ToString(myReader["amount"]);
                sOrderNumber    = Convert.ToString(myReader["ordernumber"]);
                sSVA            = Convert.ToString(myReader["sva"]);
                sAUTHCODE       = Convert.ToString(myReader["AUTHCODE"]);
                sSessionID      = Convert.ToString(myReader["sessionid"]);
                sPaymentName    = Convert.ToString(myReader["paymentname"]);
                sCardNumber     = Convert.ToString(myReader["cardnumber"]);
                sSJName         = Convert.ToString(myReader["sjname"]);
                sStreetAddress  = Convert.ToString(myReader["streetaddress"]);
                sCity           = Convert.ToString(myReader["city"]);
                sState          = Convert.ToString(myReader["state"]);
                sZip            = Convert.ToString(myReader["zip"]);

                if (sPNREF.Trim() != "")
                {
                    sDisplayPNREF = "Payment Reference Number: " + sPNREF + "<br />";
                }
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        lcl_return  = "<div>";

        //BEGIN: Branding ---------------------------------------------------------------
        if (sPaymentIMG != "../" && sPaymentIMG != "")
        {
            sPaymentIMG = "<img src=\"" + sPaymentIMG + "\" border=\"0\" /><br /><br />";
        }

        lcl_return += "<strong>" + sPaymentGatewayName + "</strong> ";
        lcl_return += "has routed, processed, and secured your payment information.";
        //END: Branding -----------------------------------------------------------------

        lcl_return += "<div id=\"processDeclinedDetails\">";
        lcl_return += "  <div id=\"processDeclinedReason\">";
        lcl_return += "    Your credit card purchase was declined for the following reason:<br />";
        lcl_return +=      sDisplayPNREF;
        lcl_return += "  </div>";
        lcl_return += "  <div style=\"font-weight: bold\">Description: (" + sResult + ") - " + sRESPMSG + "</div>";
        lcl_return += "  <table border=\"0\" class=\"processDeclinedTable\">";


        //BEGIN: Transaction Result Details ---------------------------------------------
        lcl_return += "    <tr>";
        lcl_return += "        <td colspan=\"2\"><hr /></td>";
        lcl_return += "    </tr>";
        lcl_return += "    <tr>";
        lcl_return += "        <td colspan=\"2\" class=\"processDeclinedSectionTitle\">Transaction Details</td>";
        lcl_return += "    </tr>";

        if (sProcessingPath == "POINTANDPOINT")
        {
            lcl_return += "<tr>";
            lcl_return += "    <td>Purchase Amount:</td>";
            lcl_return += "    <td>" + sAmount + "</td>";
            lcl_return += "</tr>";
            lcl_return += "<tr>";
            lcl_return += "    <td>Order Number:</td>";
            lcl_return += "    <td>" + sOrderNumber + "</td>";
            lcl_return += "</tr>";
            lcl_return += "<tr>";
            lcl_return += "    <td>SVA:</td>";
            lcl_return += "    <td>" + sSVA + "</td>";
            lcl_return += "</tr>";
        }
        else  //PayPal
        {
            lcl_return += "<tr>";
            lcl_return += "    <td>Amount:</td>";
            lcl_return += "    <td>" + sAmount + "</td>";
            lcl_return += "</tr>";
            lcl_return += "<tr>";
            lcl_return += "    <td>Reference Number:</td>";
            lcl_return += "    <td>" + sPNREF + "</td>";
            lcl_return += "</tr>";
            lcl_return += "<tr>";
            lcl_return += "    <td>Authorization Code:</td>";
            lcl_return += "    <td>" + sAUTHCODE + "</td>";
            lcl_return += "</tr>";
        }

        lcl_return += "<tr>";
        lcl_return += "    <td colspan=\"2\"><hr /></td>";
        lcl_return += "</tr>";
        //END: Transaction Result Details -----------------------------------------------

        //BEGIN: Product Information ----------------------------------------------------
        sShowClassNames = showClassNames_Text2(sSessionID);

        lcl_return += "<tr>";
        lcl_return += "    <td colspan=\"2\" class=\"processDeclinedSectionTitle\">Product Information</td>";
        lcl_return += "</tr>";
        lcl_return += "<tr>";
        lcl_return += "    <td>Item Number:</td>";
        lcl_return += "    <td>" + sSessionID +"</td>";
        lcl_return += "</tr>";
        lcl_return += "<tr>";
        lcl_return += "    <td>Payment:</td>";
        lcl_return += "    <td>" + sPaymentName +"</td>";
        lcl_return += "</tr>";
        lcl_return += "<tr valign=\"top\">";
        lcl_return += "    <td>Details:</td>";
        lcl_return += "    <td>" + sShowClassNames + "</td>";
        lcl_return += "</tr>";
        lcl_return += "<tr>";
        lcl_return += "    <td colspan=\"2\"><hr /></td>";
        lcl_return += "</tr>";
        //END: Product Information ------------------------------------------------------

        //BEGIN: Credit Card Information ------------------------------------------------
        lcl_return += "<tr>";
        lcl_return += "    <td colspan=\"2\" class=\"processDeclinedSectionTitle\">User Information</td>";
        lcl_return += "</tr>";
        lcl_return += "<tr>";
        lcl_return += "    <td>Credit Card:</td>";
        lcl_return += "    <td>XXXXXXXXXXXX" + sCardNumber.Substring(12) + "</td>";
        lcl_return += "</tr>";
        lcl_return += "<tr>";
        lcl_return += "    <td>Name:</td>";
        lcl_return += "    <td>" + sSJName + "</td>";
        lcl_return += "</tr>";
        lcl_return += "<tr>";
        lcl_return += "    <td>Address:</td>";
        lcl_return += "    <td>" + sStreetAddress + "</td>";
        lcl_return += "</tr>";
        lcl_return += "<tr>";
        lcl_return += "    <td>City:</td>";
        lcl_return += "    <td>" + sCity + "</td>";
        lcl_return += "</tr>";
        lcl_return += "<tr>";
        lcl_return += "    <td>State:</td>";
        lcl_return += "    <td>" + sState + "</td>";
        lcl_return += "</tr>";
        lcl_return += "<tr>";
        lcl_return += "    <td>Zip:</td>";
        lcl_return += "    <td>" + sZip + "</td>";
        lcl_return += "</tr>";
        lcl_return += "<tr>";
        lcl_return += "    <td colspan=\"2\"><hr /></td>";
        lcl_return += "</tr>";
        //END: Credit Card Information --------------------------------------------------

        lcl_return +=   "</table>";
        lcl_return += "</div>";


        lcl_return += "</div>";

        return lcl_return;
    }

    public static string getPaymentImage(Int32 iPaymentGatewayID,
                                         string iRelPath)
    {
        string lcl_return = "";
        string sSQL = "";

        sSQL  = "SELECT isnull(logopath, '') AS logopath ";
        sSQL += " FROM egov_payment_gateways ";
        sSQL += " WHERE paymentgatewayid = " + iPaymentGatewayID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();
            {
                if (Convert.ToString(myReader["logopath"]) != "")
                {
                    lcl_return  = iRelPath;
                    lcl_return += Convert.ToString(myReader["logopath"]);
                }
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;

    }

    public static string getPaymentGatewayName(Int32 iPaymentGatewayID)
    {
        string lcl_return = "PayPal";
        string sSQL = "";

        sSQL  = "SELECT admingatewayname ";
        sSQL += " FROM egov_payment_gateways ";
        sSQL += " WHERE paymentgatewayid = " + iPaymentGatewayID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();
            {
                lcl_return = Convert.ToString(myReader["admingatewayname"]);
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string showClassNames_Text2(string iSessionID)
    {
        Int32 sFamilyMemberID = 0;
        Int32 sOptionID       = 0;
        Int32 sItemTypeID     = 0;
        Int32 sQuantity       = 0;

        string lcl_return        = "";
        string sSessionID        = "";
        string sSQL              = "";
        string sItemType         = "";
        string sClassName        = "";
        string sStartDate        = "";
        string sFamilyMemberName = "";

        if (iSessionID != "")
        {
            sSessionID = dbSafe(iSessionID);
        }

        sSessionID = "'" + sSessionID + "'";

        sSQL  = "SELECT c.classname, ";
        sSQL += " c.startdate, ";
        sSQL += " cc.quantity, ";
        sSQL += " isnull(cc.optionid, 0) as optionid, ";
        sSQL += " isnull(cc.familymemberid, 0) as familymemberid, ";
        sSQL += " cc.itemtypeid ";
        sSQL += " FROM egov_class_cart cc, ";
        sSQL += " egov_class c ";
        sSQL += " WHERE cc.classid = c.classid ";
        sSQL += " AND cc.sessionid_csharp = " + sSessionID;
        sSQL += " ORDER BY cc.dateadded ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while(myReader.Read())
            {
                sItemTypeID     = Convert.ToInt32(myReader["itemtypeid"]);
                sQuantity       = Convert.ToInt32(myReader["quantity"]);
                sOptionID       = Convert.ToInt32(myReader["optionid"]);
                sFamilyMemberID = Convert.ToInt32(myReader["familymemberid"]);

                sClassName = Convert.ToString(myReader["classname"]);
                sStartDate = string.Format("{0:MM/dd/yyyy}", Convert.ToDateTime(myReader["startdate"]));

                sItemType = getItemType(sItemTypeID);
                sItemType = sItemType.ToUpper().Trim();

                switch (sItemType)
                {
                    case "RECREATION ACTIVITY":
                        if (lcl_return != "")
                        {
                            lcl_return += "<br />";
                        }

                        lcl_return += sClassName + "<strong> on </strong>" + sStartDate;

                        //Show qty for ticket events.
                        if (sOptionID == 2)
                        {
                            lcl_return += "&nbsp;&nbsp;<strong>Qty: </strong>" + sQuantity.ToString();
                        }
                        else
                        {
                            if (sOptionID > 0)
                            {
                                sFamilyMemberName = getFamilyMemberName(Convert.ToString(sFamilyMemberID));

                                lcl_return += "&nbsp;&nbsp;<strong>For: </strong>" + sFamilyMemberName;
                            }
                        }

                        break;
                }
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static double getProcessingFee(Int32 iPaymentID)
    {
        double lcl_return = 0.00;

        string sSQL = "";

        sSQL  = "SELECT isnull(processingfee, 0.00) as processingfee ";
        sSQL += " FROM egov_verisign_payment_information ";
        sSQL += " WHERE paymentid = " + iPaymentID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();
            {
                lcl_return = Convert.ToDouble(myReader["processingfee"]);
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Boolean isCLHomePage(Int32 iOrgID)
    {
        Boolean lcl_return = false;

        string sSQL = "";
        string sIsCLHomePage = "N";

        sSQL  = "SELECT 'Y' as isCLHomePage ";
        sSQL += " FROM egov_communitylink ";
        sSQL += " WHERE orgid = " + iOrgID.ToString();
        sSQL += " AND isEgovHomePage = 1 ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();
            {
                sIsCLHomePage = Convert.ToString(myReader["isCLHomePage"]);
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        if (sIsCLHomePage == "Y")
        {
            lcl_return = true;
        }

        return lcl_return;
    }

    public static string decodeUTFString(string iValue)
    {
        SqlString utf8EncodedString;
        string lcl_return = "";

        utf8EncodedString = iValue.ToString().Trim();
        lcl_return        = System.Text.Encoding.UTF8.GetString(utf8EncodedString.GetNonUnicodeBytes());

        return lcl_return;
    }

    public static string getFeaturePublicURL(string iOrgID,
                                         string iFeature)
    {
        Int32 sOrgID = 0;

        string lcl_return = "";
        string sFeature = iFeature;
        string sSQL = "";

        if (iOrgID != null)
        {
            try
            {
                sOrgID = Convert.ToInt32(iOrgID);
            }
            catch
            {
                sOrgID = 0;
            }
        }

        if (sFeature != "")
        {
            sFeature = iFeature.Trim();
            sFeature = dbSafe(sFeature);
        }

        sFeature = "'" + sFeature + "'";

        sSQL = "SELECT isnull(o.publicurl, f.publicurl) as publicurl ";
        sSQL += " FROM egov_organizations_to_features o, ";
        sSQL += " egov_organization_features f ";
        sSQL += " WHERE o.featureid = f.featureid ";
        sSQL += " AND f.feature = " + sFeature;
        sSQL += " AND o.orgid = " + sOrgID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();
            {
                lcl_return = Convert.ToString(myReader["publicurl"]);
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

}

using System;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using System.Text.RegularExpressions; 



/// <summary>
/// This is the common methods for egov on the admin side. You do not need to instantiate this class.
/// Try to keep these in alphabetical order, please.
/// </summary>
public static class common
{
	
    public static string dbSafe(string _Value)
    {
        string sNewString;
        sNewString = _Value.Replace("'", "''");
        sNewString = sNewString.Replace("<", "&lt;");
        return sNewString;
    }

	public static Boolean FolderHasInternalSecurity( string _Folder )
	{
		Boolean IsRestricted = false;
		string Sql = "SELECT ISNULL(DF.accessid,0) AS [issecure] ";
		Sql += "FROM DocumentFolders DF WHERE DF.folderpath = '" + dbSafe( _Folder ) + "' ";
		Sql += "AND DF.OrgID = " + common.getOrgId( );

		SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
		sqlConn.Open( );

		SqlCommand myCommand = new SqlCommand( Sql, sqlConn );
		SqlDataReader myReader;
		myReader = myCommand.ExecuteReader( );

		while ( myReader.Read( ) )
		{
			if ( Int32.Parse( myReader["issecure"].ToString( ) ) > 0 )
				IsRestricted = true;
			else
				IsRestricted = false;
		}

		myReader.Close( );
		sqlConn.Close( );

		return IsRestricted;
	}

	public static Boolean FolderHasRestrictedPublicAccess( string _Folder )
	{
		Boolean IsRestricted = false;
		string Sql = "SELECT ISNULL(DF.citizenaccessid,0) AS [isrestricted] ";
		Sql += "FROM DocumentFolders DF WHERE DF.folderpath = '" + dbSafe( _Folder ) + "' ";
		Sql += "AND DF.OrgID = " + common.getOrgId();

		SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(Sql, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

		while ( myReader.Read( ) )
		{
			if ( Int32.Parse( myReader["isrestricted"].ToString( ) ) > 0 )
				IsRestricted = true;
			else
				IsRestricted = false;
		}

		myReader.Close( );
		sqlConn.Close( );

		return IsRestricted;
	}

    public static string formatPhoneNumber( string _Phone )
    {
        // remove any unwanted formatting and the format it so the phone can call the number
        string phoneNumber = _Phone;
        phoneNumber = phoneNumber.Replace( "(", "" ).Replace( ")", "" ).Replace( "-", "" ).Replace( ".", "" ).Replace( " ", "" );

        if (phoneNumber.Length > 7) {
			string origPN = phoneNumber;
            phoneNumber = "(" + common.Left( phoneNumber, 3 ) + ") " + phoneNumber.Substring( 3, 3 ) + "-" + phoneNumber.Substring( 6,phoneNumber.Length-6 );
			//if (origPN.Length > 10)
			//{
				//phoneNumber += " " + origPN.Substring(10);
			//}
		}
        else
        {
            if (phoneNumber.Length == 7)
                phoneNumber = common.Left( phoneNumber, 3 ) + "-" + common.Right( phoneNumber, 4 );
        }

        return phoneNumber;
    }

	public static string getOrgId()
    {
        string sOrgId = "0"; 
        string sProtocol = "http://";
        if (HttpContext.Current.Request.ServerVariables["HTTPS"].ToUpper() == "ON")
            sProtocol = "https://";

        string sOrgURL = sProtocol + HttpContext.Current.Request.ServerVariables["server_name"].ToLower() + "/" + common.GetVirtualDirectyName(HttpContext.Current.Request.ServerVariables["URL"].ToLower());

        string sSql = "SELECT * FROM Organizations WHERE OrgEgovWebsiteURL = '" + sOrgURL.Replace("https:","http:") + "'";
        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSql, sqlConn);
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

    public static string GetVirtualDirectyName( string _URL )
    {
        string[] NewUrl = _URL.Split('/');
        return NewUrl[1];
    }
    
    public static Boolean IsNumeric( string _Number )
    {
        Int64 Number;
        return Int64.TryParse( _Number, out Number );
    }

    public static bool isValidEmail(string _EmailAddress)
    {
        //              @"^[\w-]+(\.[\w-]+)*@([a-z0-9-]+(\.[a-z0-9-]+)*?\.[a-z]{2,6}|(\d{1,3}\.){3}\d{1,3})(:\d{4})?$"
        string sMatch = @"^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*\.(\w{2}|(com|net|org|edu|mil|gov|biz|us))$";
        Regex myRegEx = new Regex(sMatch, RegexOptions.IgnoreCase);
        Match myMatch = myRegEx.Match(_EmailAddress);
        return myMatch.Success;
    }

    public static string Left(string _Text, int _Length)
    {
        if (_Length < 0)
            throw new ArgumentOutOfRangeException("length", _Length, "length must be > 0");
        else if (_Length == 0 || _Text.Length == 0)
            return "";
        else if (_Text.Length <= _Length)
            return _Text;
        else
            return _Text.Substring(0, _Length);
    }

    public static string Right( string _Text, int _Length )
    {
        if (_Length < 0)
            throw new ArgumentOutOfRangeException( "length", _Length, "length must be > 0" );
        else if (_Length == 0 || _Text.Length == 0)
            return "";
        else if (_Text.Length <= _Length)
            return _Text;
        else
        {
            int startPos = _Text.Length - _Length;
            return _Text.Substring( startPos );
        }
    }

    public static string RunInsertStatement(string _Sql)
    {
        string iNewId = "0";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        SqlCommand sqlCommander = new SqlCommand();
        sqlCommander.Connection = sqlConn;

        sqlConn.Open();

        sqlCommander.CommandText = _Sql + ";SELECT @@IDENTITY";
        iNewId = Convert.ToString(sqlCommander.ExecuteScalar());

        sqlConn.Close();
        sqlCommander.Dispose();

        return iNewId;
    }

    public static void RunSQLStatement(string _Sql)
    {
        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);

        SqlCommand sqlCommander = new SqlCommand();
        sqlCommander.Connection = sqlConn;

        sqlConn.Open();

        sqlCommander.CommandText = _Sql;
        sqlCommander.ExecuteNonQuery();

        sqlConn.Close();
        sqlCommander.Dispose();
    }

    public static void setOrganizationSessionVariables()
    {
        // Initialize the session variables. This prevents problems later
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

	public static Boolean UserCanViewSecureFolder( string _UserId, string _Folder )
	{
		Boolean CanView = false;

		string Sql = "SELECT COUNT(FolderID) AS hits ";
		Sql += "FROM DocumentFolders DF, FeatureAccess FA, UsersGroups UG ";
		Sql += "WHERE DF.folderpath = '" + dbSafe( _Folder ) + "' AND DF.OrgID = " + common.getOrgId( );
		Sql += " AND DF.accessid = FA.accessid AND FA.groupid = UG.groupid AND UG.userid = " + _UserId;

		SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
		sqlConn.Open( );

		SqlCommand myCommand = new SqlCommand( Sql, sqlConn );
		SqlDataReader myReader;
		myReader = myCommand.ExecuteReader( );

		while ( myReader.Read( ) )
		{
			if ( Int32.Parse( myReader["hits"].ToString( ) ) > 0 )
				CanView = true;
			else
				CanView = false;
		}

		myReader.Close( );
		sqlConn.Close( );

		return CanView;
	}
}

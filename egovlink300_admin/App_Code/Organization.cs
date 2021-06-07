using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;

/// <summary>
/// Organization class. You need to instantiate this class
/// </summary>
public class Organization
{
    private int lOrgid;
    private string sResult;
    private string[] aPath;
    private string sConnString; 

    public int iOrgId
    {
        get
        { return lOrgid; }
    }

	public Organization( string sPath)
	{
        // pass in the URL string to find the org that matches
        sConnString = ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString;

        //lOrgid
	}

    public Organization(int iOrgNo)
    {
        // pass in a number to set the orgid to this value
        lOrgid = iOrgNo;
        sConnString = ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString;
    }

    private int GetOrgParameters(string sPath)
    {
        string[] aPath;
        string sResults;

        aPath = sPath.Split('/');
        sResult = aPath[1].Replace("/", "");

        return 0;
    }

}

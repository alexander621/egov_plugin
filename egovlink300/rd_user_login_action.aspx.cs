using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_user_login_action : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //Because ASPX and ASP cookies do NOT work well together, it was decided that a new "userid" cookie be used
        //for ASPX pages.  When a cookie is created in ASP or ASPX and "destroyed" and then a cookie with the same 
        //name is attempted to be created in the other language as second cookie is actually created.  Therefore, TWO
        //cookies with the same name are created and the system doesn't know which one to use.
        //For example, if a user logs in via ASP attempts to log in via ASPX (i.e. as a different user/account)
        //  then there would be two "userid" cookies.
        //        HttpCookie sCookieUserID = new HttpCookie("useridx");
        HttpCookie sCookieUserID = new HttpCookie("userid");

        Int32 sTotal = 0;
        Int32 sOrgID = 0;

        string sResults             = "";
        string sSQL                 = "";
        string sRequestEmail        = "";
        string sRequestPassword     = "";
        string sUserEmail           = "''";
        string sLoginUserPassword   = "x.x";
        string sEncPassword 	  = "x.x";
        string sLoginUserID         = "x.x";
        string sEGovDefaultPage     = "";
        string sSessionRedirectPage = "";
        string sIPAddress           = "";


        if (Request["orgid"] != null)
        {
            try
            {
                sOrgID = Convert.ToInt32(Request["orgid"]);
            }
            catch
            {
                sOrgID = 0;
            }
        }

        string sImgBaseURL           = common.getOrgInfo(sOrgID.ToString(),"orgEgovWebsiteURL");
            sImgBaseURL         = common.getBaseURL(sImgBaseURL);
	Response.Headers.Add("P3P", "CP=This is not a P3P privacy policy!  Read the privacy policy here: " + sImgBaseURL + "/privacy_policy.asp");

        if (Request["email"] != "")
        {
            sRequestEmail = Request["email"];

            sUserEmail = sRequestEmail;
            sUserEmail = common.dbSafe(sUserEmail);
            sUserEmail = "'" + sUserEmail + "'";
        }
	if (String.IsNullOrEmpty(sRequestEmail))
	{
		sRequestEmail = "";
	}

        //--------------------------------------------------------------------
        //BEGIN: Validate required values
        //--------------------------------------------------------------------
        //if (Request["frmsubjecttext"] != null && Request["frmsubjecttext"] != "")
        if (Request["frmsubjecttext"] != "" && sRequestEmail != "")
        {
            sIPAddress = Request.ServerVariables["HTTP_X_FORWARDED_FOR"];

            if (sIPAddress == "")
            {
                sIPAddress = Request.ServerVariables["REMOTE_ADDR"];
            }

            sendSpamFlag(sRequestEmail,
                         Request["password"],
                         Request["frmsubjecttext"],
                         sOrgID,
                         sIPAddress);

            sCookieUserID.Value   = "-1";
            sCookieUserID.Expires = DateTime.Now.AddDays(-1);

            string appName = Page.ResolveUrl("~"); //Gets the application name
            sCookieUserID.Path = appName.Substring(0, appName.Length - 1); //Trims the trailing slash to match the cookie path created by Classic ASP

            Response.Cookies.Add(sCookieUserID);

            sResults = "FAILEDSPAM";
        }

        if (sOrgID == 0)
        {
            sResults = "FAILED";
        }

        if (sRequestEmail.IndexOf("'") > 0)
        {
            //Failed login.  Return to Login Page.
            sCookieUserID.Value   = "-1";
            sCookieUserID.Expires = DateTime.Now.AddDays(-1);

            string appName = Page.ResolveUrl("~"); //Gets the application name
            sCookieUserID.Path = appName.Substring(0, appName.Length - 1); //Trims the trailing slash to match the cookie path created by Classic ASP

            Response.Cookies.Add(sCookieUserID);

            sResults = "FAILED";
        }
        //--------------------------------------------------------------------
        //END: See if we need to redirect the user to the login screen.
        //--------------------------------------------------------------------

        //if (sResults != "FAILED")
        if (! sResults.Contains("FAILED"))
        {
            if (Request["password"] != null)
            {
                sRequestPassword = Request["password"];
            }

            sSQL  = "SELECT userid, ";
            sSQL += " userpassword, password ";
            sSQL += " FROM egov_users ";
            sSQL += " WHERE headofhousehold = 1 AND isdeleted = 0 ";
            sSQL += " AND useremail = " + sUserEmail;
            sSQL += " AND orgid = " + sOrgID.ToString();

            SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
            sqlConn.Open();

            SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
            SqlDataReader myReader;
            myReader = myCommand.ExecuteReader();

            if (myReader.HasRows)
            {
                myReader.Read();

                sTotal = sTotal + 1;

                sLoginUserPassword = Convert.ToString(myReader["userpassword"]);
                sLoginUserID       = Convert.ToString(myReader["userid"]);
		sEncPassword = Convert.ToString(myReader["password"]);
            }

            myReader.Close();
            sqlConn.Close();
            myReader.Dispose();
            sqlConn.Dispose();

            if (sTotal == 0)
            {
                //No such user found, redirect to registration.aspx page.
                //Response.Redirect("register.aspx?email=" + sRequestEmail);
                sResults = "NEWUSER";
            }
            else
            {
		bool bPlainPassMatch = sLoginUserPassword == sRequestPassword && !String.IsNullOrEmpty(sLoginUserPassword);
		bool bEncPassMatch = ValidateUser(sRequestPassword, sEncPassword);

                //Check User's Password
                if (bPlainPassMatch || bEncPassMatch)
                {
                    sEGovDefaultPage = common.getEGovDefaultPage(sOrgID);

                    if (Session["RedirectPage"] != null && Session["RedirectPage"].ToString() != "")
                    {
                        sSessionRedirectPage = Session["RedirectPage"].ToString();

                        sResults = "REDIRECT" + sSessionRedirectPage;
                    }
                    else
                    {
                        if (sEGovDefaultPage != null)
                        {
                            sResults = "REDIRECT" + sEGovDefaultPage.Replace("http:","https:");
                        }
                    }

                    //Add the UserID cookie.
                    sCookieUserID.Value   = sLoginUserID;
                    sCookieUserID.Expires = DateTime.Now.AddHours(8);

                    string appName = Page.ResolveUrl("~"); //Gets the application name
                    sCookieUserID.Path = appName.Substring(0, appName.Length - 1); //Trims the trailing slash to match the cookie path created by Classic ASP

                    Response.Cookies.Add(sCookieUserID);
                }
                else
                {
                    //Failed login.  Return to Login Page.
                    sCookieUserID.Value   = "-1";
                    sCookieUserID.Expires = DateTime.Now.AddDays(-1);

                    string appName = Page.ResolveUrl("~"); //Gets the application name
                    sCookieUserID.Path = appName.Substring(0, appName.Length - 1); //Trims the trailing slash to match the cookie path created by Classic ASP

                    Response.Cookies.Add(sCookieUserID);

                    sResults = "FAILED";
                }
            }
        }

        Response.Write(sResults);
    }

    public bool ValidateUser(string password, string hashedUserPassword)
    {
	    if (!String.IsNullOrEmpty(hashedUserPassword))
	    {
        	string salt = hashedUserPassword.Substring(0, 64);
        	string validHashPw = hashedUserPassword.Substring(64, 64);
	
        	string passHash = sha256Hex(salt + password);

        	if (string.Compare(passHash, validHashPw) == 0)
        	{
            		return true;
        	}
	    }
	    	
	    return false;
    }

        public static string sha256Hex(string _ToHash)
        {

            //returns the SHA256 hash of a string, formatted in hex
            SHA256Managed hash = new SHA256Managed();
            byte[] utf8 = UTF8Encoding.UTF8.GetBytes(_ToHash);

            return bytesToHex(hash.ComputeHash(utf8));

        }
	public static string bytesToHex(byte[] _ToConvert)
        {
            //Converts a byte array to a hex string
            StringBuilder s = new StringBuilder(_ToConvert.Length * 2);
            foreach (byte b in _ToConvert)
            {
                s.Append(b.ToString("x2"));
            }

            return s.ToString();
        }
    

    public void sendSpamFlag(string iEmail,
                             string iPassword,
                             string iHidden,
                             Int32 iOrgID,
                             string iIPAddress)
    {
        string sOrgName       = common.getOrgName(Convert.ToString(iOrgID));
        string sFromEmail     = "noreply@eclink.com";
        string sFromName      = sOrgName + " (E-Gov Website)";
        string sToEmail       = "egovsupport@eclink.com";
        string sSubject       = sOrgName + ": Invalid Public Login Attempt";
        string sEmailBody     = "";
        string sPriority      = "high";
        Boolean sIsHTMLFormat = true;
        string sCcEmail       = "";
        string sBccEmail      = "";

        sEmailBody  = "An attempt was made to log into the public side of " + sOrgName + ". ";
        sEmailBody += "They populated the hidden field and bypassed the JavaScript catch.<br /><br />";
        sEmailBody += "Email Address: " + iEmail     + "<br />";
        sEmailBody += "Password: "      + iPassword  + "<br />";
        sEmailBody += "Hidden Field: "  + iHidden    + "<br />";
        sEmailBody += "IP Address: "    + iIPAddress + "<br />";

        common.sendMessage(iOrgID,
                   sFromEmail,
                   sFromName,
                   sToEmail,
                   sSubject,
                   sEmailBody,
                   sPriority,
                   sIsHTMLFormat,
                   sCcEmail,
                   sBccEmail);
    }
}

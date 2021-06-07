using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_forgot_password_action : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Boolean sIsHTMLFormat = true;

        Int32 sOrgID = 0;

        string sSQL              = "";
        string sOrgName          = "";
        string sRequestUserEmail = "";
        string sUserEmail        = "";
        string sUserPassword     = "";
        //string sFromEmail        = common.getOrgInfo(Convert.ToString(sOrgID), "defaultemail");
        string sFromEmail        = "noreplies@egovlink.com";
        string sFromName         = "";
        string sSubject          = "Password Assistance";
        string sBody             = "";
        string sPriority         = "";
        string sCCEmail          = "";
        string sBCCEmail         = "";
        string lcl_return        = "NOT EXISTS";

        if (Request["orgid"] != null)
        {
            try
            {
                sOrgID   = Convert.ToInt32(Request["orgid"]);
                sOrgName = common.getOrgName(Convert.ToString(sOrgID));
            }
            catch
            {
                sOrgID = 0;
            }
        }

        if (Request["email"] != "" && Request["email"] != null)
        {
            if (! Request["email"].Contains("'"))
            {
                sRequestUserEmail = Request["email"];
                sRequestUserEmail = common.dbSafe(sRequestUserEmail);
                sRequestUserEmail = sRequestUserEmail;
            }
        }

        if(sFromEmail == "")
        {
            sFromEmail = "noreply@eclink.com";
        }

        sFromName = "E-Gov Services " + sOrgName;

        sSQL  = "SELECT useremail, ";
        sSQL += " userpassword ";
        sSQL += " FROM egov_users ";
        sSQL += " WHERE useremail LIKE ('" + sRequestUserEmail + "') ";
        sSQL += " AND orgid = " + sOrgID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            sUserEmail = Convert.ToString(myReader["useremail"]);
            sUserPassword = Convert.ToString(myReader["userpassword"]);
            //sBody = "Your " + sOrgName + " E-Gov Services Password: " + sUserPassword;

            string key = sha256Hex(Membership.GeneratePassword(50, 0));
            string url = common.getOrgFullSite(sOrgID.ToString()) + "/reset_password.asp?key=" + key;

            //save random key and date in the database
	    string sql = "Update egov_users SET pwresetdate = '" + DateTime.Now + "', pwresetkey = '" + key + "' WHERE useremail = '" + sUserEmail + "'";
	    common.RunSQLStatement(sql);



            sBody = "Reset your " + sOrgName + " password.  <a href=\"" + url + "\">Click here to reset your password.</a>  This link is valid for 2 hours.";

            common.sendMessage(sOrgID,
                               sFromEmail,
                               sFromName,
                               sUserEmail,
                               sSubject,
                               sBody,
                               sPriority,
                               sIsHTMLFormat,
                               sCCEmail,
                               sBCCEmail);

            lcl_return = "SENT";
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        Response.Write(lcl_return);
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
}

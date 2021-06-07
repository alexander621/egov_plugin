using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_classes_class_receipt : System.Web.UI.Page
{
    double startCounter = 0.00;

    static string sOrgID   = common.getOrgId();
    static string sOrgName = common.getOrgName(sOrgID);

    protected void Page_PreInit(object sender, EventArgs e)
    {
        // This is the earliest thing the page does, so set the start time here.
        startCounter = DateTime.Now.TimeOfDay.TotalSeconds;
    }

    protected void Page_PreRender(object sender, EventArgs e)
    {
        common.logThePageVisit(startCounter, "class_receipt.aspx", "public");
    }

    protected void Page_Load(object sender, EventArgs e)
    {
    }

    public void displayReceipt(Int32 iOrgID)
    {
        Boolean sShowPaymentDetails                     = false;
        Boolean sHasPaymentFee                          = false;
        Boolean sPaymentGatewayRequiresFeeCheck         = common.paymentGatewayRequiresFeeCheck(iOrgID);
        Boolean sOrgHasFeatureCitizenAccounts           = common.orgHasFeature(Convert.ToString(iOrgID), "citizen accounts");
        Boolean sOrgHasFeatureCustomRegistrationCraigCO = common.orgHasFeature(Convert.ToString(iOrgID), "custom_registration_CraigCO");

        DateTime sPaymentDate;

        double sProcessingFee = 0.00;

        //HttpCookie sUserIDX_Secure = Request.Cookies["useridx_secure"];
        //HttpCookie sUserIDx = Request.Cookies["useridx"];
        HttpCookie sUserIDx = Request.Cookies["userid"];

        Int32 sPaymentID = 0;
        Int32 sUserID = 0;
        Int32 sPaymentTotal = 0;
        Int32 sAdminLocationID = 0;
        Int32 sAdminUserID = 0;
        Int32 sJournalEntryTypeID = 0;

        string sNotes       = "";
        string sEnvironment = ConfigurationManager.AppSettings["environment"];
        string sHttpsStatus = HttpContext.Current.Request.ServerVariables["HTTPS"].ToUpper();
        string sCookieUserID = "";
        string sRequestUserID = "";
        string sShowReceiptHeader = "";
        string sJournalEntryType = "";
        string sAdminUserName = "";
        string sAdminLocation = "";
        string sPayeeAccountInfoTransactionRefundLabel = "Transaction Total";
        string sPayeeAccountInfoTransactionRefundTotal = "";
        string sShowAccountChangeDetails = "";
        string sShowUserInfo = "";
        string sShowRefundPaymentTypes = "";
        string sShowTransactions = "";

        //Depending on the Environment (i.e. DEV or PROD),
        //we need to determine what the userid is.
        if (Request.QueryString["userid"] != "")
        {
            sRequestUserID = Request.QueryString["userid"];
        }

        if (sUserIDx == null || sUserIDx.Value == "" || sUserIDx.Value == "-1")
        {
            sCookieUserID = "";
        }
        else
        {
            sCookieUserID = sUserIDx.Value;
        }

        sUserID = getUserIDByEnvironment(sEnvironment,
                                         sHttpsStatus,
                                         sCookieUserID,
                                         sRequestUserID);

        if (sUserID == 0)
        {
            Response.Redirect("../rd_user_login.aspx");
        }

        try
        {
            sPaymentID = Convert.ToInt32(Request.QueryString["paymentid"]);
        }
        catch
        {
            sPaymentID = 0;
        }
       
        sShowPaymentDetails = getPaymentDetails(iOrgID,
                                                sPaymentID,
                                                sUserID,
                                                out sUserID,
                                                out sPaymentTotal,
                                                out sAdminLocationID,
                                                out sAdminUserID,
                                                out sJournalEntryTypeID,
                                                out sPaymentDate,
                                                out sNotes);

        Response.Write("<div id=\"classReceipt\">");

        if (sShowPaymentDetails)
        {
            //BEGIN: Receipt Header ---------------------------------------------------------------
            sShowReceiptHeader = showReceiptHeader(iOrgID,
                                                   sPaymentID);
            
            sJournalEntryType = getJournalEntryType(sJournalEntryTypeID);

            if (sJournalEntryType.ToUpper() == "PURCHASE")
            {
                if (sPaymentGatewayRequiresFeeCheck)
                {
                    sHasPaymentFee = true;
                    sProcessingFee = common.getProcessingFee(sPaymentID);
                }
            }
            else if (sJournalEntryType.ToUpper() == "REFUND")
            {
                sPayeeAccountInfoTransactionRefundLabel = "Amount Refunded";
            }

            sAdminLocation = classes.getAdminLocation(sAdminLocationID);
            sAdminLocation = common.decodeUTFString(sAdminLocation);

            sAdminUserName = classes.getAdminName(sAdminUserID);
            sAdminUserName = common.decodeUTFString(sAdminUserName);

            Response.Write(sShowReceiptHeader);
            Response.Write("<fieldset class=\"fieldset\">");
            Response.Write("<span class=\"receiptLabel\">Date: </span>" + string.Format("{0:MM/dd/yyyy}", sPaymentDate) + "&nbsp;&nbsp;&nbsp;");
            Response.Write("<span class=\"receiptLabel\">Receipt #: </span>" + sPaymentID + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;");

            if (sAdminLocation != "")
            {
                Response.Write("<span class=\"receiptLabel\">Location: </span>" + sAdminLocation + "&nbsp;&nbsp;");
            }

            if (sAdminUserName != "")
            {
                Response.Write("<span class=\"receiptLabel\">Administrator: </span>" + sAdminUserName);
            }

            Response.Write("</fieldset>");
            //END: Receipt Header -----------------------------------------------------------------

            //BEGIN: Payee Information ------------------------------------------------------------
            //Payee Info
            sShowUserInfo = showUserInfo(iOrgID,
                                         sUserID,
                                         sJournalEntryType);

            //Payee Account Info
            sPayeeAccountInfoTransactionRefundTotal = string.Format("{0:C}",(sPaymentTotal + sProcessingFee));

            if (sOrgHasFeatureCitizenAccounts)
            {
                sShowAccountChangeDetails = showAccountChange(sPaymentID,
                                                              sUserID);
            }

            Response.Write("<fieldset class=\"fieldset\">");
            Response.Write("  <div id=\"receiptPayeeInfo\">" + sShowUserInfo + "</div>");
            Response.Write("  <div id=\"receiptPayeeAccountInfo\">");
            Response.Write("    <div class=\"receiptPayeeAccountInfoTotal\">");
            Response.Write("      <span class=\"receiptAccountInfoTotalLabel\">" + sPayeeAccountInfoTransactionRefundLabel + "</span>&nbsp;");
            Response.Write(       sPayeeAccountInfoTransactionRefundTotal);
            Response.Write("    </div>");
            Response.Write(     sShowAccountChangeDetails);
            Response.Write("  </div>");
            Response.Write("</fieldset>");
            //END: Payee Information --------------------------------------------------------------
            
            //BEGIN: Refund / Payment Types -------------------------------------------------------
            sShowRefundPaymentTypes = showRefundPaymentTypes(iOrgID,
                                                 sPaymentID,
                                                 sJournalEntryType,
                                                 sHasPaymentFee,
                                                 sProcessingFee);

            Response.Write("<fieldset class=\"fieldset\">");
            Response.Write(sShowRefundPaymentTypes);
            Response.Write("</fieldset>");
            //END: Refund / Payment Types ---------------------------------------------------------

            //BEGIN: Transactions -----------------------------------------------------------------
            sShowTransactions = classes.showReceiptTransactions(iOrgID,
                                                                sPaymentID,
                                                                sJournalEntryType,
                                                                sHasPaymentFee,
                                                                sProcessingFee);

            Response.Write("<fieldset class=\"fieldset\" id=\"receiptTransactionsFieldset\">");
            //Response.Write("  <div id=\"receiptTransactionsHeader\">Transactions</div>");
            Response.Write("  <legend>Transactions</legend>");
            Response.Write(sShowTransactions);
            Response.Write("</fieldset>");
            //END: Transactions -------------------------------------------------------------------

        }
        else
        {
            Response.Write("No details could be found for the requested receipt, or you do not have permission to view this receipt.");
        }

        Response.Write("</div>");
    }

    public static Boolean getPaymentDetails(Int32 iOrgID,
                                        Int32 iPaymentID,
                                        Int32 iLoggedInUserID,
                                        out Int32 sUserID,
                                        out Int32 sPaymentTotal,
                                        out Int32 sAdminLocationID,
                                        out Int32 sAdminUserID,
                                        out Int32 sJournalEntryTypeID,
                                        out DateTime sPaymentDate,
                                        out string sNotes)
    {
        Boolean lcl_return = false;
        string sSQL = "";

        sUserID = 0;
        sPaymentTotal = 0;
        sAdminLocationID = 0;
        sAdminUserID = 0;
        sJournalEntryTypeID = 0;
        sPaymentDate = new DateTime();
        sNotes = "";

        sSQL = "SELECT userid, ";
        sSQL += " paymenttotal, ";
        sSQL += " paymentdate, ";
        sSQL += " isnull(adminlocationid, 0) as adminlocationid, ";
        sSQL += " isnull(adminuserid, 0) as adminuserid, ";
        sSQL += " journalentrytypeid, ";
        sSQL += " notes ";
        sSQL += " FROM egov_class_payment ";
        sSQL += " WHERE paymentid = " + iPaymentID.ToString();
        sSQL += " AND orgid = " + iOrgID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            sUserID = Convert.ToInt32(myReader["userid"]);

            if (iLoggedInUserID == sUserID)
            {
                lcl_return = true;
                sPaymentTotal = Convert.ToInt32(myReader["paymenttotal"]);
                sAdminLocationID = Convert.ToInt32(myReader["adminlocationid"]);
                sAdminUserID = Convert.ToInt32(myReader["adminuserid"]);
                sJournalEntryTypeID = Convert.ToInt32(myReader["journalentrytypeid"]);
                sPaymentDate = Convert.ToDateTime(myReader["paymentdate"]);
                sNotes = Convert.ToString(myReader["notes"]);
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Int32 getUserIDByEnvironment(string iEnvironment,
                                               string iHttpsStatus,
                                               string iCookieUserID,
                                               string iRequestUserID)
    {
        Int32 lcl_return = 0;

        string sEnvironment = "PROD";
        string sHttpsStatus = "";

        if (iEnvironment != "")
        {
            sEnvironment = common.dbSafe(iEnvironment);
            sEnvironment = sEnvironment.ToUpper();
        }

        if (iHttpsStatus != "")
        {
            sHttpsStatus = iHttpsStatus.ToUpper();
            sHttpsStatus = common.dbSafe(sHttpsStatus);
        }
/*
        if (sEnvironment == "DEV")
        {
            //NOTE: We never go into "HTTPS" in the DEV environment!
            try
            {
                lcl_return = Convert.ToInt32(iCookieUserID);
            }
            catch
            {
                if (iRequestUserID != "" && iRequestUserID != null)
                {
                    try
                    {
                        lcl_return = Convert.ToInt32(iRequestUserID);
                    }
                    catch
                    {
                        lcl_return = 0;
                    }
                }
            }
        }
        else  //PROD
        {
            if (sHttpsStatus != "ON")
            {
                try
                {
                    lcl_return = Convert.ToInt32(iCookieUserID);
                }
                catch
                {
                    lcl_return = 0;
                }
            }
            else
            {
                if (iRequestUserID != "" && iRequestUserID != null)
                {
                    try
                    {
                        lcl_return = Convert.ToInt32(iRequestUserID);
                    }
                    catch
                    {
                        lcl_return = 0;
                    }
                }
            }
        }
*/

        //We create a new cookie for the "secure" url in Process Payment BEFORE redirecting to class_receipe.aspx.
        //The past method was a "Form - POST" and therefore the userid was passed in the post.
        try
        {
            lcl_return = Convert.ToInt32(iCookieUserID);
        }
        catch
        {
            if (iRequestUserID != "" && iRequestUserID != null)
            {
                try
                {
                    lcl_return = Convert.ToInt32(iRequestUserID);
                }
                catch
                {
                    lcl_return = 0;
                }
            }
        }

        return lcl_return;
    }

    public static string showReceiptHeader(Int32 iOrgID,
                                           Int32 iPaymentID)
    {
        Boolean sOrgHasDisplayReceiptHeader = common.orgHasDisplay(Convert.ToString(iOrgID), "receipt header");

        string lcl_return                  = "";
        string sGetReceiptHeader           = getReceiptHeader(iPaymentID);
        string sGetOrgDisplayReceiptHeader = common.getOrgDisplay(Convert.ToString(iOrgID), "receipt header");
        string sOrgName                    = common.getOrgName(Convert.ToString(iOrgID));

        if (sOrgHasDisplayReceiptHeader)
        {
            sGetOrgDisplayReceiptHeader = common.decodeUTFString(sGetOrgDisplayReceiptHeader);

            lcl_return  = "<div id=\"receiptHeader\">";
            lcl_return += sGetOrgDisplayReceiptHeader;
            lcl_return += "<br /><br >";
            lcl_return += sGetReceiptHeader;
            lcl_return += "</div>";
        }
        else
        {
            lcl_return  = "<h3>" + sOrgName + "&nbsp;" + sGetReceiptHeader + "</h3><br /><br />";
        }

        return lcl_return;
    }

    public static string getReceiptHeader(Int32 iPaymentID)
    {
        string lcl_return = Convert.ToString(iPaymentID);
        string sSQL = "";

        sSQL = "SELECT receipttitle ";
        sSQL += " FROM egov_class_payment p, ";
        sSQL += " egov_journal_entry_types j ";
        sSQL += " WHERE p.journalentrytypeid = j.journalentrytypeid ";
        sSQL += " AND p.paymentid = " + iPaymentID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToString(myReader["receipttitle"]);
            lcl_return = common.decodeUTFString(lcl_return);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getJournalEntryType(Int32 iJournalEntryTypeID)
    {
        string lcl_return = "";
        string sSQL = "";

        sSQL  = "SELECT journalentrytype ";
        sSQL += " FROM egov_journal_entry_types ";
        sSQL += " WHERE journalentrytypeid = " + iJournalEntryTypeID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToString(myReader["journalentrytype"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string showAccountChange(Int32 iPaymentID,
                                           Int32 iUserID)
    {
        double sPriorBalance = 0.00;
        double sAmount = 0.00;
        double sCurrentBalance = 0.00;

        string lcl_return = "";
        string sSQL = "";
        string sEntryType = "credit";
        string sPrefix = "+";

        sSQL  = "SELECT entrytype, ";
        sSQL += " amount, ";
        sSQL += " priorbalance, ";
        sSQL += " plusminus ";
        sSQL += " FROM egov_accounts_ledger ";
        sSQL += " WHERE accountid = " + iUserID.ToString();
        sSQL += " AND paymentid = " + iPaymentID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            sPriorBalance = Convert.ToDouble(myReader["priorbalance"]);
            sAmount       = Convert.ToDouble(myReader["amount"]);
            sEntryType    = Convert.ToString(myReader["entrytype"]);
            sPrefix       = Convert.ToString(myReader["plusminus"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        if (sEntryType.ToUpper() == "CREDIT")
        {
            sCurrentBalance = sPriorBalance + sAmount;
        }
        else
        {
            sCurrentBalance = sPriorBalance - sAmount;
        }

        lcl_return  = "<div class=\"receiptPayeeHeader\">Payee Account Information</div>";
        lcl_return += "<div>";
        lcl_return += "<table id=\"receiptPayeeTable\" width=\"100%\">";
        lcl_return += "  <tr>";
        lcl_return += "      <td>Prior Balance.............................</td>";
        lcl_return += "      <td class=\"receiptPayeeAmounts\">" + string.Format("{0:C}", sPriorBalance) + "</td>";
        lcl_return += "  </tr>";
        lcl_return += "  <tr>";
        lcl_return += "      <td>Change.....................................</td>";
        lcl_return += "      <td class=\"receiptPayeeAmounts\">" + sPrefix + " " + string.Format("{0:C}", sAmount) + "</td>";
        lcl_return += "  </tr>";
        lcl_return += "  <tr class=\"receiptTotalRow\">";
        lcl_return += "      <td>Current Balance......................</td>";
        lcl_return += "      <td class=\"receiptPayeeAmounts\">" + string.Format("{0:C}", sCurrentBalance) + "</td>";
        lcl_return += "  </tr>";
        lcl_return += "</table>";
        lcl_return += "</div>";


        return lcl_return;
    }

    public static string showUserInfo(Int32 iOrgID,
                                      Int32 iUserID,
                                      string iJournalEntryType)
    {
        string lcl_return    = "";
        string sUserType     = classes.getUserResidentType(iUserID);
        string sResidentDesc = "";
        string sEmail = "";
        string sAddress = "";
        string sAddress2 = "";
        string sUserUnit = "";
        string sCity = "";
        string sState = "";
        string sZip = "";
        string sFirstName = "";
        string sLastName = "";
        string sName = "";
        string sUserHomePhone = "";
        string sJournalEntryType = "";
        string sShowUserInfoHeader = "Payee";
        string sShowAddressFields = "";
        string sShowCityStateZip = "";
        string sFamilyEmail = classes.getFamilyEmail(iUserID);

        if(iJournalEntryType != "")
        {
            sJournalEntryType = common.dbSafe(iJournalEntryType);
            sJournalEntryType = sJournalEntryType.ToUpper();
        }

        if (sUserType != "R" && sUserType != "N")
        {
            sUserType = classes.getResidentTypeByAddress(iUserID,
                                                         iOrgID);
        }

        sResidentDesc = classes.getResidentTypeDesc(sUserType);

        classes.getUserInfo(iUserID,
                            out sEmail,
                            out sAddress,
                            out sAddress2,
                            out sUserUnit,
                            out sCity,
                            out sState,
                            out sZip,
                            out sFirstName,
                            out sLastName,
                            out sName,
                            out sUserHomePhone);

        if(sJournalEntryType == "REFUND")
        {
            sShowUserInfoHeader = "Refundee";
        }

        if (sAddress != "")
        {
            sShowAddressFields = sAddress;
        }

        if (sUserUnit != "")
        {
            if (sShowAddressFields != "")
            {
                sShowAddressFields += "&nbsp;&nbsp;";
            }

            sShowAddressFields += sUserUnit;
        }

        if (sAddress2 != "")
        {
            if (sShowAddressFields != "")
            {
                sShowAddressFields += "<br />";
            }

            sShowAddressFields += sAddress2;
        }

        if (sShowAddressFields != "")
        {
            sShowAddressFields += "<br />";
        }

        if (sCity != "")
        {
            sShowCityStateZip = sCity;
        }

        if (sState != "")
        {
            if (sShowCityStateZip != "")
            {
                sShowCityStateZip += ", ";
            }

            sShowCityStateZip += sState;

        }

        if (sZip != "")
        {
            if (sShowCityStateZip != "")
            {
                sShowCityStateZip += " ";
            }

            sShowCityStateZip += sZip;
        }

        lcl_return  = "<div id=\"receiptUserInfo\">";
        lcl_return += "<div class=\"receiptPayeeHeader\">" + sShowUserInfoHeader + "&nbsp;Information</div>";
        lcl_return += "<table border=\"0\">";
        lcl_return += "  <tr>";
        lcl_return += "      <td>&nbsp;</td>";
        lcl_return += "      <td style=\"whitespace: nowrap; font-weight: bold;\">";
        lcl_return +=            sName + "<br />";
        lcl_return +=            sShowAddressFields + "<br />";
        lcl_return +=            sShowCityStateZip;
        lcl_return += "      </td>";
        lcl_return += "  </tr>";
        lcl_return += "  <tr>";
        lcl_return += "      <td>Email:</td>";
        lcl_return += "      <td style=\"whitespace: nowrap;\">";
        lcl_return +=            sFamilyEmail;
        lcl_return += "      </td>";
        lcl_return += "  </tr>";
        lcl_return += "  <tr>";
        lcl_return += "      <td>Phone:</td>";
        lcl_return += "      <td style=\"whitespace: nowrap;\">";
        lcl_return +=            common.formatPhoneNumber(sUserHomePhone);
        lcl_return += "      </td>";
        lcl_return += "  </tr>";
        lcl_return += "</table>";
        lcl_return += "</div>";

        return lcl_return;
    }

    public static string showRefundPaymentTypes(Int32 iOrgID,
                                                Int32 iPaymentID,
                                                string iJournalEntryType,
                                                Boolean iHasPaymentFee,
                                                double iProcessingFee)
    {
        string lcl_return = "";
        string sJournalEntryType = "";

        if (iJournalEntryType != "")
        {
            sJournalEntryType = common.dbSafe(iJournalEntryType);
            sJournalEntryType = sJournalEntryType.ToUpper();
        }

        if (sJournalEntryType == "REFUND")
        {
            lcl_return = showRefundType(iOrgID,
                                        iPaymentID);
        }
        else
        {
            lcl_return = showPaymentTypes(iOrgID,
                                          iPaymentID,
                                          sJournalEntryType,
                                          iHasPaymentFee,
                                          iProcessingFee);
        }

        if(lcl_return != "")
        {
            lcl_return = "<div>" + lcl_return + "</div>";
        }

        return lcl_return;
    }

    public static string showRefundType(Int32 iOrgID,
                                        Int32 iPaymentID)
    {
        Boolean sIsPaymentAccount = false;
        Boolean sIsCCRefund = false;

        double sAmount = 0.00;

        Int32 sAccountID = 0;

        string lcl_return = "";
        string sSQL = "";
        string sRefundPaymentLabel = "";
        string sRefundPaymentValue = "";
        string sRefundPaymentCitizenName = "";

        sSQL = "SELECT isnull(accountid, 0) as accountid, ";
        sSQL += " amount, ";
        sSQL += " priorbalance, ";
        sSQL += " plusminus, ";
        sSQL += " itemid, ";
        sSQL += " ispaymentaccount, ";
        sSQL += " paymenttypeid, ";
        sSQL += " isccrefund ";
        sSQL += " FROM egov_accounts_ledger ";
        sSQL += " WHERE entrytype = 'credit' ";
        sSQL += " AND paymentid = " + iPaymentID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            sIsPaymentAccount = Convert.ToBoolean(myReader["ispaymentaccount"]);
            sIsCCRefund = Convert.ToBoolean(myReader["isccrefund"]);
            sAmount = Convert.ToDouble(myReader["amount"]);
            sAccountID = Convert.ToInt32(myReader["accountid"]);

            if (sAmount == null)
            {
                sAmount = 0.00;
            }

            if (sIsPaymentAccount)
            {
                if (sIsCCRefund)
                {
                    sRefundPaymentLabel = "Refund to Credit Card";
                }
                else
                {
                    //This is a refund voucher
                    sRefundPaymentLabel = classes.getRefundName(iOrgID);
                }

                sRefundPaymentValue = string.Format("{0:C}", sAmount);
            }
            else
            {
                if (sAmount > 0.00)
                {
                    //This is to a citizen account
                    sRefundPaymentLabel = "Citizen Account";
                    sRefundPaymentValue = string.Format("{0:C}", sAmount);

                    sRefundPaymentCitizenName = classes.getUserContactInfo(sAccountID, "userfname");
                    sRefundPaymentCitizenName += " ";
                    sRefundPaymentCitizenName += classes.getUserContactInfo(sAccountID, "userlname");
                    sRefundPaymentCitizenName = sRefundPaymentCitizenName.Trim();

                    if (sRefundPaymentCitizenName != "")
                    {
                        sRefundPaymentCitizenName = "<td>&nbsp;&nbsp;To:&nbsp;" + sRefundPaymentCitizenName + "</td>";
                    }
                }
                else
                {
                    sRefundPaymentLabel = "Removed from the WaitList";
                    sRefundPaymentValue = "&nbsp;";
                }
            }
        }
        else
        {
            sRefundPaymentLabel = "Removed from the WaitList";
            sRefundPaymentValue = "&nbsp;";
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        lcl_return = "<table border=\"0\">";
        lcl_return += "  <tr>";
        lcl_return += "      <td>" + sRefundPaymentLabel + ":&nbsp;</td>";
        lcl_return += "      <td>" + sRefundPaymentValue + "</td>";
        lcl_return +=        sRefundPaymentCitizenName;
        lcl_return += "  </tr>";
        lcl_return += "</table>";

        return lcl_return;
    }

    public static string showPaymentTypes(Int32 iOrgID,
                                          Int32 iPaymentID,
                                          string iJournalEntryType,
                                          Boolean iHasPaymentFee,
                                          double iProcessingFee)
    {
        Boolean sRequiresCheckNo = false;
        Boolean sRequiresCitizenAccount = false;

        double sTotal = 0.00;
        double sAmount = 0.00;

        Int32 sPaymentTypeID = 0;

        string lcl_return = "";
        string sJournalEntryType = "";
        string sSQL = "";
        string sWhereClause = "";
        string sPaymentTypeName = "";
        string sPaymentAmount = "";
        string sPaymentAdditionalInfo = "";
        string sCheckNo = "";
        string sAccountName = "";

        if(iJournalEntryType !="")
        {
            sJournalEntryType = common.dbSafe(iJournalEntryType);
            sJournalEntryType = sJournalEntryType.ToUpper();
        }

        if (sJournalEntryType != "REFUND")
        {
            sWhereClause  = " AND isrefundmethod = 0 ";
            sWhereClause += " AND isrefunddebit = 0 ";
        }

        sSQL  = "SELECT p.paymenttypeid, ";
        sSQL += " p.paymenttypename, ";
        sSQL += " p.requirescheckno, ";
        sSQL += " p.requirescitizenaccount ";
        sSQL += " FROM egov_paymenttypes p, ";
        sSQL +=      " egov_organizations_to_paymenttypes o ";
        sSQL += " WHERE p.paymenttypeid = o.paymenttypeid ";
        sSQL += " AND o.orgid = " + iOrgID.ToString();
        sSQL += sWhereClause;
        sSQL += " ORDER BY p.displayorder ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            lcl_return = "<table border=\"0\">";

            while (myReader.Read())
            {
                sPaymentTypeID          = Convert.ToInt32(myReader["paymenttypeid"]);
                sPaymentTypeName        = Convert.ToString(myReader["paymenttypename"]);
                sRequiresCheckNo        = Convert.ToBoolean(myReader["requirescheckno"]);
                sRequiresCitizenAccount = Convert.ToBoolean(myReader["requirescitizenaccount"]);
                sPaymentAdditionalInfo = "&nbsp;";

                sAmount = classes.getLedgerAmount(iPaymentID,
                                                  sPaymentTypeID);

                sTotal = sTotal - sAmount;

                if (iHasPaymentFee && (sAmount > Convert.ToDouble(0.00)))
                {
                    sTotal = sTotal + iProcessingFee;
                    sPaymentAmount = string.Format("{0:C}", sAmount + iProcessingFee);
                }
                else
                {
                    sPaymentAmount = string.Format("{0:C}", sAmount);
                }

                if (sRequiresCheckNo)
                {
                    sCheckNo = classes.getCheckNo(iPaymentID,
                                                  sPaymentTypeID);

                    if (sCheckNo != "")
                    {
                        sPaymentAdditionalInfo = "&nbsp;&nbsp;Check #: ";

                        if (sAmount > Convert.ToDouble(0.00))
                        {
                            sPaymentAdditionalInfo += sCheckNo;
                        }
                    }

                }

                if (sRequiresCitizenAccount)
                {
                    sAccountName = classes.getAccountName(iPaymentID,
                                                          sPaymentTypeID);

                    if (sAccountName != "")
                    {
                        sPaymentAdditionalInfo += "&nbsp;&nbsp;From: ";

                        if (sAmount > Convert.ToDouble(0.00))
                        {
                            sPaymentAdditionalInfo += sAccountName;
                        }
                    }
                }
                                                  

                lcl_return += "  <tr>";
                lcl_return += "      <td class=\"receiptLabel\">" + sPaymentTypeName + ":&nbsp;</td>";
                lcl_return += "      <td class=\"receiptPayeeAmounts\">" + sPaymentAmount + "</td>";
                lcl_return += "      <td>" + sPaymentAdditionalInfo + "</td>";
                lcl_return += "  </tr>";
            }

            lcl_return += "  <tr class=\"receiptTotalRow\">";
            lcl_return += "      <td>Total:&nbsp;</td>";
            lcl_return += "      <td class=\"receiptPayeeAmounts\">" + string.Format("{0:C}", sTotal) + "</td>";
            lcl_return += "      <td>&nbsp;</td>";
            lcl_return += "  </tr>";
            lcl_return += "</table>";

        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

}

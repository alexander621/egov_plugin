using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_classes_class_paymentform : System.Web.UI.Page
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
        common.logThePageVisit(startCounter, "class_paymentform.aspx", "public");
    }

    protected void Page_Load(object sender, EventArgs e)
    {
    }

    public void displayPaymentForm(Int32 iOrgID)
    {
        Boolean sPNPFee         = true;
        Boolean sApplyTheCredit = false;
        Boolean sCitizenPaysFee = false;
        Boolean sHasPaymentFee  = false;
        Boolean sPaymentGatewayRequiresFeeCheck     = common.paymentGatewayRequiresFeeCheck(iOrgID);
        Boolean sOrgHasFeatureAllowsAccountPayments = common.orgHasFeature(Convert.ToString(iOrgID), "public account payments");
        Boolean sOrgHasFeatureDisplayCVV            = common.orgHasFeature(Convert.ToString(iOrgID), "display cvv");


        double sItemTotal     = 0.00;
        double sAccountCredit = 0.00;
        double sTotalAmount   = 0.00;
        double sFeeAmount     = 0.00;

        Int32 sPurchaserID  = 0;
        Int32 sUserID       = 0;
        Int32 sItemNumber   = 0;
        Int32 sItemQuantity = 1;
        Int32 sTotalItems   = 0;
        Int32 i = 1;

        //string sSessionID = HttpContext.Current.Session.SessionID;
        string sSessionID       = "";
        string sApplyCredit     = "";
        string sItemDescription = Request["ITEM_NAME"];
        string sTaxTable        = "N";  //"Y" = Use Table - "N" Don't Use Table
        string sOrderString     = "";
        string sSerialNumber    = classes.getSerialNumber(5);  //(5) is hardcoded in the public-side classes "verisign_form.asp" as well
        string sCreditCardNum   = "";
        string sExpMonth        = "";
        string sExpYear         = "";
        string sCVSCode         = "";
        string sPhone           = "";
        string sAmount          = "";
        string sOrderNumber     = "";
        string sEmail           = "";
        string sAddress         = "";
        string sAddress2        = "";
        string sUserUnit        = "";
        string sCity            = "";
        string sState           = "";
        string sZip             = "";
        string sFirstName       = "";
        string sLastName        = "";
        string sName            = "";
        string sUserHomePhone   = "";
        string sDetails         = "";
        string sStateOrg        = "";
        string sRosterGrade     = "";
        string sRosterShirtSize = "";
        string sRosterPantsSize = "";
        string sRosterCoachType = "";
        string sRosterVolunteerCoachName      = "";
        string sRosterVolunteerCoachDayPhone  = "";
        string sRosterVolunteerCoachCellPhone = "";
        string sRosterVolunteerCoachEmail     = "";
        string sDisplayItemsToBePurchased     = "";
        string sDisplayItemTotal              = "";
        string sProcessingRoute               = "";
        string sErrorMsg                      = "";
        string sGatewayErrorID                = "";
        string sPaymentProcessingFailureURL   = "";
        string sDisplayStateOptions           = "";
        string sDisplayCardTypesOptions       = "";
        string sDisplayOptionsMonth           = "";
        string sDisplayOptionsYear            = "";
        string sOrgVirtualSiteName            = common.getOrgInfo(Convert.ToString(iOrgID), "orgVirtualSiteName");

        if(Request["sessionID"] != "")
		{
			sSessionID = Request["sessionID"];

            try
            {
                sSessionID = common.dbSafe(sSessionID);
            }
            catch
            {
                sSessionID = Request["sessionID"];
            }
		}

        if(Request.Form["userid"] != "")
        {
            try
            {
                sUserID = Convert.ToInt32(Request.Form["userid"]);
            }
            catch
            {
                sUserID = 0;
            }
        }

        try
        {
            sPurchaserID = Convert.ToInt32(Request["purchaserid"]);
        }
        catch
        {
            sPurchaserID = 0;
        }

        try
        {
            sTotalItems = Convert.ToInt32(Request["totalitem"]);
        }
        catch
        {
            sTotalItems = 0;
        }

        sDetails          = classes.getActivityNosForPayPalComment2(sSessionID);
        sItemTotal        = classes.getCartTotalAmount(sSessionID);
        sDisplayItemTotal = string.Format("{0:#,0.00}", sItemTotal);

        if (Request["applycredit"] != "")
        {
            sApplyCredit = Convert.ToString(Request["applycredit"]);
            sApplyCredit = common.dbSafe(sApplyCredit);
            sApplyCredit = sApplyCredit.ToString().ToLower();
        }

        if (sApplyCredit == "yes")
        {
            sApplyTheCredit = true;
        }

        if (sApplyTheCredit)
        {
            sAccountCredit = common.getCitizenAccountAmount(sPurchaserID);

            if (sAccountCredit >= Convert.ToDouble(0))
            {
                if (sAccountCredit > sItemTotal)
                {
                    sAccountCredit = sItemTotal;
                }
            }

            if (sAccountCredit == 0.00)
            {
                sApplyTheCredit = false;
            }

            if (sOrgHasFeatureAllowsAccountPayments && sApplyTheCredit)
            {
                sTotalAmount = sItemTotal - sAccountCredit;
            }
            else
            {
                sTotalAmount = sItemTotal;
                sApplyTheCredit = false;
            }
        }
        else
        {
            sTotalAmount = sItemTotal;
        }

        if (Request["ITEM_NUMBER"] != "")
        {
            try
            {
                sItemNumber = Convert.ToInt32(Request["ITEM_NUMBER"]);
            }
            catch
            {
                sItemNumber = 0;
            }
        }

        //Check if payment gateway needs a fee check for this page
        if (sPaymentGatewayRequiresFeeCheck)
        {
            sCitizenPaysFee = common.citizenPaysFee(iOrgID);

            if (sCitizenPaysFee)
            {
                sHasPaymentFee   = true;
                sProcessingRoute = common.getPaymentProcessingRoute(Convert.ToString(iOrgID));

                if (sProcessingRoute.ToUpper() == "POINTANDPAY")
                {
                    //Fetch the fee for the amount to be charged.

	    	    sFeeAmount = common.getPNPFee(sOrgID, sTotalAmount, out sErrorMsg, out sPNPFee);


                    if (!sPNPFee)  
                    {
                        //If not successful, store the error, then take them to a page 
                        //to display the error message
                        sGatewayErrorID = common.savePaymentProcessingError(Convert.ToString(iOrgID),
                                                                            sProcessingRoute,
                                                                            "feecheck",
                                                                            "fee check",
                                                                            sErrorMsg,
                                                                            string.Format("{0:#,0.00}",sTotalAmount));

                        sPaymentProcessingFailureURL  = ConfigurationManager.AppSettings["paymenturl"];
                        sPaymentProcessingFailureURL += "/" + sOrgVirtualSiteName;
                        sPaymentProcessingFailureURL += "/rd_classes/rd_processing_failure.aspx?ge=" + sGatewayErrorID;

                        Response.Redirect(sPaymentProcessingFailureURL);
                    }
                }
            }
        }

        sTotalAmount = sTotalAmount + sFeeAmount;

        //If this is Park City (orgid = 37) then show test values
        if (iOrgID == 37)
        {
            sCreditCardNum = "5555555555554444";
            sExpMonth      = "03";
            sExpYear       = "2013";
            sPhone         = "5136814030";
        }

        //Build the OrderString
        sOrderString  = sItemNumber.ToString();
        sOrderString += "-";
        sOrderString += sItemDescription;
        sOrderString += "-";
        sOrderString += sAmount;
        sOrderString += "-";
        sOrderString += sItemQuantity.ToString();
        sOrderString += "-";
        sOrderString += sTaxTable;
        sOrderString += "||";

        //Get the user info
        classes.getUserInfo(sPurchaserID,
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

        Response.Write("<form name=\"paymentForm\" id=\"paymentForm\" action=\"ProcessPayment.aspx\" method=\"post\">");
        Response.Write("<input type=\"hidden\" name=\"applycredit\" id=\"applycredit\" value=\"" + sApplyCredit + "\" />");
        Response.Write("<input type=\"hidden\" name=\"paymentname\" id=\"paymentname\" value=\"Recreation Purchase\" />");
        Response.Write("<input type=\"hidden\" name=\"paymenttype\" id=\"paymenttype\" value=\"\" />");
        Response.Write("<input type=\"hidden\" name=\"paymentlocation\" id=\"paymentlocation\" value=\"website\" />");
        Response.Write("<input type=\"hidden\" name=\"orderstring\" id=\"orderstring\" value=\"" + sOrderString + "\" />");
        Response.Write("<input type=\"hidden\" name=\"serialnumber\" id=\"serialnumber\" value=\"" + sSerialNumber + "\" />");
        Response.Write("<input type=\"hidden\" name=\"ordernumber\" id=\"ordernumber\" value=\"" + sOrderNumber + "\" />");
        Response.Write("<input type=\"hidden\" name=\"itemnumber\" id=\"itemnumber\" value=\"" + sSessionID + "\" />");
        Response.Write("<input type=\"hidden\" name=\"details\" id=\"details\" value=\"" + sDetails + "\" />");
        Response.Write("<input type=\"hidden\" name=\"userid\" id=\"userid\" value=\"" + sUserID.ToString() + "\" />");
        Response.Write("<input type=\"hidden\" name=\"purchaserID\" id=\"purchaserID\" value=\"" + sPurchaserID.ToString() + "\" />");
        Response.Write("<input type=\"hidden\" name=\"transactionamount\" id=\"transactionamount\" value=\"" + sDisplayItemTotal + "\" />");
        Response.Write("<input type=\"hidden\" name=\"sjname\" id=\"sjname\" value=\"" + sName + "\" />");

        while (i < sTotalItems)
        {
            sRosterGrade                   = Request["iRosterGrade" + i.ToString()];
            sRosterShirtSize               = Request["iRosterShirtSize" + i.ToString()];
            sRosterPantsSize               = Request["iRosterPantsSize" + i.ToString()];
            sRosterCoachType               = Request["iRosterCoachType" + i.ToString()];
            sRosterVolunteerCoachName      = Request["iRosterVolunteerCoachName" + i.ToString()];
            sRosterVolunteerCoachDayPhone  = Request["iRosterVolunteerCoachDayPhone" + i.ToString()];
            sRosterVolunteerCoachCellPhone = Request["iRosterVolunteerCoachCellPhone" + i.ToString()];
            sRosterVolunteerCoachEmail     = Request["iRosterVolunteerCoachEmail" + i.ToString()];

            Response.Write("<input type=\"hidden\" name=\"iRosterGrade" + i.ToString() + "\" id=\"iRosterGrade" + i.ToString() + "\" value=\"" + sRosterGrade + "\" />");
            Response.Write("<input type=\"hidden\" name=\"iRosterShirtSize" + i.ToString() + "\" id=\"iRosterShirtSize" + i.ToString() + "\" value=\"" + sRosterShirtSize + "\" />");
            Response.Write("<input type=\"hidden\" name=\"iRosterPantsSize" + i.ToString() + "\" id=\"iRosterPantsSize" + i.ToString() + "\" value=\"" + sRosterPantsSize + "\" />");
            Response.Write("<input type=\"hidden\" name=\"iRosterCoachType" + i.ToString() + "\" id=\"iRosterCoachType" + i.ToString() + "\" value=\"" + sRosterCoachType + "\" />");
            Response.Write("<input type=\"hidden\" name=\"iRosterVolunteerCoachName" + i.ToString() + "\" id=\"iRosterVolunteerCoachName" + i.ToString() + "\" value=\"" + sRosterVolunteerCoachName + "\" />");
            Response.Write("<input type=\"hidden\" name=\"iRosterVolunteerCoachDayPhone" + i.ToString() + "\" id=\"iRosterVolunteerCoachDayPhone" + i.ToString() + "\" value=\"" + sRosterVolunteerCoachDayPhone + "\" />");
            Response.Write("<input type=\"hidden\" name=\"iRosterVolunteerCoachCellPhone" + i.ToString() + "\" id=\"iRosterVolunteerCoachCellPhone" + i.ToString() + "\" value=\"" + sRosterVolunteerCoachCellPhone + "\" />");
            Response.Write("<input type=\"hidden\" name=\"iRosterVolunteerCoachEmail" + i.ToString() + "\" id=\"iRosterVolunteerCoachEmail" + i.ToString() + "\" value=\"" + sRosterVolunteerCoachEmail + "\" />");

            i = i + 1;
        }

        Response.Write("<input type=\"hidden\" name=\"totalrosteritems\" id=\"totalrosteritems\" value=\"" + sTotalItems.ToString() + "\" size=\"3\" maxlength=\"10\" />");

        Response.Write("<div id=\"paymentFormDiv\">");

        //BEGIN: Items to be Purchased --------------------------------------------
        sDisplayItemsToBePurchased = classes.showCartItems(sSessionID);

        Response.Write("<fieldset class=\"fieldset_paymentform\">");
        Response.Write(  "<legend>Items to be Purchased</legend>");
        Response.Write(sDisplayItemsToBePurchased);
        Response.Write("</fieldset>");
        //END: Items to be Purchased ----------------------------------------------

        //BEGIN: Charges ----------------------------------------------------------
        Response.Write("<fieldset class=\"fieldset_paymentform\">");
        Response.Write(  "<legend>Charges</legend>");
        Response.Write(  "<table border=\"0\" id=\"chargeAmounts\">");
        Response.Write(    "<tr>");
        Response.Write(        "<td class=\"labelcol\">Purchase Amount:</td>");
        Response.Write(        "<td>" + sDisplayItemTotal + "</td>");
        Response.Write(    "</tr>");

        //If they are applying credit amount then show that here and the total item charges
        if (sApplyTheCredit)
        {
            Response.Write(    "<tr>");
            Response.Write(        "<td class=\"labelcol\">Account Credit:</td>");
            Response.Write(        "<td>-&nbsp;" + string.Format("{0:#,0.00}", sAccountCredit) + "</td>");
            Response.Write(    "</tr>");
        }

        //Display the PNP Fee if there is one.
        if (sHasPaymentFee)
        {
            Response.Write(    "<tr>");
            Response.Write(        "<td class=\"labelcol\">Processing Fee:</td>");
            Response.Write(        "<td>+&nbsp;" + string.Format("{0:#,0.00}", sFeeAmount) + "</td>");
            Response.Write(    "</tr>");
        }

        if (sHasPaymentFee || sApplyTheCredit)
        {
            Response.Write(    "<tr>");
            Response.Write(        "<td class=\"labelcol totalchargesdisplay\">Total Charges:</td>");
            Response.Write(        "<td class=\"totalchargesdisplay\">" + string.Format("{0:#,0.00}", sTotalAmount) + "</td>");
            Response.Write(    "</tr>");
        }

        Response.Write(  "</table>");
        Response.Write("</fieldset>");
        //END: Charges ------------------------------------------------------------

        //BEGIN: Personal Information ---------------------------------------------
        sStateOrg            = common.getOrgInfo(Convert.ToString(iOrgID), "defaultstate");

        sDisplayStateOptions  = "<select name=\"state\" id=\"state\">";
        sDisplayStateOptions += common.buildStateDropDownOptions(sStateOrg, sState);
        sDisplayStateOptions += "</select>";

        Response.Write("<fieldset class=\"fieldset_paymentform\">");
        Response.Write("  <legend>Billing Information</legend>");
        Response.Write("  <div id=\"billingInfoMsg\">");
        Response.Write("    Please enter your billing information as it appears on your credit card statement, ");
        Response.Write("    then click on the <strong>Process Payment</strong> Button.");
        Response.Write("  </div>");
        Response.Write("  <table border=\"0\">");
        Response.Write("    <tr>");
        Response.Write("        <td class=\"personInfoLabel\">First Name:</td>");
        Response.Write("        <td><input type=\"text\" name=\"firstname\" id=\"firstname\" size=\"30\" maxlength=\"30\" value=\"" + sFirstName + "\" /></td>");
        Response.Write("    </tr>");
        Response.Write("    <tr>");
        Response.Write("        <td class=\"personInfoLabel\">Last Name:</td>");
        Response.Write("        <td><input type=\"text\" name=\"lastname\" id=\"lastname\" size=\"30\" maxlength=\"30\" value=\"" + sLastName + "\" /></td>");
        Response.Write("    </tr>");
        Response.Write("    <tr>");
        Response.Write("        <td class=\"personInfoLabel\">E-mail:</td>");
        Response.Write("        <td><input type=\"text\" name=\"email\" id=\"email\" size=\"50\" maxlength=\"50\" value=\"" + sEmail + "\" /></td>");
        Response.Write("    </tr>");
        Response.Write("    <tr>");
        Response.Write("        <td class=\"personInfoLabel\">Address:</td>");
        Response.Write("        <td><input type=\"text\" name=\"streetaddress\" id=\"streetaddress\" size=\"50\" maxlength=\"50\" value=\"" + sAddress + "\" /></td>");
        Response.Write("    </tr>");
        Response.Write("    <tr>");
        Response.Write("        <td class=\"personInfoLabel\">City:</td>");
        Response.Write("        <td><input type=\"text\" name=\"city\" id=\"city\" size=\"20\" maxlength=\"20\" value=\"" + sCity + "\" /></td>");
        Response.Write("    </tr>");
        Response.Write("    <tr>");
        Response.Write("        <td class=\"personInfoLabel\">State:</td>");
        Response.Write("        <td>" + sDisplayStateOptions + "</td>");
        Response.Write("    </tr>");
        Response.Write("    <tr>");
        Response.Write("        <td class=\"personInfoLabel\">Zip:</td>");
        Response.Write("        <td><input type=\"text\" name=\"zipcode\" id=\"zipcode\" size=\"15\" maxlength=\"15\" value=\"" + sZip + "\" /></td>");
        Response.Write("    </tr>");
        Response.Write(  "</table>");
        Response.Write("</fieldset>");
        //END: Personal Information -----------------------------------------------

        //BEGIN: Credit Card Information ------------------------------------------
        sDisplayCardTypesOptions  = "<select name=\"cardtype\" id=\"cardtype\">";
        sDisplayCardTypesOptions +=    common.buildCreditCardOptions(iOrgID);
        sDisplayCardTypesOptions += "</select>";

        sDisplayOptionsMonth  = "<select name=\"month\" id=\"month\">";
        sDisplayOptionsMonth +=    common.buildMonthOptions();
        sDisplayOptionsMonth += "</select>";

        sDisplayOptionsYear  = "<select name=\"year\" id=\"year\">";
        sDisplayOptionsYear +=    common.buildYearOptions();
        sDisplayOptionsYear += "</select>";

        Response.Write("<fieldset class=\"fieldset_paymentform\">");
        Response.Write("<legend>Credit Card Information</legend>");
        Response.Write("  <table border=\"0\">");
        Response.Write("    <tr>");
        Response.Write("        <td class=\"personInfoLabel\">Credit Card Type:</td>");
        Response.Write("        <td>" + sDisplayCardTypesOptions + "</td>");
        Response.Write("    </tr>");
        Response.Write("    <tr>");
        Response.Write("        <td class=\"personInfoLabel\">Credit Card Number:</td>");
        Response.Write("        <td>");
        Response.Write("            <input type=\"text\" name=\"accountnumber\" id=\"accountnumber\" size=\"30\" maxlength=\"22\" value=\"" + sCreditCardNum + "\" />");
        Response.Write("            <div id=\"creditCardNumberMsg\">Please enter without dashes or spaces (i.e. 0000111188887777)</div>");
        Response.Write("        </td>");
        Response.Write("    </tr>");
        Response.Write("    <tr>");
        Response.Write("        <td class=\"personInfoLabel\">Expiration Month/Year:</td>");
        Response.Write("        <td>" + sDisplayOptionsMonth + "&nbsp;/&nbsp;" + sDisplayOptionsYear + "</td>");
        Response.Write("    </tr>");
        Response.Write("    <tr>");
        Response.Write("        <td class=\"personInfoLabel\">CVV Code:</td>");
        Response.Write("        <td><input type=\"text\" name=\"cvv2\" id=\"cvv2\" size=\"4\" maxlength=\"4\" value=\"\" /></td>");
        Response.Write("    </tr>");
        Response.Write("</table>");
        Response.Write("</fieldset>");
        //END: Credit Card Information --------------------------------------------

        //BEGIN: Process Payment Button -------------------------------------------
        Response.Write("<fieldset class=\"fieldset_paymentform\">");
        Response.Write(  "<legend>Process Payment</legend>");
        Response.Write(  "<div id=\"processPaymentButtonMsg\">");
	Response.Write( "<div class=\"g-recaptcha\" data-sitekey=\"6LcVxxwUAAAAAEYHUr3XZt3fghgcbZOXS6PZflD-\"></div>");
        Response.Write(    "To prevent double billing, please press the \"Process Payment\" button only once ");
        Response.Write(    "and wait for the authorization page to be displayed.  Be patient, it may take up to ");
        Response.Write(    "2 minutes to process your transaction.");
        Response.Write(  "</div>");
        Response.Write(  "<div id=\"processPaymentButton\">");
        Response.Write(  "   <input type=\"button\" name=\"COMPLETE_PAYMENT\" id=\"COMPLETE_PAYMENT\" value=\"PROCESS PAYMENT\" onclick=\"eGovLink.Class.processPayment();\" />");
        Response.Write(  "</div>");
        Response.Write("</fieldset>");
        //END: Process Payment Button ---------------------------------------------

        //BEGIN: IP Address Message -----------------------------------------------
        Response.Write("<div id=\"paymentFormIPMsg\">");
        Response.Write("NOTE: Your IP Address [" + HttpContext.Current.Request.UserHostAddress + "] has been logged with this transaction.");
        Response.Write("</div>");
        //END: IP Address Message -------------------------------------------------

        Response.Write("</div>");
        Response.Write("</form>");
    }

}

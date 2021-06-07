using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_classes_class_cart : System.Web.UI.Page
{
    double startCounter = 0.00;

    protected void Page_PreInit(object sender, EventArgs e)
    {
        // This is the earliest thing the page does, so set the start time here.
        startCounter = DateTime.Now.TimeOfDay.TotalSeconds;
    }

    protected void Page_PreRender(object sender, EventArgs e)
    {
        common.logThePageVisit(startCounter, "class_cart.aspx", "public");
    }

    protected void Page_Load(object sender, EventArgs e)
    {
    }

    public void displayCart(Int32 iOrgID, string iSessionID)
    {
        Boolean sIsClassFilled = false;
        Boolean sCartHasMerchandise                     = common.cartHasMerchandise(iSessionID);
        Boolean sFeatureIsTurnedOnForPublic_merchandise = common.featureIsTurnedOnForPublic(iOrgID, "merchandise");
        Boolean sOrgHasFeature_merchandise              = common.orgHasFeature(Convert.ToString(iOrgID), "merchandise");
        Boolean sOrgHasFeature_publicAccountPayments    = common.orgHasFeature(Convert.ToString(iOrgID), "public account payments");
        Boolean sOrgHasDisplay_refundPolicy             = common.orgHasDisplay(Convert.ToString(iOrgID), "refund policy");

        double sTotal                 = 0.00;
        double sTotalDue              = 0.00;
        double sAccountCredit         = 0.00;
        double sAccountCreditOriginal = 0.00;

        Int32 sRowCount        = 0;
        Int32 sPurchaserID     = 0;
        Int32 sUserID          = 0;
        Int32 sClassID         = 0;
        Int32 sPriceDiscountID = 0;
        //Int32 sFamilyChildAge  = 0;
        Int32 sAvail           = 0;
        Int32 sClassTimeID     = 0;
        Int32 sAge             = 0;

        string sOrgVirtualSiteName           = common.getOrgInfo(Convert.ToString(iOrgID), "orgVirtualSiteName");
        string sSQL                          = "";
        string sFormActionURL                = "";
        string sItemType                     = "";
        string sClassName                    = "";
        string sRowClassName                 = "";
        string sRowClassNameDisplay          = "";
        string sRowActivityNo                = "";
        string sDiscount                     = "";
        string sFamilyMemberName             = "";
        string sFamilyMemberBirthDate        = "";
        string sQuantity                     = "";
        string sBuyOrWait                    = "";
        string sPurchaseWaitListDisplay      = "";
        string sAgeCompareDate               = "";
        string sDisplayTotal                 = "";
        string sAccountCreditDisplay         = "";
        string sAccountCreditOriginalDisplay = "";
        string sTotalDueDisplay              = "";
        string sRefundPolicy                 = "";
		string sOrgName                      = common.getOrgName(Convert.ToString(iOrgID));

        if (Request.QueryString["userid"] == "" || Request.QueryString["userid"] == null)
        {
            sUserID = common.getCartUserID(iSessionID);
        }
        else
        {
            try
            {
                sUserID = Convert.ToInt32(Request.QueryString["userid"]);
            }
            catch
            {
                Response.Redirect("class_list.aspx");
            }
        }

        if (Request.QueryString["iClassID"] != "")
        {
            try
            {
                sClassID   = Convert.ToInt32(Request.QueryString["iClassID"]);
                sClassName = classes.getClassName(sClassID);
                sClassName = common.decodeUTFString(sClassName);
            }
            catch
            {
                sClassID = 0;
            }
        }

        sIsClassFilled = classes.classIsFilled(sClassID);

        Response.Write("<div id=\"cartPageHeader\">");

        //if(sClassName != "" && ! iIsRegattaTeam) {
        if (sClassName != "")
        {
            if (!sIsClassFilled)
            {
                Response.Write("  <input type=\"button\" name=\"classSignUpButton\" id=\"classSignUpButton\" value=\"Purchase Another '" + sClassName + "'\" />");
            }
        }

        Response.Write("  <input type=\"button\" name=\"classListButton\" id=\"classListButton\" value=\"Purchase Another Class/Event\" />");

        if (sOrgHasFeature_merchandise && sFeatureIsTurnedOnForPublic_merchandise && !sCartHasMerchandise)
        {
            Response.Write("  <input type=\"button\" name=\"purchaseMerchandiseButton\" id=\"purchaseMerchandiseButton\" value=\"Purchase Merchandise Here\" onclick=\"location.href='../merchandise/merchandiseofferrings.asp';\" />");
        }

        Response.Write("</div>");

        sSQL  = "SELECT ";
        sSQL += " cartid, ";
        sSQL += " classid, ";
        sSQL += " isnull(userid,0) as userid, ";
        sSQL += " isnull(familymemberid,0) as familymemberid, ";
        sSQL += " quantity, ";
        sSQL += " amount, ";
        sSQL += " optionid, ";
        sSQL += " classtimeid, ";
        sSQL += " pricetypeid, ";
        sSQL += " buyorwait, ";
        sSQL += " sessionid, ";
        sSQL += " orgid, ";
        sSQL += " rostergrade, ";
        sSQL += " rostershirtsize, ";
        sSQL += " rosterpantssize, ";
        sSQL += " rostercoachtype, ";
        sSQL += " rostervolunteercoachname, ";
        sSQL += " rostervolunteercoachdayphone, ";
        sSQL += " rostervolunteercoachcellphone, ";
        sSQL += " rostervolunteercoachemail, ";
        sSQL += " isregatta, ";
        sSQL += " regattateamid, ";
        sSQL += " I.itemtype, ";
        sSQL += " I.isshippingfee, ";
        sSQL += " I.issalestax ";
        sSQL += " FROM egov_class_cart C, ";
        sSQL +=      " egov_item_types I ";
        sSQL += " WHERE C.itemtypeid = I.itemtypeid ";
        sSQL += " AND sessionid_csharp = '" + iSessionID + "' ";
        sSQL += " ORDER BY I.cartdisplayorder, C.cartid";
	//Response.Write(sSQL);
	

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                sRowCount       = sRowCount + 1;
                sItemType       = Convert.ToString(myReader["itemtype"]);
                sQuantity       = Convert.ToString(myReader["quantity"]);
                sBuyOrWait      = Convert.ToString(myReader["buyorwait"]);
                sAgeCompareDate = Convert.ToString(DateTime.Now);
                sClassTimeID    = Convert.ToInt32(myReader["classtimeid"]);

                sFamilyMemberName = "";
                sFamilyMemberBirthDate = "";
                sPurchaseWaitListDisplay = "Wait List";

                sAvail                 = classes.getCurrentAvailability(sClassTimeID);
                sPriceDiscountID       = classes.getClassPriceDiscountID(Convert.ToInt32(myReader["classid"]));
                sDiscount              = common.decodeUTFString(classes.getDiscountPhrase(sPriceDiscountID));
                sRowActivityNo         = common.decodeUTFString(classes.getActivityNo(sClassTimeID));
                sFamilyMemberName      = classes.getFamilyMemberName(Convert.ToInt32(myReader["familymemberid"]));
                sFamilyMemberBirthDate = classes.getBirthDate(Convert.ToInt32(myReader["familymemberid"]));

                sAge = Convert.ToInt32(classes.getAgeOnDate(Convert.ToDateTime(sFamilyMemberBirthDate),
                                                            Convert.ToDateTime(sAgeCompareDate)));

                //sRowClassName = "<strong>";
                //sRowClassName += classes.getClassName(Convert.ToInt32(myReader["classid"]));
                //sRowClassName += " (" + sRowActivityNo + ")";
                //sRowClassName += "</strong>";

                sRowClassName  = classes.getClassName(Convert.ToInt32(myReader["classid"]));
                sRowClassName  = common.decodeUTFString(sRowClassName);
                sRowClassName += " (" + sRowActivityNo + ")";

                sRowClassNameDisplay  = "<strong>";
                sRowClassNameDisplay += sRowClassName;
                sRowClassNameDisplay += "</strong>";

                if (sDiscount != "")
                {
                    sRowClassNameDisplay += "<br /><span class=\"discounttext\">(" + sDiscount + ")</span>";
                }

                if (sBuyOrWait == "B")
                {
                    sPurchaseWaitListDisplay = string.Format("{0:0.00}", Convert.ToInt32(myReader["amount"]));

                    sTotal = sTotal + Convert.ToDouble(myReader["amount"]);
                }

                if (sRowCount == 1)
                {
                    sPurchaserID = Convert.ToInt32(myReader["userid"]);

                    sFormActionURL = ConfigurationManager.AppSettings["paymenturl"];
                    sFormActionURL += "/";
                    sFormActionURL += sOrgVirtualSiteName;
                    sFormActionURL += "/rd_classes/class_paymentform.aspx";

                    Response.Write("<form name=\"cartForm\" id=\"cartForm\" method=\"post\" action=\"" + sFormActionURL + "\">");
                    Response.Write("  <input type=\"hidden\" name=\"purchaserID\" id=\"purchaserID\" value=\"" + sPurchaserID.ToString() + "\" />");
                    Response.Write("  <input type=\"hidden\" name=\"userid\" id=\"userid\" value=\"" + sUserID.ToString() + "\" />");
                    Response.Write("  <input type=\"hidden\" name=\"iClassID\" id=\"iClassID\" value=\"" + sClassID.ToString() + "\" />");
                    Response.Write("  <input type=\"hidden\" name=\"sessionID\" id=\"sessionID\" value=\"" + iSessionID + "\" />");
                    Response.Write("  <input type=\"hidden\" name=\"init\" id=\"init\" value=\"N\" size=\"1\" maxlength=\"1\" />");

                    Response.Write("<div id=\"cartDetails\">");
                    Response.Write("  <fieldset class=\"fieldset\">");
                    Response.Write("    <legend>" + sOrgName + " Cart</legend>");

                    Response.Write("  <div class=\"cartTableContainer\">");
                    Response.Write("    <table id=\"cartTable\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\">");
                    Response.Write("      <thead>");
                    Response.Write("      <tr>");
                    Response.Write("          <td colspan=\"2\" class=\"cartTableColumn1\">Class/Event/Program</td>");
                    Response.Write("          <td>Participant</td>");
                    Response.Write("          <td>Age</td>");
                    Response.Write("          <td>Qty</td>");
                    Response.Write("          <td class=\"cartTablePrice\">Price</td>");
                    //Response.Write("          <td class=\"cartTableColumnLast\">&nbsp;</td>");
                    Response.Write("      </tr>");
                    Response.Write("      </thead>");
                    Response.Write("      <tbody>");
                }

                Response.Write("      <tr>");

                switch (sItemType)
                {
                    case "recreation activity":
                        Response.Write("          <td><button name=\"removeButton" + sRowCount + "\" id=\"removeButton" + sRowCount + "\" class=\"cartTableRemoveButton\" onclick=\"eGovLink.Class.removeItem(" + Convert.ToString(myReader["cartid"]) + ", " + Convert.ToString(myReader["classtimeid"]) + ", '" + Convert.ToString(myReader["buyorwait"]) + "');return false;\">Remove</button>&nbsp;</td>");
                        //Response.Write(              "<input type=\"button\" name=\"removeButton\" id=\"removeButton\" value=\"Remove\" onclick=\"eGovLink.Class.RemoveItem(" + Convert.ToString(myReader["cartid"]) + ", " + Convert.ToString(myReader["classtimeid"]) + ", '" + Convert.ToString(myReader["buyorwait"]) + "');\" />");
                        Response.Write("          <td class=\"cartTableClassName\">");
                        Response.Write("              <input type=\"hidden\" name=\"cartid" + sRowCount + "\" id=\"cartid" + sRowCount + "\" value=\"" + Convert.ToString(myReader["cartid"]) + "\" />");
                        Response.Write("              <input type=\"hidden\" name=\"classname" + sRowCount + "\" id=\"classname" + sRowCount + "\" value=\"" + sRowClassName + "\" />");
                        Response.Write("              <input type=\"hidden\" name=\"iRosterGrade\"" + sRowCount + "\" id=\"iRosterGrade\"" + sRowCount + "\" value=\"" + Convert.ToString(myReader["rostergrade"]) + "\" />");
                        Response.Write("              <input type=\"hidden\" name=\"iRosterShirtSize\"" + sRowCount + "\" id=\"iRosterShirtSize\"" + sRowCount + "\" value=\"" + Convert.ToString(myReader["rostershirtsize"]) + "\" />");
                        Response.Write("              <input type=\"hidden\" name=\"iRosterPantsSize\"" + sRowCount + "\" id=\"iRosterPantsSize\"" + sRowCount + "\" value=\"" + Convert.ToString(myReader["rosterpantssize"]) + "\" />");
                        Response.Write("              <input type=\"hidden\" name=\"iRosterCoachType\"" + sRowCount + "\" id=\"iRosterCoachType\"" + sRowCount + "\" value=\"" + Convert.ToString(myReader["rostercoachtype"]) + "\" />");
                        Response.Write("              <input type=\"hidden\" name=\"iRosterVolunteerCoachName\"" + sRowCount + "\" id=\"iRosterVolunteerCoachName\"" + sRowCount + "\" value=\"" + Convert.ToString(myReader["rostervolunteercoachname"]) + "\" />");
                        Response.Write("              <input type=\"hidden\" name=\"iRosterVolunteerCoachDayPhone\"" + sRowCount + "\" id=\"iRosterVolunteerCoachDayPhone\"" + sRowCount + "\" value=\"" + Convert.ToString(myReader["rostervolunteercoachdayphone"]) + "\" />");
                        Response.Write("              <input type=\"hidden\" name=\"iRosterVolunteerCoachCellPhone\"" + sRowCount + "\" id=\"iRosterVolunteerCoachCellPhone\"" + sRowCount + "\" value=\"" + Convert.ToString(myReader["rostervolunteercoachcellphone"]) + "\" />");
                        Response.Write("              <input type=\"hidden\" name=\"iRosterVolunteerCoachEmail\"" + sRowCount + "\" id=\"iRosterVolunteerCoachEmail\"" + sRowCount + "\" value=\"" + Convert.ToString(myReader["rostervolunteercoachemail"]) + "\" />");
                        Response.Write("              <input type=\"hidden\" name=\"quantity" + sRowCount + "\" id=\"quantity" + sRowCount + "\" value=\"" + sQuantity + "\" />");
                        Response.Write("              <input type=\"hidden\" name=\"original" + sRowCount + "\" id=\"original" + sRowCount + "\" value=\"" + sQuantity + "\" />");
                        Response.Write("              <input type=\"hidden\" name=\"avail" + sRowCount + "\" id=\"avail" + sRowCount + "\" value=\"" + Convert.ToString(sAvail) + "\" />");
                        Response.Write(sRowClassNameDisplay);
                        Response.Write("          </td>");
                        Response.Write("          <td>" + sFamilyMemberName + "</td>");
                        Response.Write("          <td>" + sAge.ToString() + "</td>");
                        Response.Write("          <td>" + sQuantity + "</td>");
                        Response.Write("          <td class=\"cartTablePrice\">" + sPurchaseWaitListDisplay + "</td>");

                        break;
                }

                Response.Write("      </tr>");
            }

            if (sRowCount > 0)
            {
                if (sTotal == null)
                {
                    sTotal = 0;
                }

                sDisplayTotal = string.Format("{0:0.00}", sTotal);

                Response.Write("      </tbody>");
                Response.Write("      <tfoot>");
                Response.Write("      <tr>");
                Response.Write("          <td colspan=\"5\" class=\"cartTableFooterColumn1\">Total Charges:</td>");
                Response.Write("          <td class=\"cartTableFooterColumnTotal\">" + sDisplayTotal + "</td>");
                Response.Write("      </tr>");
                Response.Write("      </tfoot>");
                Response.Write("    </table>");
                Response.Write("      </div>");

                //BEGIN: Account Credit --------------------------------------------------
                //If the org allows payment with citizen account amounts
                if (sOrgHasFeature_publicAccountPayments)
                {
                    sTotalDue = sTotal;

                    sAccountCredit = common.getCitizenAccountAmount(sPurchaserID);
                    sAccountCreditOriginal = sAccountCredit;
                    sAccountCreditOriginalDisplay = string.Format("{0:0.00}", Convert.ToInt32(sAccountCreditOriginal));

                    if (sAccountCredit < 0)
                    {
                        sAccountCredit = 0.00;
                    }
                    else
                    {
                        if (sAccountCredit > Convert.ToDouble(sTotal))
                        {
                            sAccountCredit = Convert.ToDouble(sTotal);
                        }
                    }

                    Response.Write("    <table id=\"cartTableAccountCredit\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\">");
                    Response.Write("      <thead>");
                    Response.Write("      <tr>");
                    Response.Write("          <td class=\"cartTableColumn1\">Account Credit</td>");
                    Response.Write("          <td class=\"cartTableColumnLast\">&nbsp;</td>");
                    Response.Write("      </tr>");
                    Response.Write("      </thead>");

                    //If the citizen has a positive account balance then...
                    if (sAccountCredit > 0.00)
                    {
                        //...display the "Apply Credit" and "Total Due" rows.
                        sAccountCreditDisplay = string.Format("{0:0.00}", Convert.ToInt32(sAccountCredit));

                        sTotalDue = sTotalDue - sAccountCredit;
                        sTotalDueDisplay = string.Format("{0:0.00}", Convert.ToInt32(sTotalDue));

                        Response.Write("      <tbody>");
                        Response.Write("      <tr>");
                        Response.Write("          <td class=\"cartTableColumn1\" id=\"columnApplyCredit\">");
                        Response.Write("<input type=\"checkbox\" id=\"checkapplycredit\" value=\"1\" checked=\"checked=\" onclick=\"eGovLink.Class.toggleApplyCredit()\" />");
                        Response.Write("<input type=\"hidden\" name=\"applycredit\" id=\"applycredit\" value=\"yes\" />");
                        Response.Write("<input type=\"hidden\" name=\"accountcredit\" id=\"accountcredit\" value=\"" + Convert.ToString(sAccountCredit) + "\" />");
                        Response.Write("<input type=\"hidden\" name=\"totaldue\" id=\"totaldue\" value=\"" + Convert.ToString(sTotalDue) + "\" />");
                        Response.Write("Apply Your Account Credit to This Purchase:");
                        Response.Write("          </td>");
                        Response.Write("          <td class=\"cartTableColumnLast\" id=\"columnApplyCreditAmount\">" + sAccountCreditDisplay + "</td>");
                        Response.Write("      </tr>");
                        Response.Write("      <tr>");
                        Response.Write("          <td class=\"cartTableColumn1\" id=\"columnAmountOnAccount\">(Total Amount on Account: " + sAccountCreditOriginalDisplay + ")</td>");
                        Response.Write("          <td class=\"cartTableColumnLast\">&nbsp;</td>");
                        Response.Write("      </tr>");
                        Response.Write("      </tbody>");
                        Response.Write("      <tfoot>");
                        Response.Write("      <tr>");
                        Response.Write("          <td class=\"cartTableFooterColumn1\">Total Due:</td>");
                        Response.Write("          <td class=\"cartTableFooterColumnTotal\"><span id=\"totalDueDisplay\">" + sTotalDueDisplay + "</span></td>");
                        Response.Write("      </tr>");
                        Response.Write("      </tfoot>");

                    }
                    else
                    {
                        //Set the "credit applied" to ($0) in a hidden field AND...
                        //set the "total due" to the "total charges"
                        Response.Write("      <tbody>");
                        Response.Write("      <tr>");
                        Response.Write("          <td class=\"cartTableColumn1\" id=\"columnAmountOnAccount\">(total Amount on Account: " + sAccountCreditOriginalDisplay + ")</td>");
                        Response.Write("          <td class=\"cartTableColumnLast\">");
                        Response.Write("<input type=\"hidden\" name=\"applycredit\" id=\"applycredit\" value=\"no\" />");
                        Response.Write("<input type=\"hidden\" name=\"accountcredit\" id=\"accountcredit\" value=\"0.00\" />");
                        Response.Write("<input type=\"hidden\" name=\"totaldue\" id=\"totaldue\" value=\"" + Convert.ToString(sTotalDue) + "\" />");
                        Response.Write("          </td>");
                        Response.Write("      </tr>");
                        Response.Write("      </tbody>");
                    }

                    Response.Write("    </table>");
                }
                //END: Account Credit ----------------------------------------------------

                //BEGIN: Continue Purchase Button ----------------------------------------
                Response.Write("<div class=\"cartBottomButtonRow\">");
                Response.Write("<input type=\"button\" name=\"completeButton\" id=\"completeButton\" value=\"Continue Purchase\" />");
                Response.Write("</div>");
                //END: Continue Purchase Button ------------------------------------------

                //BEGIN: Refund Policy ---------------------------------------------------
                if (sOrgHasDisplay_refundPolicy)
                {
                    sRefundPolicy = common.getOrgDisplay(Convert.ToString(iOrgID), "refund policy");
                    sRefundPolicy = common.decodeUTFString(sRefundPolicy);

                    Response.Write("<fieldset id=\"cartRefundPolicy\" class=\"fieldset\">");
                    Response.Write("<legend>Our Refund Policy</legend>");
                    Response.Write(sRefundPolicy);
                    Response.Write("</fieldset>");
                }
                //END: Refund Policy -----------------------------------------------------

                Response.Write("  </fieldset>");
                Response.Write("</div>");

                Response.Write("  <input type=\"hidden\" name=\"totalitems\" id=\"totalitems\" value=\"" + Convert.ToString(sRowCount - 1) + "\" />");
                Response.Write("  <input type=\"hidden\" name=\"nTotal\" id=\"nTotal\" value=\"" + Convert.ToString(sTotal) + "\" />");
                Response.Write("</form>");
            }

//            Response.Write("<div style=\"border:1px solid red; text-align:center;\"><span style=\"color:#FF0000; font-size:10pt;\">*** We're sorry but, your session has timed out. You must start your purchase again. ***</span><br /><span style=\"color:#FF0000; font-size:11pt; font-weight:bold;\">*** You have not been charged for this transaction. ***</span></div>");
        }
        else
        {
            //if (sOrgHasFeature_merchandise && sFeatureIsTurnedOnForPublic_merchandise)
            //{
            //    Response.Write("<input type=\"button\" name=\"purchaseMerchandiseButton\" id=\"purchaseMerchandiseButton\" value=\"Purchase Merchandise Here\" onclick=\"location.href='../merchandise/merchandiseofferings.asp';\" />");
            //}

            if (Request.QueryString["message"] == "timeout")
            {
                Response.Write("<div id=\"cartTimeoutMsg\">");
                Response.Write(  "*** We are sorry but your session has timed out.  You must start your purchase again. ***<br />");
                Response.Write  ("<div id=\"cartTimeoutMsgNoTransaction\">*** You have not been charged for this transaction. ***</div>");
                Response.Write("</div>");
            }

            Response.Write("<div id=\"cartNoItems\">There are no items in the cart</div>");
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
    }
}

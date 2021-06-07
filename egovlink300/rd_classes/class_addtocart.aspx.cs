using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class class_addtocart : System.Web.UI.Page
{
    double startCounter = 0.00;

    static string sOrgID = common.getOrgId();
    static string sOrgName = common.getOrgName(sOrgID);

    protected void Page_PreInit(object sender, EventArgs e)
    {
        // This is the earliest thing the page does, so set the start time here.
        startCounter = DateTime.Now.TimeOfDay.TotalSeconds;
    }

    protected void Page_PreRender(object sender, EventArgs e)
    {
        common.logThePageVisit(startCounter, "class_addtocart.aspx", "public");
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        Boolean sIsParent                           = false;
        Boolean sDisplayRosterPublic                = false;
        Boolean sOrgHasFeatureDiscounts             = false;
        Boolean sOrgHasFeatureEmergencyInfoRequired = false;

        double sAmount = 0.00;

        Int32 sOrgID          = 0;
        Int32 sUserID         = 0;
        Int32 sClassID        = 0;
        Int32 sTimeID         = 0;
        Int32 sPriceTypeID    = 0;
        Int32 sOptionID       = 0;
        Int32 sFamilyMemberID = 0;
        Int32 sAttendeeUserID = 0;
        Int32 sQuantity       = 0;
        Int32 sIsDropIn       = 0;  //since we will have to convert this to a 1 or 0 to insert, and it's not changed in the code, we just set it to "0"
        Int32 sCartID         = 0;
        Int32 sClassTypeID    = 0;
        Int32 sCategoryID     = 0;
        Int32 sItemTypeID     = common.getItemTypeID("recreation activity");  //this is what kind of thing they are buying

        //string sCartID = "";

        string sBuyOrWait  = "";
        string sStatus     = "";
        string sDropInDate = "";
        string sRosterGrade = "";
        string sRosterShirtSize = "";
        string sRosterPantsSize = "";
        string sRosterCoachType = "";
        string sRosterVolunteerCoachName = "";
        string sRosterVolunteerCoachDayPhone = "";
        string sRosterVolunteerCoachCellPhone = "";
        string sRosterVolunteerCoachEmail = "";
        string sSkipVolCoachDayAreaCode = "";
        string sSkipVolCoachDayExchange = "";
        string sSkipVolCoachDayLine = "";
        string sSkipVolCoachCellAreaCode = "";
        string sSkipVolCoachCellExchange = "";
        string sSkipVolCoachCellLine = "";
        string sEmergencyContact = "";
        string sEmergencyPhone = "";
        string sRedirectURL = "";

        try
        {
            sOrgID = Convert.ToInt32(Request.Form["orgid"]);
        }
        catch
        {
            sOrgID = 0;
        }
		if (sOrgID == 0)
		{
			try
			{
				sOrgID = Int32.Parse(common.getOrgId());
			}
			catch
			{
				//Nothing
			}
		}
		

        try
        {
            sUserID = Convert.ToInt32(Request.Form["userid"]);
        }
        catch
        {
            sUserID = 0;
        }

        try
        {
            sClassID = Convert.ToInt32(Request.Form["classid"]);
        }
        catch
        {
            sClassID = 0;
        }

        try
        {
            sTimeID = Convert.ToInt32(Request.Form["timeid"]);
        }
        catch
        {
            sTimeID = 0;
        }

        try
        {
            sPriceTypeID = Convert.ToInt32(Request.Form["pricetypeid"]);
        }
        catch
        {
            sPriceTypeID = 0;
        }

        try
        {
            sOptionID = Convert.ToInt32(Request.Form["optionid"]);
        }
        catch
        {
            sOptionID = 0;
        }

        try
        {
            sIsParent = Convert.ToBoolean(Request.Form["isparent"]);
        }
        catch
        {
            sIsParent = false;
        }

        try
        {
            sClassTypeID = Convert.ToInt32(Request.Form["classtypeid"]);
        }
        catch
        {
            sClassTypeID = 0;
        }

        try
        {
            sCategoryID = Convert.ToInt32(Request.Form["categoryid"]);
        }
        catch
        {
            sCategoryID = 0;
        }

        try
        {
            sDisplayRosterPublic = Convert.ToBoolean(Request.Form["displayrosterpublic"]);
        }
        catch
        {
            sDisplayRosterPublic = false;
        }

        if (sOptionID == 2)
        {
            sAttendeeUserID = sUserID;  //Purchaser is the attendee
            sFamilyMemberID = classes.getCitizenFamilyID(sUserID);

            try
            {
                sQuantity = Convert.ToInt32(Request.Form["quantity"]);
            }
            catch
            {
                sQuantity = 0;
            }
        }
        else
        {
            try
            {
                sFamilyMemberID = Convert.ToInt32(Request.Form["familymemberid"]);
            }
            catch
            {
                sFamilyMemberID = 0;
            }

            sAttendeeUserID = classes.getAttendeeUserID(sFamilyMemberID);
            sQuantity       = 1;
        }

        sBuyOrWait = Request.Form["buyorwait"];

        if (sBuyOrWait == "W")
        {
            sAmount      = 0.00;  //The waitlist is free
            sStatus      = "WAITLIST";
            sPriceTypeID = 0;
        }
        else
        {
            sAmount = classes.getAmount(sPriceTypeID, sClassID);
            sStatus = "ACTIVE";
        }

        if (sDisplayRosterPublic)
        {
            if (Request.Form["rostergrade"] != null)
            {
                sRosterGrade = Request.Form["rostergrade"];
                sRosterGrade = common.dbready_string(sRosterGrade, 2);
            }

            if (Request.Form["rostershirtsize"] != null)
            {
                sRosterShirtSize = Request.Form["rostershirtsize"];
                sRosterShirtSize = common.dbready_string(sRosterShirtSize, 50);
            }

            if (Request.Form["rosterpantssize"] != null)
            {
                sRosterPantsSize = Request.Form["rosterpantssize"];
                sRosterPantsSize = common.dbready_string(sRosterPantsSize, 50);
            }

            if (Request.Form["rostercoachtype"] != null)
            {
                sRosterCoachType = Request.Form["rostercoachtype"];
                sRosterCoachType = common.dbready_string(sRosterCoachType, 50);
            }

            if (Request.Form["rostervolunteercoachname"] != null)
            {
                sRosterVolunteerCoachName = Request.Form["rostervolunteercoachname"];
                sRosterVolunteerCoachName = common.dbready_string(sRosterVolunteerCoachName, 100);
            }

            //BEGIN: Build the Volunteer Coach Day Phone --------------------------------
            if (Request.Form["skip_volcoachday_areacode"] != null)
            {
                sSkipVolCoachDayAreaCode = common.dbready_string(Request.Form["skip_volcoachday_areacode"], 3);
            }

            if (Request.Form["skip_volcoachday_exchange"] != null)
            {
                sSkipVolCoachDayExchange = common.dbready_string(Request.Form["skip_volcoachday_exchange"], 3);
            }

            if (Request.Form["skip_volcoachday_line"] != null)
            {
                sSkipVolCoachDayLine = common.dbready_string(Request.Form["skip_volcoachday_line"], 4);
            }

            sRosterVolunteerCoachDayPhone = sSkipVolCoachDayAreaCode + sSkipVolCoachDayExchange + sSkipVolCoachDayLine;
            sRosterVolunteerCoachDayPhone = sRosterVolunteerCoachDayPhone.Trim();

            if (sRosterVolunteerCoachDayPhone != null)
            {
                sRosterVolunteerCoachDayPhone = common.dbready_string(sRosterVolunteerCoachDayPhone, 10);
            }
            else
            {
                sRosterVolunteerCoachDayPhone = "NULL";
            }
            //END: Build the Volunteer Coach Day Phone ----------------------------------

            //BEGIN: Build the Volunteer Coach Cell Phone -------------------------------
            if (Request.Form["skip_volcoachcell_areacode"] != null)
            {
                sSkipVolCoachCellAreaCode = common.dbready_string(Request.Form["skip_volcoachcell_areacode"], 3);
            }

            if (Request.Form["skip_volcoachcell_exchange"] != null)
            {
                sSkipVolCoachCellExchange = common.dbready_string(Request.Form["skip_volcoachcell_exchange"], 3);
            }

            if (Request.Form["skip_volcoachcell_line"] != null)
            {
                sSkipVolCoachCellLine = common.dbready_string(Request.Form["skip_volcoachcell_line"], 4);
            }

            sRosterVolunteerCoachCellPhone = sSkipVolCoachCellAreaCode + sSkipVolCoachCellExchange + sSkipVolCoachCellLine;
            sRosterVolunteerCoachCellPhone = sRosterVolunteerCoachCellPhone.Trim();

            if (sRosterVolunteerCoachCellPhone != null)
            {
                sRosterVolunteerCoachCellPhone = common.dbready_string(sRosterVolunteerCoachCellPhone, 10);
            }
            else
            {
                sRosterVolunteerCoachCellPhone = "NULL";
            }
            //END: Build the Volunteer Coach Cell Phone ---------------------------------

            if (Request.Form["rostervolunteercoachemail"] != null)
            {
                sRosterVolunteerCoachEmail = Request.Form["rostervolunteercoachemail"];
                sRosterVolunteerCoachEmail = common.dbready_string(sRosterVolunteerCoachEmail, 100);
            }
        }

        sRosterGrade                   = "'" + sRosterGrade                          + "'";
        sRosterShirtSize               = "'" + sRosterShirtSize                      + "'";
        sRosterPantsSize               = "'" + sRosterPantsSize                      + "'";
        sRosterCoachType               = "'" + sRosterCoachType                      + "'";
        sRosterVolunteerCoachName      = "'" + sRosterVolunteerCoachName             + "'";
        sRosterVolunteerCoachDayPhone  = "'" + sRosterVolunteerCoachDayPhone.Trim()  + "'";
        sRosterVolunteerCoachCellPhone = "'" + sRosterVolunteerCoachCellPhone.Trim() + "'";
        sRosterVolunteerCoachEmail     = "'" + sRosterVolunteerCoachEmail            + "'";

        sOrgHasFeatureDiscounts             = common.orgHasFeature(Convert.ToString(sOrgID), "discounts");
        sOrgHasFeatureEmergencyInfoRequired = common.orgHasFeature(Convert.ToString(sOrgID), "emergency info required");

        sCartID = addToCart(sOrgID,
                            sClassID,
                            sUserID,
                            sTimeID,
                            sFamilyMemberID,
                            sQuantity,
                            sAmount,
                            sPriceTypeID,
                            sOptionID,
                            sBuyOrWait,
                            sIsParent,
                            sClassTypeID,
                            sItemTypeID,
                            sIsDropIn,
                            sDropInDate,
                            sDisplayRosterPublic,
                            sRosterGrade,
                            sRosterShirtSize,
                            sRosterPantsSize,
                            sRosterCoachType,
                            sRosterVolunteerCoachName,
                            sRosterVolunteerCoachDayPhone,
                            sRosterVolunteerCoachCellPhone,
                            sRosterVolunteerCoachEmail);

        //Waitlist place holder - They are not changed in the reset prices call below.
        if (sBuyOrWait == "W")
        {
            classes.addToCartPrice(sCartID,
                                   0,        //pricetypeid
                                   0.00,     //amount
                                   sQuantity);
        }

        //Increment egov_class_time counts
        classes.updateClassTime(sTimeID,
                                sQuantity,
                                sBuyOrWait);

        //If this is a series parent, update the children
        if (sIsParent && sClassTypeID == 1)
        {
            classes.updateClassTimeSeriesChildren(sClassID,
                                                  sQuantity,
                                                  sBuyOrWait);
        }

        //Recalculate the prices - This puts prices into the egov_class_cart_price
        //table for the Buys
        classes.resetCartPrices();

        //Recalculate any discounts
        if (sOrgHasFeatureDiscounts)
        {
            classes.determineDiscounts();
        }

        //Update the Emergency Info....if needed (required)
        if (sOrgHasFeatureEmergencyInfoRequired)
        {
            if (Request.Form["emergencycontact"] != "")
            {
                sEmergencyContact = Request.Form["emergencycontact"];
            }

            if (Request.Form["emergencyphone"] != "")
            {
                sEmergencyPhone = Request.Form["emergencyphone"];
            }

            classes.updateEmergencyInfo(sOrgID,
                                        sFamilyMemberID,
                                        sEmergencyContact,
                                        sEmergencyPhone);
        }

        //BEGIN: Build redirect url -------------------------------------------------------
        sRedirectURL  = "?userid="     + sUserID.ToString();
        sRedirectURL += "&iClassID="   + sClassID.ToString();
        sRedirectURL += "&categoryID=" + sCategoryID.ToString();
        

//We no longer need to redirect to "confirminfo" as it has been built into class_signup.aspx
//        if (sOrgHasFeatureEmergencyInfoRequired)
//        {
            //Take them to the page that shows the emergency contact info for the enrollee.
//            sRedirectURL  = "confirminfo.aspx" + sRedirectURL;
//            sRedirectURL += "&attendeeUserID=" + sAttendeeUserID.ToString();
//        }
//        else
//        {
            //Redirect to the cart page.
            sRedirectURL = "class_cart.aspx" + sRedirectURL;
//        }

        Response.Redirect(sRedirectURL);
        //END: Build redirect url ---------------------------------------------------------

        /*
        Response.Write("orgid: [" + sOrgID.ToString() + "]<br />");
        Response.Write("userid: [" + sUserID.ToString() + "]<br />");
        Response.Write("timeid: [" + sTimeID.ToString() + "]<br />");
        Response.Write("pricetypeid: [" + sPriceTypeID.ToString() + "]<br />");
        Response.Write("optionid: [" + sOptionID.ToString() + "]<br />");
        Response.Write("attendeeuserid: [" + sAttendeeUserID.ToString() + "]<br />");
        Response.Write("familymemberid: [" + sFamilyMemberID.ToString() + "]<br />");
        Response.Write("quantity: [" + sQuantity.ToString() + "]<br />");
        Response.Write("buyorwait: [" + sBuyOrWait.ToString() + "]<br />");
        Response.Write("amount: [" + sAmount.ToString() + "]<br />");
        Response.Write("status: [" + sStatus + "]<br />");
        Response.Write("displayrosterpublic: [" + sDisplayRosterPublic.ToString() + "]<br />");
        Response.Write("rostergrade: [" + sRosterGrade.ToString() + "]<br />");
        Response.Write("rostershirtsize: [" + sRosterShirtSize.ToString() + "]<br />");
        Response.Write("rosterpantssize: [" + sRosterPantsSize.ToString() + "]<br />");
        Response.Write("rostercoachtype: [" + sRosterCoachType.ToString() + "]<br />");
        Response.Write("rostervolunteercoachname: [" + sRosterVolunteerCoachName.ToString() + "]<br />");
        Response.Write("rostervolunteercoachdayphone: [" + sRosterVolunteerCoachDayPhone.ToString() + "]<br />");
        Response.Write("rostervolunteercoachcellphone: [" + sRosterVolunteerCoachCellPhone.ToString() + "]<br />");
        Response.Write("rostervolunteercoachemail: [" + sRosterVolunteerCoachEmail.ToString() + "]<br />");
        Response.Write("cartid: [" + sCartID.ToString() + "]<br />");
        */
    }

    public static Int32 addToCart(Int32 iOrgID,
    //public static string addToCart(Int32 iOrgID,
                                  Int32 iClassID,
                                  Int32 iUserID,
                                  Int32 iTimeID,
                                  Int32 iFamilyMemberID,
                                  Int32 iQuantity,
                                  double iAmount,
                                  Int32 iPriceTypeID,
                                  Int32 iOptionID,
                                  string iBuyOrWait,
                                  Boolean iIsParent,
                                  Int32 iClassTypeID,
                                  Int32 iItemTypeID,
                                  Int32 iIsDropIn,
                                  string iDropInDate,
                                  Boolean iDisplayRosterPublic,
                                  string iRosterGrade,
                                  string iRosterShirtSize,
                                  string iRosterPantsSize,
                                  string iRosterCoachType,
                                  string iRosterVolunteerCoachName,
                                  string iRosterVolunteerCoachDayPhone,
                                  string iRosterVolunteerCoachCellPhone,
                                  string iRosterVolunteerCoachEmail)
    {
        Int32 lcl_return = 0;
        Int32 sIsParent  = 0;

        //string lcl_return = "";

        string sSQL = "";
        string sBuyOrWait = "";
        string sDropInDate = "";
        string sSessionID = "'" + HttpContext.Current.Session.SessionID + "'";

        if(iBuyOrWait != null) {
            sBuyOrWait = common.dbSafe(iBuyOrWait);
            sBuyOrWait = "'" + sBuyOrWait + "'";
        }
		else
		{
			sBuyOrWait = "null";
		}

        if (iIsParent)
        {
            sIsParent = 1;
        }

        if(iDropInDate != null)
        {
            sDropInDate = iDropInDate;
            sDropInDate = common.dbSafe(sDropInDate);
        }
        else
        {
            sDropInDate = "NULL";
        }

        sDropInDate = "'" + sDropInDate + "'";

        sSQL  = "INSERT INTO egov_class_cart (";
        sSQL += "classid,";
        sSQL += "userid,";
        sSQL += "classtimeid,";
        sSQL += "familymemberid,";
        sSQL += "quantity,";
        sSQL += "amount,";
        sSQL += "pricetypeid,";
        sSQL += "optionid,";
        sSQL += "buyorwait,";
        sSQL += "orgid,";
        sSQL += "sessionid_csharp,";
        sSQL += "isparent,";
        sSQL += "classtypeid,";
        sSQL += "itemtypeid,";
        sSQL += "dateadded,";
        sSQL += "isdropin,";
        sSQL += "dropindate";

        if (iDisplayRosterPublic)
        {
            sSQL += ",";
            sSQL += "rostergrade,";
            sSQL += "rostershirtsize,";
            sSQL += "rosterpantssize,";
            sSQL += "rostercoachtype,";
            sSQL += "rostervolunteercoachname,";
            sSQL += "rostervolunteercoachdayphone,";
            sSQL += "rostervolunteercoachcellphone,";
            sSQL += "rostervolunteercoachemail";
        }

        sSQL += ") VALUES (";

        sSQL += iClassID.ToString() + ", ";
        sSQL += iUserID.ToString() + ", ";
        sSQL += iTimeID.ToString() + ", ";
        sSQL += iFamilyMemberID + ", ";
        sSQL += iQuantity.ToString() + ", ";
        sSQL += iAmount.ToString() + ", ";
        sSQL += iPriceTypeID.ToString() + ", ";
        sSQL += iOptionID.ToString() + ", ";
        sSQL += sBuyOrWait + ", ";
        sSQL += iOrgID.ToString() + ", ";
        sSQL += sSessionID + ", ";
        sSQL += sIsParent.ToString() + ", ";
        sSQL += iClassTypeID.ToString() + ", ";
        sSQL += iItemTypeID.ToString() + ", ";
        sSQL += "dbo.getLocalDate(" + iOrgID.ToString() + ", getdate()), ";
        sSQL += iIsDropIn.ToString() + ", ";
        sSQL += sDropInDate;

        if(iDisplayRosterPublic)
        {
            sSQL += ",";
            sSQL += iRosterGrade + ", ";
            sSQL += iRosterShirtSize + ", ";
            sSQL += iRosterPantsSize + ", ";
            sSQL += iRosterCoachType + ", ";
            sSQL += iRosterVolunteerCoachName + ", ";
            sSQL += iRosterVolunteerCoachDayPhone + ", ";
            sSQL += iRosterVolunteerCoachCellPhone + ", ";
            sSQL += iRosterVolunteerCoachEmail;
        }

        sSQL += ");";

        //lcl_return = sSQL;

        lcl_return = Convert.ToInt32(common.RunInsertStatement(sSQL));

        return lcl_return;
    }
}

using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_classes_class_signup : System.Web.UI.Page
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
        common.logThePageVisit(startCounter, "class_signup.aspx", "public");
    }


    protected void Page_Load(object sender, EventArgs e)
    {

    }

    protected void displayClassSignUp(Int32 iOrgID,
                                      Int32 iUserID,
                                      Int32 iRootCategoryID,
                                      Int32 iCategoryID,
                                      //string iCategoryTitle,
                                      Int32 iClassID)
    {
        Boolean sCodeFinalFieldset              = true;
        Boolean sAllMembers                     = false;
        Boolean sCanRegister                    = false;
        Boolean sRegistrationStarted            = false;
        Boolean sIsUserNotBlocked               = true;
        Boolean sMemberRequirement              = false;
        Boolean sAgeRequirementsMet             = false;
        Boolean sClassFeesIncludeContainerDIV   = false;
        Boolean sClassFeesIncludeContainerTABLE = false;
        Boolean sSetupCoachFields               = false;
        Boolean sShowWaiverText                 = true;
        Boolean sShowWaiverName                 = true;
        Boolean sShowWaiverDesc                 = true;
        Boolean sShowWaiverLink                 = true;

        Boolean iOrgHasDisplay_classTeamRegistrationTshirtLabel        = common.orgHasDisplay(Convert.ToString(iOrgID), "class_teamregistration_tshirt_label");
        Boolean iOrgHasDisplay_classTeamRegistrationVolunteerCoachDesc = common.orgHasDisplay(Convert.ToString(iOrgID), "class_teamregistration_volunteercoachdesc");

        Boolean iOrgHasFeature_genderRestriction         = common.orgHasFeature(Convert.ToString(iOrgID), "gender restriction");
        Boolean iOrgHasFeature_registrationBlocking      = common.orgHasFeature(Convert.ToString(iOrgID), "registration blocking");
        Boolean sOrgHasFeatureEmergencyInfoRequired      = common.orgHasFeature(Convert.ToString(iOrgID), "emergency info required");
        Boolean iOrgHasFeature_customRegistrationCraigCO = common.orgHasFeature(Convert.ToString(iOrgID), "custom_registration_craigco");

        double sMinAge = 0;
        double sMaxAge = 0;

        Int32 sGenderNotRequiredID    = classes.getGenderNotRequiredID();
        Int32 iMemberCount            = 0;
        Int32 iSelectedFamilyMemberID = 0;

        string sSQL                   = "";
        string sClassName             = "";
        string sEnabledTshirt         = "BOTH";
        string sEnabledPants          = "BOTH";
        string sEnabledGrade          = "BOTH";
        string sEnabledCoach          = "BOTH";
        string sInputTypeTshirt       = "LOV";
        string sInputTypePants        = "LOV";
        string sInputTypeGrade        = "TEXT";
        string sUserType              = "";
        string sStartDate             = "";
        string sEndDate               = "";
        string sRegistrationEndDate   = "";
        string sRegStartDate          = "";
        string sPriceType             = "";
        string sAgeRestrictionsDate   = "&nbsp;";
        string sGenderRestriction     = "N";
        string sGenderRestrictionText = "";
        string sClassFeeLine          = "";
        string sDropDownList_familyMembers = "";
        string sCostOptions                = "";
        string sEmergencyInfo              = "";
        string sTeamRosterType             = "";
        string sTeamRosterName             = "";
        string sDisplayTeamRosterTshirt    = "";
        string sDisplayTeamRosterPants     = "";
        string sDisplayTeamRosterGrade     = "";
        string sLabelTshirt                = "T-Shirt";
        string sVolunteerCoachText         = "";
        string sDisplayActivityTime        = "";
        string sDisplayWaiverList          = "";
        string sDisplayTerms               = "";
        string sDisplayPublicBlockedNote   = "";

        System.TimeSpan sRegEndDateDiff = DateTime.Now - DateTime.Now;

        sUserType = classes.getUserResidentType(iUserID);

					Response.Write("<!--TWF:" + sUserType +"-->");
	if (sUserType == "E" && iOrgID == 60)
		sUserType = "R";

        if (sUserType != "R" && sUserType != "N" && sUserType != "U")
        {
            sUserType = classes.getResidentTypeByAddress(iUserID,
                                                         iOrgID);
        }

        //sResidentDesc = getResidentTypeDesc(sUserType);   <<--- still need to create when needed.

        Response.Write("<div>");
        Response.Write("  <input type=\"button\" name=\"class_signup_returnButtonDetails\" id=\"class_signup_returnButtonDetails\" value=\"Return to Details\" onclick=\"goToDetails();\" />");
        Response.Write("</div>");

        sSQL  = "SELECT c.classname, ";
        sSQL += " isnull(c.startdate,'') as startdate, ";
        sSQL += " c.classdescription, ";
        sSQL += " isnull(c.enddate,'') as enddate, ";
        sSQL += " isnull(minageprecisionid,0) as minageprecisionid, ";
        sSQL += " isnull(maxageprecisionid,0) as maxageprecisionid, ";
        sSQL += " o.optionid, ";
        sSQL += " o.optionname, ";
        sSQL += " o.optiondescription, ";
        sSQL += " c.isparent, ";
        sSQL += " c.classtypeid, ";
        sSQL += " isnull(c.membershipid,0) as membershipid, ";
        sSQL += " c.agecomparedate, ";
        sSQL += " l.name as locationname, ";
        sSQL += " l.address1, ";
        sSQL += " isnull(c.minage,0) as minage, ";
        sSQL += " isnull(c.maxage,99) as maxage, ";
        sSQL += " isnull(c.pricediscountid,0) as pricediscountid, ";
        sSQL += " allowearlyregistration, ";
        sSQL += " earlyregistrationdate, ";
        sSQL += " earlyregistrationclassid, ";
        sSQL += " displayrosterpublic, ";
        sSQL += " teamreg_tshirt_enabled, ";
        sSQL += " teamreg_pants_enabled, ";
        sSQL += " teamreg_grade_enabled, ";
        sSQL += " teamreg_coach_enabled, ";
        sSQL += " teamreg_tshirt_inputtype, ";
        sSQL += " teamreg_pants_inputtype, ";
        sSQL += " teamreg_grade_inputtype, ";
        sSQL += " showTerms, ";
        sSQL += " isnull(genderrestrictionid," + sGenderNotRequiredID + ") as genderrestrictionid ";
        sSQL += " FROM egov_class c, ";
        sSQL +=      " egov_registration_option o, ";
        sSQL +=      " egov_class_location l, ";
        sSQL +=      " egov_class_seasons s ";
        sSQL += " WHERE c.optionid = o.optionid ";
        sSQL += " AND c.locationid = l.locationid ";
        sSQL += " AND c.classseasonid = s.classseasonid ";
        sSQL += " AND s.showpublic = 1 ";
        sSQL += " AND classid = " + iClassID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            //Setup values
            sClassName = common.decodeUTFString(myReader["classname"].ToString().Trim());

            if (myReader["teamreg_tshirt_enabled"].ToString() != "")
            {
                sEnabledTshirt = myReader["teamreg_tshirt_enabled"].ToString().ToUpper();
            }

            if (myReader["teamreg_pants_enabled"].ToString() != "")
            {
                sEnabledPants = myReader["teamreg_pants_enabled"].ToString().ToUpper();
            }

            if (myReader["teamreg_grade_enabled"].ToString() != "")
            {
                sEnabledGrade = myReader["teamreg_grade_enabled"].ToString().ToUpper();
            }

            if (myReader["teamreg_coach_enabled"].ToString() != "")
            {
                sEnabledCoach = myReader["teamreg_coach_enabled"].ToString().ToUpper();
            }

            if (myReader["teamreg_tshirt_inputtype"].ToString() != "")
            {
                sInputTypeTshirt = myReader["teamreg_tshirt_inputtype"].ToString().ToUpper();
            }

            if (myReader["teamreg_pants_inputtype"].ToString() != "")
            {
                sInputTypePants = myReader["teamreg_pants_inputtype"].ToString().ToUpper();
            }

            if (myReader["teamreg_grade_inputtype"].ToString() != "")
            {
                sInputTypeGrade = myReader["teamreg_grade_inputtype"].ToString().ToUpper();
            }

            //Get Member Information
            sAllMembers = classes.getMemberInformation(iUserID,
                                                       Convert.ToInt32(myReader["membershipid"]),
                                                       iMemberCount,
                                                       out iMemberCount);

            //Build the Start and End Dates
            if (myReader["startdate"].ToString() != "")
            {
                sStartDate = string.Format("{0:MMMM d}", Convert.ToDateTime(myReader["startdate"]));
            }

            if (myReader["enddate"].ToString() != "")
            {
                if (Convert.ToDateTime(myReader["enddate"]) != Convert.ToDateTime(myReader["startdate"]))
                {
                    sEndDate = " - " + string.Format("{0:MMMM d}", Convert.ToDateTime(myReader["enddate"]));
                }
            }

            //Check that we are still able to register
            sRegistrationEndDate = classes.getRegistrationEndDate(iClassID);

            if (sRegistrationEndDate == "")
            {
                sCanRegister = true;
            }
            else
            {
                sRegEndDateDiff = Convert.ToDateTime(sRegistrationEndDate) - DateTime.Now;
                //Response.Write("<br />" + sRegistrationEndDate + " - " + DateTime.Now.ToString() + " [" + sRegEndDateDiff.TotalDays + "]<br />");

                if (sRegEndDateDiff.TotalDays > 0)
                {
                    sCanRegister = true;
                }
            }

            Response.Write("<fieldset class=\"class_signup_fieldset\">");
            Response.Write("  <legend>" + sClassName + "</legend>");
            Response.Write("  <div id=\"class_signup_dates\">" + sStartDate + sEndDate + "</div>");

            if (!sCanRegister)
            {
                Response.Write("<div class=\"class_signup_message\">Registration for this event closed on " + string.Format("{0:M/d/yyyy}", Convert.ToDateTime(sRegistrationEndDate)) + ".</div>");
            }
            else
            {
                //If not yet normal registration period and early registration is available,
                //and going on, see if user qualifies.
                sRegistrationStarted = classes.registrationStarted(iClassID,
                                                                   iUserID,
                                                                   out sRegStartDate,
                                                                   out sPriceType);

                if (! sRegistrationStarted && Convert.ToBoolean(myReader["allowearlyregistration"]))
                {
                    if (Convert.ToDateTime(myReader["earlyregistrationdate"]) <= DateTime.Now)
                    {
                        sRegistrationStarted = classes.userCanRegisterEarly(iClassID,
                                                                            iUserID);
                    }
                }

                if (sRegistrationStarted)
                {
                    //Determine if registration, ticket, or free.
                    switch (Convert.ToInt32(myReader["optionid"]))
                    {
                        case 1:  //Handle registration required
                            //The "FIELDSET" that is opened above (class = "class_signup_fieldset") is closed
                            //within this CASE statement.  Therefore we do NOT need the original closing of
                            //it and therefore set sCodeFinalFieldset = FALSE.
                            sCodeFinalFieldset = false;

                            //See if they are blocked before going on.
                            sIsUserNotBlocked = classes.userNotBlocked(iOrgHasFeature_registrationBlocking,
                                                                       iUserID);
                            if (sIsUserNotBlocked)
                            {
                                sMemberRequirement = classes.checkMemberRequirement(iClassID,
                                                                                    iOrgID,
                                                                                    iMemberCount,
                                                                                    sUserType,
                                                                                    sAllMembers,
                                                                                    out sPriceType);

                                if (sMemberRequirement)
                                {
                                    sAgeRequirementsMet = classes.ageRequirementsMet(iUserID,
                                                                                     iClassID);
				    Response.Write("<!--HERE " + DateTime.Now.ToString() + "-->");

                                    //Form for selecting either ticket qty or selecting a family member
                                    Response.Write("<form name=\"PurchaseForm\" id=\"PurchaseForm\" method=\"post\" action=\"class_addtocart.aspx\">");
                                    Response.Write(  "<input type=\"hidden\" name=\"orgid\" id=\"orgid\" value=\"" + iOrgID.ToString() + "\" />");
                                    Response.Write(  "<input type=\"hidden\" name=\"classid\" id=\"classid\" value=\"" + iClassID.ToString() + "\" />");
                                    Response.Write(  "<input type=\"hidden\" name=\"userid\" id=\"userid\" value=\"" + iUserID.ToString() + "\" />");
                                    Response.Write(  "<input type=\"hidden\" name=\"optionid\" id=\"optionid\" value=\"" + myReader["optionid"].ToString() + "\" />");
                                    Response.Write(  "<input type=\"hidden\" name=\"isparent\" id=\"isparent\" value=\"" + myReader["isparent"].ToString() + "\" />");
                                    Response.Write(  "<input type=\"hidden\" name=\"classtypeid\" id=\"classtypeid\" value=\"" + myReader["classtypeid"].ToString() + "\" />");
                                    Response.Write(  "<input type=\"hidden\" name=\"classname\" id=\"classname\" value=\"" + myReader["classname"].ToString() + "\" />");
                                    Response.Write(  "<input type=\"hidden\" name=\"categoryid\" id=\"categoryid\" value=\"" + iCategoryID.ToString() + "\" />");
                                    //Response.Write(  "<input type=\"hidden\" name=\"categorytitle\" id=\"categorytitle\" value=\"" + iCategoryTitle + "\" />");
                                    Response.Write(  "<input type=\"hidden\" name=\"teamreg_tshirt_enabled\" id=\"teamreg_tshirt_enabled\" value=\"" + sEnabledTshirt + "\" />");
                                    Response.Write(  "<input type=\"hidden\" name=\"teamreg_pants_enabled\" id=\"teamreg_pants_enabled\" value=\"" + sEnabledPants + "\" />");
                                    Response.Write(  "<input type=\"hidden\" name=\"teamreg_grade_enabled\" id=\"teamreg_grade_enabled\" value=\"" + sEnabledGrade + "\" />");
                                    Response.Write(  "<input type=\"hidden\" name=\"teamreg_coach_enabled\" id=\"teamreg_coach_enabled\" value=\"" + sEnabledCoach + "\" />");
                                    Response.Write(  "<input type=\"hidden\" name=\"displayrosterpublic\" id=\"displayrosterpublic\" value=\"" + myReader["displayrosterpublic"].ToString() + "\" />");
                                    Response.Write(  "<input type=\"hidden\" name=\"quantity\" id=\"quantity\" value=\"1\" />");

                                    if (sOrgHasFeatureEmergencyInfoRequired)
                                    {
                                        //sEmergencyContact = getUserContactInfo(Convert.ToInt32(sUserID), "emergencycontact");
                                        //sEmergencyPhone = getUserContactInfo(Convert.ToInt32(sUserID), "emergencyphone");

                                        Response.Write("<input type=\"hidden\" name=\"emergencycontact\" id=\"emergencycontact\" value=\"\" size=\"30\" />");
                                        Response.Write("<input type=\"hidden\" name=\"emergencyphone\" id=\"emergencyphone\" value=\"\" size=\"30\" />");
                                    }

                                    Response.Write("<table id=\"class_signup_details\">");

                                    //Age Restrictions
                                    sMinAge = Convert.ToDouble(myReader["minage"]);
                                    sMaxAge = Convert.ToDouble(myReader["maxage"]);

                                    if (sMinAge != Convert.ToDouble(0) || sMaxAge != Convert.ToDouble(99))
                                    {
                                        if (Convert.ToString(myReader["agecomparedate"]) != "")
                                        {
                                            sAgeRestrictionsDate = "<strong>Age Restrictions:</strong>";
                                            sAgeRestrictionsDate += "<div class=\"class_signup_ageRestrictDate\">(as of " + string.Format("{0:M/d/yyyy}", Convert.ToDateTime(myReader["agecomparedate"])) + ")</div>";
                                        }

                                        Response.Write("  <tr valign=\"top\">");
                                        Response.Write("      <td>" + sAgeRestrictionsDate + "</td>");
                                        Response.Write("      <td>");
                                        Response.Write("          <table id=\"class_signup_age_restriction\">");

                                        if (sMinAge != Convert.ToDouble(0))
                                        {
                                            Response.Write("          <tr>");
                                            Response.Write("              <td class=\"age_restrictions_label\">Minimum:</td>");
                                            Response.Write("              <td>" + sMinAge.ToString() + " years of age</td>");
                                            Response.Write("          </tr>");
                                        }

                                        if (sMaxAge != Convert.ToDouble(99))
                                        {
                                            Response.Write("          <tr>");
                                            Response.Write("              <td class=\"age_restrictions_label\">Maximum:</td>");
                                            Response.Write("              <td>" + sMaxAge.ToString() + " years of age</td>");
                                            Response.Write("          </tr>");
                                        }

                                        Response.Write("          </table>");
                                        Response.Write("      </td>");
                                        Response.Write("  </tr>");
                                    }
                                    
                                    //Gender Restrictions
                                    if (iOrgHasFeature_genderRestriction)
                                    {
                                        if (Convert.ToInt32(myReader["genderrestrictionid"]) != sGenderNotRequiredID)
                                        {
                                            sGenderRestrictionText = classes.getGenderRestrictionText(Convert.ToInt32(myReader["genderrestrictionid"]));

                                            Response.Write("  <tr>");
                                            Response.Write("      <td><strong>Gender Restriction:</strong></td>");
                                            Response.Write("      <td>" + sGenderRestrictionText + "</td>");
                                            Response.Write("  </tr>");
                                        }
                                    }

				    bool bMeetsGenderRequirement = true;
                                    if (sAgeRequirementsMet)
                                    {
                                        sClassFeeLine = classes.buildClassFeeLine(iClassID,
                                                                                  sClassFeesIncludeContainerDIV,
                                                                                  sClassFeesIncludeContainerTABLE);

                                        if (iOrgHasFeature_genderRestriction)
                                        {
                                            sGenderRestriction = classes.getGenderRestriction(Convert.ToInt32(myReader["genderrestrictionid"]));

					    //Response.Write("<!--GS:" + sGenderRestriction + " " + iUserID + "-->");
					    if (sGenderRestriction == "M" || sGenderRestriction == "F")
					    {
					    	bMeetsGenderRequirement = classes.genderRequirementsMet(iUserID, sGenderRestriction);
					    }
                                        }

                                        //Class Fees
                                        Response.Write(sClassFeeLine);
                                    }

                                    Response.Write("</table>");
                                    Response.Write("</fieldset>");
                                    Response.Write("<div class=\"classdetails\">");

                                    if (sAgeRequirementsMet && bMeetsGenderRequirement)
                                    {
                                        //Select a Family Member
                                        sAllMembers = classes.showFamilyMembers(iUserID,
                                                                                iOrgID,
                                                                                iMemberCount,
                                                                                sMinAge,
                                                                                sMaxAge,
                                                                                Convert.ToInt32(myReader["membershipid"]),
                                                                                Convert.ToString(myReader["agecomparedate"]),
                                                                                Convert.ToInt32(myReader["minageprecisionid"]),
                                                                                Convert.ToInt32(myReader["maxageprecisionid"]),
                                                                                sGenderRestriction,
                                                                                out iMemberCount,
                                                                                out iSelectedFamilyMemberID,
                                                                                out sDropDownList_familyMembers);

                                        //Cost Options
                                        sCostOptions = classes.showCostOptions(iClassID,
                                                                               sUserType,
                                                                               iOrgID,
                                                                               sAllMembers,
                                                                               iMemberCount,
                                                                               Convert.ToInt32(myReader["pricediscountid"]));

                                        //Emergency Info Required
                                        sEmergencyInfo  = "<div id=\"classEmergencyInfoDiv\">";
                                        sEmergencyInfo +=   classes.showEmergencyInfo(iOrgID,
                                                                                      iSelectedFamilyMemberID);
                                        sEmergencyInfo += "</div>";

                                        Response.Write("<fieldset class=\"class_signup_fieldset\">");
                                        Response.Write("  <legend>Select a Family Member to Register*</legend>");
                                        Response.Write(sDropDownList_familyMembers);
                                        Response.Write(sEmergencyInfo);
                                        Response.Write(sCostOptions);
                                        Response.Write("</fieldset>");

                                        //Team Registration
                                        if (iOrgHasFeature_customRegistrationCraigCO && Convert.ToBoolean(myReader["displayrosterpublic"]))
                                        {
                                            if (sEnabledTshirt == "BOTH" ||
                                                sEnabledPants == "BOTH" ||
                                                sEnabledGrade == "BOTH" ||
                                                sEnabledCoach == "BOTH")
                                            {
                                                Response.Write("<fieldset class=\"class_signup_fieldset\">");
                                                Response.Write("  <legend>Team Registration - Additional Info</legend>");

                                                if (sEnabledTshirt == "BOTH" ||
                                                    sEnabledPants == "BOTH" ||
                                                    sEnabledGrade == "BOTH")
                                                {
                                                    Response.Write("<table width=\"100%\">");

                                                    //Grade
                                                    if (sEnabledGrade == "BOTH")
                                                    {
                                                        sTeamRosterType = "GRADE";
                                                        sTeamRosterName = "rostergrade";
                                                        sDisplayTeamRosterGrade = classes.buildTeamRosterAccessories(iOrgID,
                                                                                                                     iClassID,
                                                                                                                     sEnabledGrade,
                                                                                                                     sInputTypeGrade,
                                                                                                                     sTeamRosterType,
                                                                                                                     sTeamRosterName);

                                                        Response.Write("  <tr>");
                                                        Response.Write("      <td class=\"requiredField\" align=\"right\">*</td>");
                                                        Response.Write("      <td class=\"coachFieldsLabel\">Grade:</td>");
                                                        Response.Write("      <td width=\"90%\">" + sDisplayTeamRosterGrade + "</td>");
                                                        Response.Write("  </tr>");
                                                    }

                                                    //T-Shirt
                                                    if (sEnabledTshirt == "BOTH")
                                                    {
                                                        sTeamRosterType = "TSHIRT";
                                                        sTeamRosterName = "rostershirtsize";
                                                        sDisplayTeamRosterTshirt = classes.buildTeamRosterAccessories(iOrgID,
                                                                                                                     iClassID,
                                                                                                                     sEnabledTshirt,
                                                                                                                     sInputTypeTshirt,
                                                                                                                     sTeamRosterType,
                                                                                                                     sTeamRosterName);

                                                        if (iOrgHasDisplay_classTeamRegistrationTshirtLabel)
                                                        {
                                                            sLabelTshirt = common.getOrgDisplay(Convert.ToString(iOrgID), "class_teamregistration_tshirt_label");
                                                        }

                                                        Response.Write("  <tr>");
                                                        Response.Write("      <td>&nbsp;</td>");
                                                        Response.Write("      <td class=\"coachFieldsLabel\">" + sLabelTshirt + " Size:</td>");
                                                        Response.Write("      <td>" + sDisplayTeamRosterTshirt + "</td>");
                                                        Response.Write("  </tr>");
                                                    }

                                                    //Pants
                                                    if (sEnabledPants == "BOTH")
                                                    {
                                                        sTeamRosterType = "PANTS";
                                                        sTeamRosterName = "rosterpantssize";
                                                        sDisplayTeamRosterPants = classes.buildTeamRosterAccessories(iOrgID,
                                                                                                                     iClassID,
                                                                                                                     sEnabledPants,
                                                                                                                     sInputTypePants,
                                                                                                                     sTeamRosterType,
                                                                                                                     sTeamRosterName);

                                                        Response.Write("  <tr>");
                                                        Response.Write("      <td>&nbsp;</td>");
                                                        Response.Write("      <td class=\"coachFieldsLabel\">Pants:</td>");
                                                        Response.Write("      <td>" + sDisplayTeamRosterPants + "</td>");
                                                        Response.Write("  </tr>");
                                                    }

                                                    //Coach
                                                    if (sEnabledCoach == "BOTH")
                                                    {
                                                        if (iOrgHasDisplay_classTeamRegistrationVolunteerCoachDesc)
                                                        {
                                                            sVolunteerCoachText = common.getOrgDisplay(Convert.ToString(iOrgID), "class_teamregistration_volunteercoachdesc");
                                                        }

                                                        if (sVolunteerCoachText != "")
                                                        {
                                                            Response.Write("  <tr>");
                                                            Response.Write("      <td colspan=\"3\">" + sVolunteerCoachText + "</td>");
                                                            Response.Write("  </tr>");
                                                        }

                                                        Response.Write("  <tr>");
                                                        Response.Write("      <td colspan=\"2\" class=\"coachFieldsLabel\">I would like to:</td>");
                                                        Response.Write("      <td>");
                                                        Response.Write("          <select name=\"rostercoachtype\" id=\"rostercoachtype\" onchange=\"clearMsg('rostercoachtype');setupCoachFields();\">");
                                                        Response.Write("            <option value=\"\"></option>");
                                                        Response.Write("            <option value=\"Head Coach\">Head Coach</option>");
                                                        Response.Write("            <option value=\"Assistant Coach\">Assistant Coach</option>");
                                                        Response.Write("          </select>");
                                                        Response.Write("      </td>");
                                                        Response.Write("  </tr>");
                                                        Response.Write("  <tr>");
                                                        Response.Write("      <td colspan=\"3\">");
                                                        Response.Write("          <div id=\"volunteerCoachInfo\">");
                                                        Response.Write("            <table cellspacing=\"0\" cellpadding=\"2\">");
                                                        Response.Write("              <tr>");
                                                        Response.Write("                  <td class=\"requiredField\">*</td>");
                                                        Response.Write("                  <td class=\"coachFieldsLabel\">Coach's Name:</td>");
                                                        Response.Write("                  <td width=\"85%\">");
                                                        Response.Write("                      <input type=\"text\" name=\"rostervolunteercoachname\" id=\"rostervolunteercoachname\" maxlength=\"100\" class=\"coachLargeInputField\" onchange=\"clearMsg('rostervolunteercoachname');\" />");
                                                        Response.Write("                  </td>");
                                                        Response.Write("              </tr>");
                                                        Response.Write("              <tr>");
                                                        Response.Write("                  <td>&nbsp;</td>");
                                                        Response.Write("                  <td class=\"coachFieldsLabel\">Daytime Phone:</td>");
                                                        Response.Write("                  <td>");
                                                        Response.Write("                     (<input type=\"text\" name=\"skip_volcoachday_areacode\" id=\"skip_volcoachday_areacode\" size=\"3\" maxlength=\"3\" onKeyUp=\"return autoTab(this, 3, event);\" onchange=\"clearMsg('skip_volcoachday_line');\" />)");
                                                        Response.Write("                      <input type=\"text\" name=\"skip_volcoachday_exchange\" id=\"skip_volcoachday_exchange\" size=\"3\" maxlength=\"3\" onKeyUp=\"return autoTab(this, 3, event);\" onchange=\"clearMsg('skip_volcoachday_line');\" />");
                                                        Response.Write("                      &ndash;");
                                                        Response.Write("                      <input type=\"text\" name=\"skip_volcoachday_line\" id=\"skip_volcoachday_line\" size=\"4\" maxlength=\"4\" onKeyUp=\"return autoTab(this, 4, event);\" onchange=\"clearMsg('skip_volcoachday_line');\" />");
                                                        Response.Write("                  </td>");
                                                        Response.Write("              </tr>");
                                                        Response.Write("              <tr>");
                                                        Response.Write("                  <td>&nbsp;</td>");
                                                        Response.Write("                  <td class=\"coachFieldsLabel\">Cell Phone:</td>");
                                                        Response.Write("                  <td>");
                                                        Response.Write("                     (<input type=\"text\" name=\"skip_volcoachcell_areacode\" id=\"skip_volcoachcell_areacode\" size=\"3\" maxlength=\"3\" onKeyUp=\"return autoTab(this, 3, event);\" onchange=\"clearMsg('skip_volcoachcell_line');\" />)");
                                                        Response.Write("                      <input type=\"text\" name=\"skip_volcoachcell_exchange\" id=\"skip_volcoachcell_exchange\" size=\"3\" maxlength=\"3\" onKeyUp=\"return autoTab(this, 3, event);\" onchange=\"clearMsg('skip_volcoachcell_line');\" />");
                                                        Response.Write("                      &ndash;");
                                                        Response.Write("                      <input type=\"text\" name=\"skip_volcoachcell_line\" id=\"skip_volcoachcell_line\" size=\"4\" maxlength=\"4\" onKeyUp=\"return autoTab(this, 4, event);\" onchange=\"clearMsg('skip_volcoachcell_line');\" />");
                                                        Response.Write("                  </td>");
                                                        Response.Write("              </tr>");
                                                        Response.Write("            </table>");
                                                        Response.Write("            <div>");
                                                        Response.Write("              Please list an email address, so you can be contacted for more information:<br />");
                                                        Response.Write("              <input type=\"text\" name=\"rostervolunteercoachemail\" id=\"rostervolunteercoachemail\" maxlength=\"100\" class=\"coachLargeInputField\" onchange=\"clearMsg('rostervolunteercoachemail');\" />");
                                                        Response.Write("            </div>");
                                                        Response.Write("          </div>");
                                                        Response.Write("      </td>");
                                                        Response.Write("  </tr>");

                                                        sSetupCoachFields = true;
                                                    }

                                                    Response.Write("</table>");
                                                }

                                                //Determine if there are any HIDDEN fields we need to build.
                                                if (sEnabledGrade != "BOTH")
                                                {
                                                    Response.Write("<input type=\"hidden\" name=\"rostergrade\" id=\"rostergrade\" value=\"\" />");
                                                }

                                                if (sEnabledTshirt != "BOTH")
                                                {
                                                    Response.Write("<input type=\"hidden\" name=\"rostershirtsize\" id=\"rostershirtsize\" value=\"\" />");
                                                }

                                                if (sEnabledPants != "BOTH")
                                                {
                                                    Response.Write("<input type=\"hidden\" name=\"rosterpantssize\" id=\"rosterpantssize\" value=\"\" />");
                                                }

                                                Response.Write("</fieldset>");
                                            }
                                        }
                                        else
                                        {
                                            Response.Write("<input type=\"hidden\" name=\"rostergrade\" id=\"rostergrade\" value=\"\" />");
                                            Response.Write("<input type=\"hidden\" name=\"rostershirtsize\" id=\"rostershirtsize\" value=\"\" />");
                                            Response.Write("<input type=\"hidden\" name=\"rosterpantssize\" id=\"rosterpantssize\" value=\"\" />");
                                            Response.Write("<input type=\"hidden\" name=\"rostercoachtype\" id=\"rostercoachtype\" value=\"\" />");
                                            Response.Write("<input type=\"hidden\" name=\"rostervolunteercoachname\" id=\"rostervolunteercoachname\" value=\"\" />");
                                            Response.Write("<input type=\"hidden\" name=\"skip_volcoachday_areacode\" id=\"skip_volcoachday_areacode\" value=\"\" />");
                                            Response.Write("<input type=\"hidden\" name=\"skip_volcoachday_exchange\" id=\"skip_volcoachday_exchange\" value=\"\" />");
                                            Response.Write("<input type=\"hidden\" name=\"skip_volcoachday_line\" id=\"skip_volcoachday_line\" value=\"\" />");
                                            Response.Write("<input type=\"hidden\" name=\"skip_volcoachcell_areacode\" id=\"skip_volcoachcell_areacode\" value=\"\" />");
                                            Response.Write("<input type=\"hidden\" name=\"skip_volcoachcell_exchange\" id=\"skip_volcoachcell_exchange\" value=\"\" />");
                                            Response.Write("<input type=\"hidden\" name=\"skip_volcoachcell_line\" id=\"skip_volcoachcell_line\" value=\"\" />");
                                            Response.Write("<input type=\"hidden\" name=\"rostervolunteercoachemail\" id=\"rostervolunteercoachemail\" value=\"\" />");
                                        }

                                        //Availability
                                        sDisplayActivityTime = classes.displayActivityTimes(iClassID,
                                                                                            Convert.ToBoolean(myReader["isParent"]));

                                        if (sDisplayActivityTime != "")
                                        {
                                            Response.Write("<fieldset class=\"class_signup_fieldset\">");
                                            Response.Write("  <legend>Select an Activity Time</legend>");
                                    		Response.Write("<div class=\"offeringActivitiesContainer\">");
                                            Response.Write(sDisplayActivityTime);
                                            Response.Write("</div>");
                                            Response.Write("</fieldset>");
                                        }

                                        //Waivers
                                        sDisplayWaiverList = classes.showWaiverList(iOrgID,
                                                                                    iClassID,
                                                                                    sShowWaiverText,
                                                                                    sShowWaiverName,
                                                                                    sShowWaiverDesc,
                                                                                    sShowWaiverLink);
                                        
                                        if (sDisplayWaiverList != "")
                                        {
                                            Response.Write("<fieldset class=\"class_signup_fieldset_waivers\">");
                                            Response.Write("  <legend>Waivers</legend>");
                                            Response.Write(sDisplayWaiverList);
                                            Response.Write("</fieldset>");
                                        }

                                        //Terms
                                        if (Convert.ToBoolean(myReader["showTerms"]))
                                        {
                                            sDisplayTerms = classes.showTermsList(iOrgID);

                                            Response.Write("<fieldset class=\"class_signup_fieldset\">");
                                            Response.Write("  <legend>Terms</legend>");
                                            Response.Write(sDisplayTerms);
                                            Response.Write("</fieldset>");
                                        }

                                        //Add to Cart
                                        Response.Write("<input type=\"button\" value=\"Add to Cart\" name=\"addToCartButton\" id=\"addToCartButton\" />");
                                    }
                                    else
                                    {
                                        Response.Write("<div id=\"class_signup_selectfamilymembers\">");
					if (!sAgeRequirementsMet)
					{
                                        	Response.Write("<p class=\"requiredField\">The age requirement could not be met by anyone in your family.</p>");
					}
					if (!bMeetsGenderRequirement)
					{
                                        	Response.Write("<p class=\"requiredField\">The gender requirement could not be met by anyone in your family.</p>");
					}
                                        Response.Write("<input type=\"button\" name=\"updateFamilyMembersButton\" id=\"updateFamilyMembersButton\" value=\"Update Family Members\" onclick=\"updateFamily('" + iUserID.ToString() + "')\" />");
                                        Response.Write("</div>");
                                    }

                                    Response.Write("</div>");
                                    Response.Write("</form>");
                                }
                                else
                                {
                                    //Failure to meet residency or membership requirements
                                    Response.Write("<div class=\"class_signup_filledStatus\">You do not meet the requirements to purchase this.</div>");

                                    if (sPriceType == "M")
                                    {
                                        Response.Write("<div class=\"class_signup_filledStatus\">Non Member Registration Is Not Available.</div>");
                                    }

                                    if (sPriceType == "R")
                                    {
                                        Response.Write("<div class=\"class_signup_filledStatus\">Residency is required.</div>");
                                    }

                                    Response.Write("<div>");
                                    Response.Write("<input type=\"button\" name=\"class_signup_returnButtonClassList\" id=\"class_signup_returnButtonClassList\" value=\"Return to Class List\" onclick=\"goToList()\" />");
                                    Response.Write("</div>");
                                }
                            }
                            else
                            {
                                //Blocked Message
                                sDisplayPublicBlockedNote = classes.displayPublicBlockedNote(iUserID);

                                Response.Write("<div class=\"class_signup_filledStatus\">Your account has been blocked from online purchases.</div>");
                                Response.Write(sDisplayPublicBlockedNote);
                            }

                            break;
                        case 2:  //Ticketed event

                            sIsUserNotBlocked = classes.userNotBlocked(iOrgHasFeature_registrationBlocking,
                                                                       iUserID);

                            if (sIsUserNotBlocked)
                            {
                                sMemberRequirement = classes.checkMemberRequirement(iClassID,
                                                                                    iOrgID,
                                                                                    iMemberCount,
                                                                                    sUserType,
                                                                                    sAllMembers,
                                                                                    out sPriceType);

                                if (sMemberRequirement)
                                {
                                    Response.Write("<form name=\"PurchaseForm\" id=\"PurchaseForm\" method=\"post\" action=\"class_addtocart.aspx\">");
                                    Response.Write("  <input type=\"hidden\" name=\"orgid\" id=\"orgid\" value=\"" + iOrgID.ToString() + "\" />");
                                    Response.Write("  <input type=\"hidden\" name=\"classid\" id=\"classid\" value=\"" + iClassID.ToString() + "\" />");
                                    Response.Write("  <input type=\"hidden\" name=\"userid\" id=\"userid\" value=\"" + iUserID.ToString() + "\" />");
                                    Response.Write("  <input type=\"hidden\" name=\"optionid\" id=\"optionid\" value=\"" + myReader["optionid"].ToString() + "\" />");
                                    Response.Write("  <input type=\"hidden\" name=\"isparent\" id=\"isparent\" value=\"" + myReader["isparent"].ToString() + "\" />");
                                    Response.Write("  <input type=\"hidden\" name=\"classtypeid\" id=\"classtypeid\" value=\"" + myReader["classtypeid"].ToString() + "\" />");
                                    Response.Write("  <input type=\"hidden\" name=\"classname\" id=\"classname\" value=\"" + myReader["classname"].ToString() + "\" />");
                                    Response.Write("  <input type=\"hidden\" name=\"categoryid\" id=\"categoryid\" value=\"" + iCategoryID.ToString() + "\" />");
                                    //Response.Write("  <input type=\"hidden\" name=\"categorytitle\" id=\"categorytitle\" value=\"" + iCategoryTitle + "\" />");
                                    Response.Write("  <input type=\"hidden\" name=\"displayrosterpublic\" id=\"displayrosterpublic\" value=\"" + myReader["displayrosterpublic"].ToString() + "\" />");

                                    if (sOrgHasFeatureEmergencyInfoRequired)
                                    {
                                        //sEmergencyContact = getUserContactInfo(Convert.ToInt32(sUserID), "emergencycontact");
                                        //sEmergencyPhone = getUserContactInfo(Convert.ToInt32(sUserID), "emergencyphone");

                                        Response.Write("<input type=\"hidden\" name=\"emergencycontact\" id=\"emergencycontact\" value=\"\" size=\"30\" />");
                                        Response.Write("<input type=\"hidden\" name=\"emergencyphone\" id=\"emergencyphone\" value=\"\" size=\"30\" />");
                                    }

                                    Response.Write("<table id=\"class_signup_details\">");

                                    //Age Restrictions
                                    sMinAge = Convert.ToDouble(myReader["minage"]);
                                    sMaxAge = Convert.ToDouble(myReader["maxage"]);

                                    if (sMinAge != Convert.ToDouble(0) || sMaxAge != Convert.ToDouble(99))
                                    {
                                        if (Convert.ToString(myReader["agecomparedate"]) != "")
                                        {
                                            sAgeRestrictionsDate = "<strong>Age Restrictions:</strong>";
                                            sAgeRestrictionsDate += "<div class=\"class_signup_ageRestrictDate\">(as of " + string.Format("{0:M/d/yyyy}", Convert.ToDateTime(myReader["agecomparedate"])) + ")</div>";
                                        }

                                        Response.Write("  <tr valign=\"top\">");
                                        Response.Write("      <td>" + sAgeRestrictionsDate + "</td>");
                                        Response.Write("      <td>");
                                        Response.Write("          <table id=\"class_signup_age_restriction\">");

                                        if (sMinAge != Convert.ToDouble(0))
                                        {
                                            Response.Write("          <tr>");
                                            Response.Write("              <td class=\"age_restrictions_label\">Minimum:</td>");
                                            Response.Write("              <td>" + sMinAge.ToString() + " years of age</td>");
                                            Response.Write("          </tr>");
                                        }

                                        if (sMaxAge != Convert.ToDouble(99))
                                        {
                                            Response.Write("          <tr>");
                                            Response.Write("              <td class=\"age_restrictions_label\">Maximum:</td>");
                                            Response.Write("              <td>" + sMaxAge.ToString() + " years of age</td>");
                                            Response.Write("          </tr>");
                                        }

                                        Response.Write("          </table>");
                                        Response.Write("      </td>");
                                        Response.Write("  </tr>");
                                    }

                                    //Class Fees
                                    sClassFeeLine = classes.buildClassFeeLine(iClassID,
                                                                              sClassFeesIncludeContainerDIV,
                                                                              sClassFeesIncludeContainerTABLE);


                                    Response.Write(sClassFeeLine);
                                    Response.Write("</table>");

                                    //Cost Options
                                    sCostOptions = classes.showCostOptions(iClassID,
                                                                           sUserType,
                                                                           iOrgID,
                                                                           sAllMembers,
                                                                           iMemberCount,
                                                                           Convert.ToInt32(myReader["pricediscountid"]));

                                    //Emergency Info Required
                                    iSelectedFamilyMemberID = classes.getCitizenFamilyID(iUserID);

                                    sEmergencyInfo = "<div id=\"classEmergencyInfoDiv\">";
                                    sEmergencyInfo += classes.showEmergencyInfo(iOrgID,
                                                                                iSelectedFamilyMemberID);
                                    sEmergencyInfo += "</div>";

                                    Response.Write(sCostOptions);
                                    Response.Write(sEmergencyInfo);

                                    //Response.Write("</table>");
                                    Response.Write("</fieldset>");
                                    Response.Write("<div class=\"classdetails\">");

                                    //Availability
                                    sDisplayActivityTime = classes.displayActivityTimes(iClassID,
                                                                                        Convert.ToBoolean(myReader["isParent"]));

                                    Response.Write("<fieldset class=\"class_signup_fieldset\">");
                                    Response.Write("  <legend>Select an Activity Time</legend>");
                                    Response.Write("<div class=\"offeringActivitiesContainer\">");
                                    Response.Write(sDisplayActivityTime);
                                    Response.Write("</div>");
                                    Response.Write("</fieldset>");

                                    //Ticket Quantity
                                    Response.Write("<fieldset class=\"class_signup_fieldset\">");
                                    Response.Write("  <legend>Tickets</legend>");
                                    Response.Write("  <strong>Tickets: </strong>&nbsp;&nbsp;<input type=\"text\" name=\"quantity\" id=\"quantity\" value=\"1\" size=\"6\" maxlength=\"6\" onchange=\"clearMsg('quantity')\" />");
                                    Response.Write("</fieldset>");

                                    //Waivers
                                    sDisplayWaiverList = classes.showWaiverList(iOrgID,
                                                                                iClassID,
                                                                                sShowWaiverText,
                                                                                sShowWaiverName,
                                                                                sShowWaiverDesc,
                                                                                sShowWaiverLink);

                                    if (sDisplayWaiverList != "")
                                    {
                                        Response.Write("<fieldset class=\"class_signup_fieldset_waivers\">");
                                        Response.Write("  <legend>Waivers</legend>");
                                        Response.Write(sDisplayWaiverList);
                                        Response.Write("</fieldset>");
                                    }

                                    //Terms
                                    if (Convert.ToBoolean(myReader["showTerms"]))
                                    {
                                        sDisplayTerms = classes.showTermsList(iOrgID);

                                        Response.Write("<fieldset class=\"class_signup_fieldset\">");
                                        Response.Write("  <legend>Terms</legend>");
                                        Response.Write(sDisplayTerms);
                                        Response.Write("</fieldset>");
                                    }

                                    //Add to Cart
                                    Response.Write("<input type=\"button\" value=\"Add to Cart\" name=\"addToCartButton\" id=\"addToCartButton\" />");

                                    Response.Write("</div>");
                                    Response.Write("</form>");
                                }
                                else
                                {
                                    //Failure to meet residency or membership requirements
                                    Response.Write("<div class=\"class_signup_filledStatus\">You do not meet the requirements to purchase this.</div>");

                                    if (sPriceType == "M")
                                    {
                                        Response.Write("<div class=\"class_signup_filledStatus\">Non Member Registration Is Not Available.</div>");
                                    }

                                    if (sPriceType == "R")
                                    {
                                        Response.Write("<div class=\"class_signup_filledStatus\">Residency is required.</div>");
                                    }

                                    Response.Write("<div>");
                                    Response.Write("<input type=\"button\" name=\"class_signup_returnButtonClassList\" id=\"class_signup_returnButtonClassList\" value=\"Return to Class List\" onclick=\"goToList()\" />");
                                    Response.Write("</div>");
                                }
                            }
                            else
                            {
                                //Blocked Message
                                sDisplayPublicBlockedNote = classes.displayPublicBlockedNote(iUserID);

                                Response.Write("<div class=\"class_signup_filledStatus\">Your account has been blocked from online purchases.</div>");
                                Response.Write(sDisplayPublicBlockedNote);
                            }

                            break;
                        case 3:  //Open Attendance
                            Response.Write("<div>");
                            Response.Write("  <strong>" + myReader["optionname"].ToString() + " - " + myReader["optiondescription"].ToString() + "</strong>");
                            Response.Write("</div>");
                            break;
                        case 4:  //Information Only
                            goto case 3;
                    }
                }
                else
                {
                    if (sRegStartDate != "")
                    {
                        Response.Write("<div>");
                        Response.Write("  Registration for " + sPriceType);
                        Response.Write("  will start on " + string.Format("{0:MM/dd/yyyy}", Convert.ToDateTime(sRegStartDate)) + ".");
                        Response.Write("  Please try again at that time.");
                        Response.Write("</div>");
                    }
                    else
                    {
                        if (sPriceType == "Non Members")
                        {
                            Response.Write("<div>Non Member registration is not available.</div>");
                        }
                        else
                        {
                            Response.Write("<div>Registration is not available.</div>");
                        }
                    }
                }
            }

            if (sCodeFinalFieldset)
            {
                Response.Write("</fieldset>");
            }
        }
        else
        {
            Response.Write("<div>No information for the class or event could be found.</div>");
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        if (sSetupCoachFields)
        {
            Response.Write("<script type=\"text/javascript\">");
            Response.Write("  setupCoachFields();");
            Response.Write("</script>");
        }
    }
}

/*
------------------------------------------------------------------------------------------------------
 LEFT OFF ON field validation.  Just finished the "checkaddress" function call for the
 "Validate Address" button.  Still need to finish all of the field validation and then submit the form.

 NOTE: As of December 13, 2012...
       Peter stated that I was to NOT include this page in the Menlo Park project and to no longer 
       continue creating it.  I am to have mobile users create an account using the existing 
       registration page (register.asp).  Furthermore, it was decided that because of this that it 
       would be better to create a new "userid" cookie for the ASP.NET environment.
       That new cookie is "useridx".
       
       Peter has decided that the process for setting up a new user is that the user must use the 
       old ASP page (register.asp).  If coming from an ASP.NET page, then  once the account has been 
       created then redirect the user to the ASP.NET login screen (rd_user_login.aspx).
       THAT is the new process for setting up accounts via the ASP.NET environment as of 12/13/2012.
------------------------------------------------------------------------------------------------------
*/
using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_register : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    public void displayUserRegister(Int32 iOrgID)
    {
        Boolean sUsesDisplayName        = false;
        Boolean sIsBusinessNameRequired = false;
        Boolean sHasResidentStreets     = hasResidentTypeStreets(iOrgID, "R");
        Boolean sHasBusinessStreets     = hasResidentTypeStreets(iOrgID, "B");
        Boolean sOrgHasNeighborhoods    = common.orgHasNeighborhoods(iOrgID);

        Boolean sOrgHasFeaturePayments           = common.orgHasFeature(iOrgID.ToString(), "payments");
        Boolean sOrgHasFeatureActionLine         = common.orgHasFeature(iOrgID.ToString(), "action line");
        Boolean sOrgHasFeatureShowGenderPicks    = common.orgHasFeature(iOrgID.ToString(), "display gender pick");
        Boolean sOrgHasFeatureGenderRequired     = common.orgHasFeature(iOrgID.ToString(), "gender required");
        Boolean sOrgHasFeatureDoNotKnock         = common.orgHasFeature(iOrgID.ToString(), "donotknock");
        Boolean sOrgHasFeatureAddressRequired    = common.orgHasFeature(iOrgID.ToString(), "registration req address");
        Boolean sOrgHasFeatureNoEmergencyContact = common.orgHasFeature(iOrgID.ToString(), "no emergency contact");
        Boolean sOrgHasFeatureLargeAddressList   = common.orgHasFeature(iOrgID.ToString(), "large address list");
        Boolean sOrgHasFeatureBidPostingsViewPlanHoldersRequireFields = common.orgHasFeature(iOrgID.ToString(), "bidpostings_viewplanholders_requirefields");

        Boolean sOrgHasDisplayDoNotKnockListDescription = common.orgHasDisplay(iOrgID.ToString(), "donotknock_list_description");

        Int32 sDefaultRelationshipID = getDefaultRelationshipID(iOrgID);
        Int32 sAddressInfoDisplayID  = common.getDisplayID("citizen_register_maint_addressinfo");

        string sOrgName                   = common.getOrgName(iOrgID.ToString());
        string sThisIsARequiredField      = "<span class=\"requiredField\" title=\"This field is required\">*</span>&nbsp;";
        string sGenderPicksRequiredField  = "&nbsp;";
        string sAddressRequiredField      = "&nbsp;";
        string sBusinessNameRequiredField = "&nbsp;";
        string sFeatureNameActionLine     = "";
        string sTransactionMsg            = "You can access your transaction history.";
        string sLabelFirstLastName        = "";
        string sGenderPicks               = "";
        string sUserDefaultCity           = getOrgDefaultInfo(iOrgID, "defaultcity");
        string sUserDefaultState          = getOrgDefaultInfo(iOrgID, "defaultstate");
        string sUserDefaultZip            = getOrgDefaultInfo(iOrgID, "defaultzip");
        string sUserDefaultAreaCode       = getOrgDefaultInfo(iOrgID, "defaultareacode");
        string sOnChangeBusinessName      = "";
        string sDoNotKnockDescription     = "&nbsp;";
        string sDisplayNeighborhoods      = "&nbsp;";
        string sDisplayAddresses          = "&nbsp;";
        string sAddressLabel              = "&nbsp;";
        string sAddressInfo               = common.getOrgDisplayWithID(iOrgID,
                                                                       sAddressInfoDisplayID,
                                                                       sUsesDisplayName);
        
        //Setup Transaction Message
        if (sOrgHasFeaturePayments || sOrgHasFeatureActionLine)
        {
            //sTransactionMsg += "<div>";
            sTransactionMsg += "  For example,";

            if (sOrgHasFeaturePayments)
            {
                sTransactionMsg += " history of online payments using " + sOrgName + " E-Gov Services";
            }

            if (sOrgHasFeatureActionLine)
            {
                if (sOrgHasFeaturePayments)
                {
                    sTransactionMsg += " or";
                }

                sFeatureNameActionLine = common.getFeatureName(iOrgID.ToString(), "action line");

                sTransactionMsg += " requests submitted via the " + sFeatureNameActionLine;
            }

            //sTransactionMsg += "</div>";
        }

        if (Request["fromPostings"] == "Y" && sOrgHasFeatureBidPostingsViewPlanHoldersRequireFields)
        {
            sIsBusinessNameRequired = true;
            sLabelFirstLastName     = "Contact ";
        }

        Response.Write("<div id=\"registerInfo\">");
        Response.Write("  <div id=\"registerWelcomeMsg\">Welcome to the " + sOrgName + " Registration</div>");
        Response.Write("  <div id=\"registerRegistrationMsg\">Registering to use " + sOrgName + " E-Gov Services is FREE, quick, and easy to establish!</div>");
        Response.Write("  <div id=\"registerTransactionMsg\">" + sTransactionMsg + "</div>");
        Response.Write("  <div>");
        Response.Write("    You can choose to have contact information, such as address and telephone number, saved with your ");
        Response.Write("    membership thereby eliminating the requirement to re-type this information into online forms.");
        Response.Write("  </div>");
        Response.Write("</div>");
        Response.Write("<div id=\"registerFields\">");
        Response.Write("  <fieldset class=\"register_fieldset\">");
        Response.Write("    <legend>" + sOrgName + " Registration</legend>");
        Response.Write("    <form name=\"register\" id=\"register\" method=\"post\" action=\"rd_register.aspx\">");
        Response.Write("      <input type=\"hidden\" name=\"isFinalCheck\" id=\"isFinalCheck\" value=\"N\" size=\"1\" maxlength=\"1\" />");
        Response.Write("      <input type=\"hidden\" name=\"columnnameid\" id=\"columnnameid\" value=\"userid\" />");
        Response.Write("      <input type=\"hidden\" name=\"egov_users_userregistered\" id=\"egov_users_userregistered\" value=\"1\" />");
        Response.Write("      <input type=\"hidden\" name=\"egov_users_orgid\" id=\"egov_users_orgid\" value=\"" + iOrgID.ToString() + "\" />");
        Response.Write("      <input type=\"hidden\" name=\"egov_users_relationshipid\" id=\"egov_users_relationshipid\" value=\"" + sDefaultRelationshipID + "\" />");
        Response.Write("      <input type=\"hidden\" name=\"ef:egov_users_useremail-text/req\" id=\"ef:egov_users_useremail-text/req\" value=\"Email Address\" />");
        Response.Write("      <input type=\"hidden\" name=\"ef:egov_users_userpassword-text/req\" id=\"ef:egov_users_userpassword-text/req\" value=\"Password 1\" />");
        Response.Write("      <input type=\"hidden\" name=\"ef:skip_userpassword2-text/req\" id=\"ef:skip_userpassword2-text/req\" value=\"Password 2\" />");
        Response.Write("      <input type=\"hidden\" name=\"ef:egov_users_userhomephone-text/req/phone\" id=\"ef:egov_users_userhomephone-text/req/phone\" value=\"Phone Number\" />");
        Response.Write("      <input type=\"hidden\" name=\"ef:egov_users_userfname-text/req\" id=\"ef:egov_users_userfname-text/req\" value=\"First name\" />");
        Response.Write("      <input type=\"hidden\" name=\"ef:egov_users_userlname-text/req\" id=\"ef:egov_users_userlname-text/req\" value=\"Last name\" />");
        Response.Write("      <input type=\"hidden\" name=\"egov_users_residenttype\" id=\"egov_users_residenttype\" value=\"N\" />");
        Response.Write("      <input type=\"hidden\" name=\"egov_users_neighborhoodid\" id=\"egov_users_neighborhoodid\" value=\"0\" />");
        Response.Write("      <input type=\"hidden\" name=\"egov_users_headofhousehold\" id=\"egov_users_headofhousehold\" value=\"1\" />");

        if (!sOrgHasFeatureShowGenderPicks)
        {
            Response.Write("      <input type=\"hidden\" name=\"egov_users_gender\" id=\"egov_users_gender\" value=\"N\" />");
        }

        if (!sOrgHasFeatureDoNotKnock)
        {
            Response.Write("      <input type=\"hidden\" name=\"isOnDoNotKnockList_peddlers\" id=\"isOnDoNotKnockList_peddlers\" value=\"\" />");
            Response.Write("      <input type=\"hidden\" name=\"isOnDoNotKnockList_solicitors\" id=\"isOnDoNotKnockList_solicitors\" value=\"\" />");
        }

        if (!sOrgHasNeighborhoods)
        {
            Response.Write("      <input type=\"hidden\" name=\"skip_neighborhood\" id=\"skip_neighborhood\" value=\"0\" />");
        }

        Response.Write("    <a name=\"errorMsg\"><div id=\"registerErrorMsgDiv\"></div></a>");
        //Response.Write("    <div id=\"screenMsg\"></div>");
        Response.Write("    <div id=\"requiredFieldMsg\"><span class=\"requiredField\">*</span> Indicates a required field that must have a value in order to complete your registration.</div>");
        Response.Write("    <table id=\"registrationTable\" border=\"0\">");

        //Email
        Response.Write("      <tr>");
        Response.Write("          <td class=\"registerLabel\">" + sThisIsARequiredField + "Email:</td>");
        Response.Write("          <td>");
        Response.Write("              <input type=\"text\" name=\"egov_users_useremail\" id=\"egov_users_useremail\" value=\"" + Request["egov_users_useremail"] + "\" class=\"registerInputFieldLarge\" onchange=\"clearMsg('egov_users_useremail');\" />");
        Response.Write("          </td>");
        Response.Write("      </tr>");

        //Password
        Response.Write("      <tr>");
        Response.Write("          <td class=\"registerLabel\">" + sThisIsARequiredField + "Password:</td>");
        Response.Write("          <td>");
        Response.Write("              <input type=\"password\" name=\"egov_users_userpassword\" id=\"egov_users_userpassword\" value=\"" + Request["egov_users_userpassword"] + "\" class=\"registerInputFieldLarge\" onchange=\"clearMsg('egov_users_userpassword');\" />");
        Response.Write("          </td>");
        Response.Write("      </tr>");

        //Verify Password
        Response.Write("      <tr>");
        Response.Write("          <td class=\"registerLabel\">" + sThisIsARequiredField + "Verify Password:</td>");
        Response.Write("          <td>");
        Response.Write("              <input type=\"password\" name=\"skip_userpassword2\" id=\"skip_userpassword2\" value=\"" + Request["skip_userpassword2"] + "\" class=\"registerInputFieldLarge\" onchange=\"clearMsg('skip_userpassword2');\" />");
        Response.Write("          </td>");
        Response.Write("      </tr>");

        //First Name
        Response.Write("      <tr>");
        Response.Write("          <td class=\"registerLabel\">" + sThisIsARequiredField + sLabelFirstLastName + "First Name:</td>");
        Response.Write("          <td>");
        Response.Write("              <input type=\"text\" name=\"egov_users_userfname\" id=\"egov_users_userfname\" value=\"" + Request["egov_users_userfname"] + "\" class=\"registerInputFieldLarge\" onchange=\"clearMsg('egov_users_userfname');\" />");
        Response.Write("          </td>");
        Response.Write("      </tr>");

        //Last Name
        Response.Write("      <tr>");
        Response.Write("          <td class=\"registerLabel\">" + sThisIsARequiredField + sLabelFirstLastName + "Last Name:</td>");
        Response.Write("          <td>");
        Response.Write("              <input type=\"text\" name=\"egov_users_userlname\" id=\"egov_users_userlname\" value=\"" + Request["egov_users_userfname"] + "\" class=\"registerInputFieldLarge\" onchange=\"clearMsg('egov_users_userlname');\" />");
        Response.Write("          </td>");
        Response.Write("      </tr>");

        //Gender Picks
        if (sOrgHasFeatureShowGenderPicks)
        {
            sGenderPicks = common.displayGenderPicks("egov_users_gender", "N");

            if (sOrgHasFeatureGenderRequired)
            {
                sGenderPicksRequiredField = sThisIsARequiredField;
            }

            Response.Write("      <tr>");
            Response.Write("          <td class=\"registerLabel\">" + sGenderPicksRequiredField + "Gender:</td>");
            Response.Write("          <td>" + sGenderPicks + "</td>");
            Response.Write("      </tr>");
        }

        //Phone Number
        Response.Write("      <tr>");
        Response.Write("          <td class=\"registerLabel\">" + sThisIsARequiredField + "Phone Number:</td>");
        Response.Write("          <td>");
        Response.Write("              <input type=\"hidden\" name=\"egov_users_userhomephone\" id=\"egov_users_userhomephone\" value=\"\" />");
        Response.Write("             (<input type=\"text\" value=\"\" size=\"3\" maxlength=\"3\" name=\"skip_user_areacode\" id=\"skip_user_areacode\" onkeyup=\"return autoTab(this, 3, event);\" onchange=\"clearMsg('skip_user_areacode');\" />)");
        Response.Write("              <input type=\"text\" value=\"\" size=\"3\" maxlength=\"3\" name=\"skip_user_exchange\" id=\"skip_user_exchange\" onkeyup=\"return autoTab(this, 3, event);\" onchange=\"clearMsg('skip_user_exchange');\" /> &dash;");
        Response.Write("              <input type=\"text\" value=\"\" size=\"4\" maxlength=\"4\" name=\"skip_user_line\" id=\"skip_user_line\" onkeyup=\"return autoTab(this, 4, event);\" onchange=\"clearMsg('skip_user_line');\" />");
        Response.Write("          </td>");
        Response.Write("      </tr>");

        //Mobile Phone
        Response.Write("      <tr>");
        Response.Write("          <td class=\"registerLabel\">" + sThisIsARequiredField + "Mobile Phone:</td>");
        Response.Write("          <td>");
        Response.Write("              <input type=\"hidden\" name=\"egov_users_usercell\" id=\"egov_users_userhomecell\" value=\"\" />");
        Response.Write("             (<input type=\"text\" value=\"\" size=\"3\" maxlength=\"3\" name=\"skip_cell_areacode\" id=\"skip_cell_areacode\" onkeyup=\"return autoTab(this, 3, event);\" onchange=\"clearMsg('skip_cell_areacode');\" />)");
        Response.Write("              <input type=\"text\" value=\"\" size=\"3\" maxlength=\"3\" name=\"skip_cell_exchange\" id=\"skip_cell_exchange\" onkeyup=\"return autoTab(this, 3, event);\" onchange=\"clearMsg('skip_cell_exchange');\" /> &dash;");
        Response.Write("              <input type=\"text\" value=\"\" size=\"4\" maxlength=\"4\" name=\"skip_cell_line\" id=\"skip_cell_line\" onkeyup=\"return autoTab(this, 4, event);\" onchange=\"clearMsg('skip_cell_line');\" />");
        Response.Write("          </td>");
        Response.Write("      </tr>");

        //Fax
        Response.Write("      <tr>");
        Response.Write("          <td class=\"registerLabel\">" + sThisIsARequiredField + "Fax:</td>");
        Response.Write("          <td>");
        Response.Write("              <input type=\"hidden\" name=\"egov_users_userfax\" id=\"egov_users_userfax\" value=\"\" />");
        Response.Write("             (<input type=\"text\" value=\"\" size=\"3\" maxlength=\"3\" name=\"skip_fax_areacode\" id=\"skip_fax_areacode\" onkeyup=\"return autoTab(this, 3, event);\" onchange=\"clearMsg('skip_fax_areacode');\" />)");
        Response.Write("              <input type=\"text\" value=\"\" size=\"3\" maxlength=\"3\" name=\"skip_fax_exchange\" id=\"skip_fax_exchange\" onkeyup=\"return autoTab(this, 3, event);\" onchange=\"clearMsg('skip_fax_exchange');\" /> &dash;");
        Response.Write("              <input type=\"text\" value=\"\" size=\"4\" maxlength=\"4\" name=\"skip_fax_line\" id=\"skip_fax_line\" onkeyup=\"return autoTab(this, 4, event);\" onchange=\"clearMsg('skip_fax_line');\" />");
        Response.Write("          </td>");
        Response.Write("      </tr>");

        //Show additional address info if org has "edit display"
        if (sAddressInfo != "")
        {
            Response.Write("      <tr id=\"registerAddressInfo\">");
            Response.Write("          <td>&nbsp;</td>");
            Response.Write("          <td>" + sAddressInfo + "</td>");
            Response.Write("      </tr>");
        }

        //Resident Street
        if (sHasResidentStreets)
        {
            if (!sOrgHasFeatureLargeAddressList)
            {
                sAddressLabel     = "Resident Street:";
                sDisplayAddresses = displayAddresses(iOrgID, "R");
            }
            else
            {
                if (sOrgHasFeatureAddressRequired)
                {
                    sAddressRequiredField = "<span class=\"requiredField\" title=\"This field is required\">*</span>&nbsp;";
                }

                sAddressLabel     = "Resident Address:";
                sDisplayAddresses = displayLargeAddressList(iOrgID, "R");
            }

            Response.Write("      <tr valign=\"top\">");
            Response.Write("          <td class=\"registerLabel\">" + sAddressRequiredField + sAddressLabel + "</td>");
            Response.Write("          <td class=\"registerAddresses\" id=\"residentLargeAddressFields\">");
            Response.Write("              <a name=\"addressError\"></a><div id=\"registerAddressError\"></div>" + sDisplayAddresses + "</td>");
            Response.Write("      </tr>");
        }

        //The result of THIS check will be used with the following fields:
        //  Address - Not Listed, City, State, and Zip
        if (sOrgHasFeatureAddressRequired)
        {
            sAddressRequiredField = "<span class=\"requiredField\" title=\"This field is required\">*</span>&nbsp;";
        }

        //Address - Not Listed
        sAddressLabel     = "Address";
        sDisplayAddresses = "<input type=\"text\" name=\"egov_users_useraddress\" id=\"egov_users_useraddress\" value=\"\" maxlength=\"100\" class=\"registerInputFieldLarge\" onchange=\"cleanUpAddressFields('egov_users_useraddress');\" />";

        if (sHasResidentStreets)
        {
            sAddressLabel += " (if not listed)";
        }

        sAddressLabel += ":";

        Response.Write("      <tr valign=\"top\">");
        Response.Write("          <td class=\"registerLabel\">" + sAddressRequiredField + sAddressLabel + "</td>");
        Response.Write("          <td class=\"registerAddresses\">" + sDisplayAddresses + "</td>");
        Response.Write("      </tr>");

        //Resident Unit
        Response.Write("      <tr>");
        Response.Write("          <td class=\"registerLabel\">Resident Unit:</td>");
        Response.Write("          <td>");
        Response.Write("              <input type=\"text\" name=\"egov_users_userunit\" id=\"egov_users_userunit\" value=\"\" size=\"10\" maxlength=\"10\" />");
        Response.Write("          </td>");
        Response.Write("      </tr>");

        //Neighborhood
        if (sOrgHasNeighborhoods)
        {
            sDisplayNeighborhoods = displayNeighborhoods(iOrgID);

            Response.Write("      <tr>");
            Response.Write("          <td class=\"registerLabel\">Neighborhood:</td>");
            Response.Write("          <td>" + sDisplayNeighborhoods + "</td>");
            Response.Write("      </tr>");
        }

        //City
        Response.Write("      <tr>");
        Response.Write("          <td class=\"registerLabel\">" + sAddressRequiredField + "City:</td>");
        Response.Write("          <td>");
        Response.Write("              <input type=\"text\" name=\"egov_users_usercity\" id=\"egov_users_usercity\" value=\"" + sUserDefaultCity + "\" size=\"20\" maxlength=\"40\" />");
        Response.Write("          </td>");
        Response.Write("      </tr>");

        //State
        Response.Write("      <tr>");
        Response.Write("          <td class=\"registerLabel\">" + sAddressRequiredField + "State:</td>");
        Response.Write("          <td>");
        Response.Write("              <input type=\"text\" name=\"egov_users_userstate\" id=\"egov_users_userstate\" value=\"" + sUserDefaultState + "\" size=\"3\" maxlength=\"2\" />");
        Response.Write("          </td>");
        Response.Write("      </tr>");

        //Zip
        Response.Write("      <tr>");
        Response.Write("          <td class=\"registerLabel\">" + sAddressRequiredField + "Zip:</td>");
        Response.Write("          <td>");
        Response.Write("              <input type=\"text\" name=\"egov_users_userzip\" id=\"egov_users_userzip\" value=\"" + sUserDefaultZip + "\" size=\"10\" maxlength=\"10\" />");
        Response.Write("          </td>");
        Response.Write("      </tr>");

        //Business Name
        if (sIsBusinessNameRequired)
        {
            sOnChangeBusinessName      = " onchange=\"clearMsg('egov_users_userbusinessname')\"";
            sBusinessNameRequiredField = "<span class=\"requiredField\" title=\"This field is required\">*</span>&nbsp;";
        }
                
        Response.Write("      <tr>");
        Response.Write("          <td class=\"registerLabel\">" + sBusinessNameRequiredField + "Business Name:</td>");
        Response.Write("          <td>");
        Response.Write("              <input type=\"text\" name=\"egov_users_userbusinessname\" id=\"egov_users_userbusinessname\" value=\"" + Request["egov_users_businessname"] + "\" maxlength=\"100\" class=\"registerInputFieldLarge\"" + sOnChangeBusinessName + " />");
        Response.Write("          </td>");
        Response.Write("      </tr>");

        //Business Street
        if (sHasBusinessStreets)
        {
            sDisplayAddresses = displayAddresses(iOrgID, "B");

            Response.Write("      <tr>");
            Response.Write("          <td class=\"registerLabel\">Business Street:</td>");
            Response.Write("          <td>" + sDisplayAddresses + "</td>");
            Response.Write("      </tr>");
        }

        //Business Street - Not Listed
        sAddressLabel     = "Business Street:";
        sDisplayAddresses = "<input type=\"text\" name=\"egov_users_userbusinessaddress\" id=\"egov_users_userbusinessaddress\" value=\"" + Request["egov_users_userbusinessaddress"] + "\" maxlength=\"100\" class=\"registerInputFieldLarge\" />";

        if (sHasBusinessStreets)
        {
            sAddressLabel = "Street (if not listed):";
        }

        Response.Write("      <tr valign=\"top\">");
        Response.Write("          <td class=\"registerLabel\">" + sAddressLabel + "</td>");
        Response.Write("          <td class=\"registerAddresses\">" + sDisplayAddresses + "</td>");
        Response.Write("      </tr>");

        //Work Phone
        Response.Write("      <tr>");
        Response.Write("          <td class=\"registerLabel\">Work Phone:</td>");
        Response.Write("          <td>");
        Response.Write("              <input type=\"hidden\" name=\"egov_users_userworkphone\" id=\"egov_users_userfax\" value=\"\" />");
        Response.Write("             (<input type=\"text\" value=\"\" size=\"3\" maxlength=\"3\" name=\"skip_work_areacode\" id=\"skip_work_areacode\" onkeyup=\"return autoTab(this, 3, event);\" onchange=\"clearMsg('skip_work_areacode');\" />)");
        Response.Write("              <input type=\"text\" value=\"\" size=\"3\" maxlength=\"3\" name=\"skip_work_exchange\" id=\"skip_work_exchange\" onkeyup=\"return autoTab(this, 3, event);\" onchange=\"clearMsg('skip_work_exchange');\" /> &dash;");
        Response.Write("              <input type=\"text\" value=\"\" size=\"4\" maxlength=\"4\" name=\"skip_work_line\" id=\"skip_work_line\" onkeyup=\"return autoTab(this, 4, event);\" onchange=\"clearMsg('skip_work_line');\" />");
        Response.Write("         ext. <input type=\"text\" value=\"\" size=\"4\" maxlength=\"4\" name=\"skip_work_ext\" id=\"skip_work_ext\" onkeyup=\"return autoTab(this, 4, event);\" />");
        Response.Write("          </td>");
        Response.Write("      </tr>");

        if (!sOrgHasFeatureNoEmergencyContact)
        {
            //Emergency Contact
            Response.Write("      <tr>");
            Response.Write("          <td class=\"registerLabel\">Emergency Contact:</td>");
            Response.Write("          <td>");
            Response.Write("              <input type=\"text\" name=\"egov_users_emergencycontact\" id=\"egov_users_emergencycontact\" value=\"" + Request["egov_users_emergencycontact"] + "\" maxlength=\"100\" class=\"registerInputFieldLarge\" />");
            Response.Write("          </td>");
            Response.Write("      </tr>");

            //Emergency Phone
            Response.Write("      <tr>");
            Response.Write("          <td class=\"registerLabel\">Emergency Phone:</td>");
            Response.Write("          <td>");
            Response.Write("              <input type=\"hidden\" name=\"egov_users_emergencyphone\" id=\"egov_users_emergencyphone\" value=\"\" />");
            Response.Write("             (<input type=\"text\" value=\"\" size=\"3\" maxlength=\"3\" name=\"skip_emergencyphone_areacode\" id=\"skip_emergencyphone_areacode\" onkeyup=\"return autoTab(this, 3, event);\" onchange=\"clearMsg('skip_emergencyphone_areacode');\" />)");
            Response.Write("              <input type=\"text\" value=\"\" size=\"3\" maxlength=\"3\" name=\"skip_emergencyphone_exchange\" id=\"skip_emergencyphone_exchange\" onkeyup=\"return autoTab(this, 3, event);\" onchange=\"clearMsg('skip_emergencyphone_exchange');\" /> &dash;");
            Response.Write("              <input type=\"text\" value=\"\" size=\"4\" maxlength=\"4\" name=\"skip_emergencyphone_line\" id=\"skip_emergencyphone_line\" onkeyup=\"return autoTab(this, 4, event);\" onchange=\"clearMsg('skip_emergencyphone_line');\" />");
            Response.Write("          </td>");
            Response.Write("      </tr>");
        }

        if (sOrgHasFeatureDoNotKnock)
        {
            if (sOrgHasDisplayDoNotKnockListDescription)
            {
                sDoNotKnockDescription = common.getOrgDisplay(iOrgID.ToString(), "donotknock_list_description");

                sDoNotKnockDescription = "<tr><td>" + sDoNotKnockDescription + "</td></tr>";
            }

            sDoNotKnockDescription = "blahblahblah";

            Response.Write("      <tr>");
            Response.Write("          <td colspan=\"2\">");
            Response.Write("             <fieldset class=\"fieldset\">");
            Response.Write("               <legend>\"Do Not Knock\" List(s)</legend>");
            //Response.Write("               <div id=\"registerDNKVendorTitle\">" + sDNKVendorTitle + "</div>");
            Response.Write("               <table>");
            Response.Write(sDoNotKnockDescription);
            Response.Write("                 <tr>");
            Response.Write("                     <td>");
            Response.Write("                         <input type=\"checkbox\" name=\"isOnDoNotKnockList_peddlers\" id=\"isOnDoNotKnockList_peddlers\" value=\"on\" />&nbsp;Do Not Knock - Peddlers<br />");
            Response.Write("                         <input type=\"checkbox\" name=\"isOnDoNotKnockList_solicitors\" id\"isOnDoNotKnockList_solicitors\" value\"on\" />&nbsp;Do Not Knock - Solicitors");
            Response.Write("                     </td>");
            Response.Write("                 </tr>");
            Response.Write("               </table>");
            Response.Write("             </fieldset>");
            Response.Write("          </td>");
            Response.Write("      </tr>");
        }

        //Subscription Mailing Lists <<--- GO HERE.
        //Didn't implement them because this registration is ONLY for classes in .NET.

        //Submit Registration Button
        Response.Write("          <tr>");
        Response.Write("              <td colspan=\"2\" id=\"registerButtonDiv\">");
        Response.Write("                  <input type=\"button\" name=\"registerSubmitButton\" id=\"registerSubmitButton\" value=\"Submit Registration Form\" />");
        Response.Write("              </td>");
        Response.Write("          </tr>");
        Response.Write("    </table>");

        //Problem Field
        Response.Write("    <div id=\"registerProblemField\">");
        Response.Write("      Internal Use Only, Leave Blank: ");
        Response.Write("      <input type=\"text\" name=\"subjecttext\" id=\"problemtextinput\" value=\"\" size=\"6\" />");
        Response.Write("      <input type=\"hidden\" name=\"problemorg\" id=\"problemorg\" value=\"" + iOrgID.ToString() + "\" /><br />");
        Response.Write("      <strong>Please leave this field blank and remove any values that have been populated for it.</strong>");
        Response.Write("    </div>");

        Response.Write("    </form>");
        Response.Write("  </fieldset>");
        Response.Write("</div>");
    }

    public static Int32 getDefaultRelationshipID(Int32 iOrgID)
    {
        Int32 lcl_return = 0;

        string sSQL = "";

        sSQL  = "SELECT relationshipid ";
        sSQL += " FROM egov_familymember_relationships ";
        sSQL += " WHERE orgid = " + iOrgID.ToString();
        sSQL += " AND isdefault = 1 ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToInt32(myReader["relationshipid"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getOrgDefaultInfo(Int32 iOrgID,
                                           string iDBColumn)
    {
        string lcl_return = "";
        string sSQL       = "";

        sSQL  = "SELECT " + iDBColumn + " as defaultColumn";
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

            lcl_return = Convert.ToString(myReader["defaultColumn"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Boolean hasResidentTypeStreets(Int32 iOrgID,
                                                 string iResidentType)
    {
        Boolean lcl_return = false;

        string sSQL          = "";
        string sResidentType = "''";

        if (iResidentType != "")
        {
            sResidentType = common.dbSafe(iResidentType);
            sResidentType = "'" + sResidentType + "'";
        }

        sSQL  = "SELECT count(residentaddressid) as hits ";
        sSQL += " FROM egov_residentaddresses ";
        sSQL += " WHERE orgid = " + iOrgID.ToString();
        sSQL += " AND residenttype = " + sResidentType;

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

    public static string displayAddresses(Int32 iOrgID,
                                          string iResidentType)
    {
        string lcl_return           = "";
        string sSQL                 = "";
        string sResidentType        = "";
        string sResidentTypeFieldID = "";
        string sOptionValue         = "";

        if (iResidentType != "")
        {
            sResidentType = common.dbSafe(iResidentType);

            sResidentTypeFieldID = sResidentType;
            sResidentType        = "'" + sResidentType + "'";
        }

        sSQL  = "SELECT residentstreetnumber, ";
        sSQL += " residentstreetname ";
        sSQL += " FROM egov_residentaddresses_list ";
        sSQL += " WHERE orgid = " + iOrgID.ToString();
        sSQL += " AND residenttype = " + sResidentType;
        sSQL += " ORDER BY sortstreetname, ";
        sSQL +=          " residentstreetprefix, ";
        sSQL +=          " CAST(residentstreetnumber as int) ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            lcl_return = "<select name=\"skip_" + sResidentTypeFieldID + "address\" id=\"skip_" + sResidentTypeFieldID + "address\">";
            lcl_return += "  <option value=\"0000\">Please select an address...</option>";

            while (myReader.Read())
            {
                sOptionValue = Convert.ToString(myReader["residentstreetnumber"]) + " " + Convert.ToString(myReader["residentstreetname"]);
                sOptionValue = sOptionValue.Trim();

                lcl_return += "  <option value=\"" + sOptionValue + "\">" + sOptionValue + "</option>";
            }

            lcl_return += "</select>";
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string displayLargeAddressList(Int32 iOrgID,
                                                 string iResidentType)
    {
        Boolean sOrgHasFeatureCitizenRegistrationNoValidateAddress = common.orgHasFeature(iOrgID.ToString(), "citizenregistration_novalidate_address");

        string lcl_return     = "";
        string sSQL           = "";
        string sResidentType  = "''";
        string sOptionValue   = "";

        if (iResidentType != "")
        {
            sResidentType = common.dbSafe(iResidentType);
            sResidentType = "'" + sResidentType + "'";
        }

        sSQL  = "SELECT distinct sortstreetname, ";
        sSQL += " isnull(residentstreetprefix, '') as residentstreetprefix, ";
        sSQL += " residentstreetname, ";
        sSQL += " isnull(streetsuffix, '') as streetsuffix, ";
        sSQL += " isnull(streetdirection, '') as streetdirection ";
        sSQL += " FROM egov_residentaddresses ";
        sSQL += " WHERE orgid = " + iOrgID.ToString();
        sSQL += " AND residenttype = " + sResidentType;
        sSQL += " AND residentstreetname IS NOT NULL ";
        sSQL += " ORDER BY sortstreetname ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            lcl_return = "<input type=\"text\" name=\"residentstreetnumber\" id=\"residentstreetnumber\" value=\"\" size=\"8\" maxlength=\"10\" onchange=\"cleanUpAddressFields('residentstreetnumber');\" />&nbsp;";
            lcl_return += "<select name=\"streetaddress\" id=\"streetaddress\" onchange=\"cleanUpAddressFields('streetaddress');\">";
            lcl_return += "  <option value=\"0000\">Choose street from dropdown</option>";

            while (myReader.Read())
            {
                sOptionValue = "";

                //Resident Street Prefix
                if (Convert.ToString(myReader["residentstreetprefix"]) != "")
                {
                    sOptionValue += Convert.ToString(myReader["residentstreetprefix"]);
                    sOptionValue += " ";
                }

                //Resident Street Name
                sOptionValue += Convert.ToString(myReader["residentstreetname"]);

                //Resident Street Suffix
                if (Convert.ToString(myReader["streetsuffix"]) != "")
                {
                    sOptionValue += " ";
                    sOptionValue += Convert.ToString(myReader["streetsuffix"]);
                }

                //Resident Street Direction
                if (Convert.ToString(myReader["streetdirection"]) != "")
                {
                    sOptionValue += " ";
                    sOptionValue += Convert.ToString(myReader["streetdirection"]);
                }

                lcl_return += "  <option value=\"" + sOptionValue + "\">" + sOptionValue + "</option>";
            }

            lcl_return += "</select>";

            if (!sOrgHasFeatureCitizenRegistrationNoValidateAddress)
            {
                lcl_return += "&nbsp;";
                lcl_return += "<input type=\"button\" name=\"registerValidateAddressButton\" id=\"registerValidateAddressButton\" value=\"Validate Address\" onclick=\"checkAddress('CheckResults', 'no');\" />";
            }

            lcl_return += "<fieldset id=\"validAddressList\" class=\"fieldset\">";
            lcl_return += "  <legend>Invalid Address</legend>";
            lcl_return += "  <p>The address you entered does not match any in the system. ";
            lcl_return += "  You can select a valid address from the list, or if you are ";
            lcl_return += "  certain the address you entered is correct, click the \"Use the ";
            lcl_return += "  address I entered\" button to continue.</p>";
            lcl_return += "  <strong>The address you entered: </strong><span id=\"registerDisplayAddressEntered\"></span><br />";
            lcl_return += "  <input type=\"text\" value=\"\" name=\"oldstnumber\" id=\"oldstnumber\" size=\"8\" maxlength=\"10\" />";
            lcl_return += "  <input type=\"text\" value=\"\" name=\"stname\" id=\"stname\" size=\"50\" maxlength=\"50\" />";
            lcl_return += "  <div id=\"addresspicklist\"></div>";
            lcl_return += "  <input type=\"button\" name=\"registerValidPickButton\" id=\"registerValidPickButton\" value=\"Use the valid address selected\" onclick=\"doSelect();\" />";
            lcl_return += "  <input type=\"button\" name=\"registerInvalidPickButton\" id=\"registerInvalidPickButton\" value=\"Use the address I entered\" onclick=\"doKeep();\" />";
            lcl_return += "  <input type=\"button\" name=\"registerCancelPickButton\" id=\"registerCancelPickButton\" value=\"Cancel\" onclick=\"cancelPick();\" />";
            lcl_return += "</fieldset>";
        }
        else
        {
            lcl_return = "<input type=\"hidden\" name=\"residentstreetnumber\" id=\"residentstreetnumber\" value=\"\" size=\"8\" maxlength=\"10\" />&nbsp;";
            lcl_return += "<input type=\"hidden\" name=\"streetaddress\" id=\"streetaddress\" value=\"0000\" />";
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string displayNeighborhoods(Int32 iOrgID)
    {
        string lcl_return = "";
        string sSQL       = "";

        sSQL  = "SELECT neighborhoodid, ";
        sSQL += " isnull(neighborhood, '') as neighborhood ";
        sSQL += " FROM egov_neighborhoods ";
        sSQL += " WHERE orgid = " + iOrgID.ToString();
        sSQL += " ORDER BY neighborhood ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            lcl_return = "<select name=\"skip_neighborhoodid\" id=\"skip_neighborhoodid\">";
            lcl_return += "  <option value=\"0\">Not on List...</option>";

            while (myReader.Read())
            {
                lcl_return += "  <option value=\"" + Convert.ToString(myReader["neighborhoodid"]) +"\">" + Convert.ToString(myReader["neighborhoodid"]) + "</option>";
            }

            lcl_return += "</select>";
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }
}

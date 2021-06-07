using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_classes_class_details : System.Web.UI.Page
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
        common.logThePageVisit(startCounter, "class_details.aspx", "public");
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        string sPublicURL = "";

        sPublicURL = common.getFeaturePublicURL(sOrgID, "activities");

        if (sPublicURL != "rd_classes/class_categories.aspx")
        {
            Response.Redirect("../" + sPublicURL);
        }
    }

    public void displayClassInfo(Int32 iOrgID,
                                 Int32 iClassID,
                                 Int32 iCategoryID)
    {
        string mySessionID = Session.SessionID;

        Boolean lcl_orghasfeature_gender_restriction = common.orgHasFeature(iOrgID.ToString(), "gender restriction");
        Boolean sRegistrationEnded              = false;
        Boolean sDisplayDirections              = true;
        Boolean sCanPurchase                    = true;
        Boolean sIsClassFilled                  = false;
        Boolean sPublicCanOnlyView              = true;
        Boolean sClassFeesIncludeContainerDIV   = true;
        Boolean sClassFeesIncludeContainerTABLE = true;

        DateTime sEarlistRegistrationStart = classes.getEarliestRegistrationStart(iClassID);

        Int32 sTimeID = 0;
        double sMinage = 0;
        double sMaxage = 0;
        Int32 sGenderNotRequiredID = 1;
        Int32 sLocationID = 0;

        string sSQL                       = "";
        string sDateLine                  = "";
        string sClassFeeLine              = "";
        string sRegistrationLine          = "";
        string sBreadCrumbsLine           = "";
        string sBreadCrumbsReturnLocation = "CLASSDETAILS";

        string sClassName              = "";
        string sClassDescription       = "";
        string sEarlyClassLabel        = "";
        string sRegistrationDate_Start = "";
        string sRegistrationDate_End   = "";
        string sGenderRestriction      = "";
        string sPhoneNumber            = "";
        string sPOCName                = "";
        string sPOCEmail               = "";
        string sDisplayMinage          = "&nbsp;";
        string sDisplayMaxage          = "&nbsp;";
        string sImgURL                 = "";
        string sImgALT                 = "";
        string sExternalURL            = "";
        string sExternalLinkText       = "";

        //SqlString utf8EncodedString;
        
        sSQL = "SELECT C.classid, ";
        sSQL += " C.classname, ";
        sSQL += " C.classdescription, ";
        sSQL += " C.orgid, ";
        sSQL += " C.classformid, ";
        sSQL += " C.parentclassid, ";
        sSQL += " C.isparent, ";
        sSQL += " C.statusid, ";
        sSQL += " C.cancelreason, ";
        sSQL += " C.imgalttag, ";
        sSQL += " C.publishstartdate, ";
        sSQL += " C.publishenddate, ";
        sSQL += " C.promotiondate, ";
        sSQL += " C.evaluationdate, ";
        sSQL += " C.alternatedate, ";
        sSQL += " C.agecomparedate, ";
        sSQL += " C.minage, ";
        sSQL += " C.maxage, ";
        sSQL += " C.genderrestrictionid, ";
        sSQL += " C.locationid, ";
        sSQL += " C.pocid, ";
        sSQL += " C.evaluationemailbody, ";
        sSQL += " C.searchkeywords, ";
        sSQL += " C.externalurl, ";
        sSQL += " C.externallinktext, ";
        sSQL += " C.canceldate, ";
        sSQL += " C.noenddate, ";
        sSQL += " C.classtypeid, ";
        sSQL += " C.optionid, ";
        sSQL += " C.sequenceid, ";
        sSQL += " C.ispublishable, ";
        sSQL += " C.promotionmsg, ";
        sSQL += " C.membershipid, ";
        sSQL += " C.pricediscountid, ";
        sSQL += " C.itemtypeid, ";
        sSQL += " C.minageprecisionid, ";
        sSQL += " C.maxageprecisionid, ";
        sSQL += " C.supervisorid, ";
        sSQL += " C.mingrade, ";
        sSQL += " C.maxgrade, ";
        sSQL += " C.notes, ";
        sSQL += " C.activitynumber, ";
        sSQL += " C.classseasonid, ";
        sSQL += " C.allowearlyregistration, ";
        sSQL += " C.earlyregistrationdate, ";
        sSQL += " C.earlyregistrationclassseasonid, ";
        sSQL += " C.earlyregistrationclassid, ";
        sSQL += " C.created, ";
        sSQL += " C.displayrosterpublic, ";
        sSQL += " C.teamreg_tshirt_enabled, ";
        sSQL += " C.teamreg_pants_enabled, ";
        sSQL += " C.teamreg_grade_enabled, ";
        sSQL += " C.teamreg_coach_enabled, ";
        sSQL += " C.teamreg_tshirt_inputtype, ";
        sSQL += " C.teamreg_pants_inputtype, ";
        sSQL += " C.teamreg_grade_inputtype, ";
        sSQL += " C.teamreg_tshirt_enabled_original, ";
        sSQL += " C.teamreg_pants_enabled_original, ";
        sSQL += " C.isregatta, ";
        sSQL += " C.regattasignuptypeid, ";
        sSQL += " C.showTerms, ";
        sSQL += " C.publiccanonlyview, ";
        sSQL += " t.timeid, ";
        sSQL += " t.activityno, ";
        sSQL += " t.startdate AS classtime_startdate, ";
        sSQL += " t.enddate AS classtime_enddate, ";
        sSQL += " t.starttime, ";
        sSQL += " t.endtime, ";
        sSQL += " t.min, ";
        sSQL += " t.max, ";
        sSQL += " t.waitlistmax, ";
        sSQL += " t.instructorid, ";
        sSQL += " t.enrollmentsize, ";
        sSQL += " t.waitlistsize, ";
        sSQL += " t.iscanceled, ";
        sSQL += " t.meetingcount, ";
        sSQL += " t.totalhours, ";
        sSQL += " t.rentalid, ";
        sSQL += " t.reservationid, ";
        sSQL += " i.instructorid AS classinstructor_instructorid, ";
        sSQL += " i.firstname, ";
        sSQL += " i.middle, ";
        sSQL += " i.lastname, ";
        sSQL += " i.bio, ";
        sSQL += " i.imgurl AS classinstructor_imgurl, ";
        sSQL += " i.imgalt, ";
        sSQL += " i.email, ";
        sSQL += " i.isemailpublic, ";
        sSQL += " i.phone, ";
        sSQL += " i.isphonepublic, ";
        sSQL += " i.cellphone, ";
        sSQL += " i.iscellpublic, ";
        sSQL += " i.websiteurl, ";
        sSQL += " i.userid, ";
        sSQL += " poc.name, ";
        sSQL += " poc.phone AS pointofcontact_phone, ";
        sSQL += " poc.email AS pointofcontact_email, ";
        sSQL += " o.optionid AS pointofcontact_optionid, ";
        sSQL += " o.optionname, ";
        sSQL += " o.optiondescription, ";
        sSQL += " o.canpurchase, ";
        sSQL += " o.displaytopublic, ";
        sSQL += " o.buttonprefix, ";
        sSQL += " o.requirestime, ";
        sSQL += " o.optiontype, ";
        sSQL += " o.isregistrationrequired, ";
        sSQL += " S.seasonid, ";
        sSQL += " S.seasonyear, ";
        sSQL += " S.seasonname, ";
        sSQL += " S.registrationstartdate AS classseasons_registrationstartdate, ";
        sSQL += " S.registrationenddate AS classseasons_registrationenddate, ";
        sSQL += " S.publicationstartdate, ";
        sSQL += " S.publicationenddate, ";
        sSQL += " S.isclosed, ";
        sSQL += " S.showpublic, ";
        sSQL += " S.isrosterdefault, ";
        sSQL += " ISNULL(C.startdate, 0) AS startdate, ";
        sSQL += " ISNULL(C.enddate, 0) AS enddate, ";
        sSQL += " ISNULL(C.imgurl, 'EMPTY') AS imgurl, ";
        sSQL += " i.firstname + ' ' + i.lastname AS instructor, ";
        sSQL += " ISNULL(C.registrationstartdate, GETDATE()) AS registrationstartdate, ";
        sSQL += " ISNULL(C.registrationenddate, 0) AS registrationenddate ";
        sSQL += " FROM egov_class AS C ";
        sSQL +=      " LEFT OUTER JOIN egov_class_time as t ON C.classid = t.classid AND t.iscanceled = 0 ";
        sSQL +=      " LEFT OUTER JOIN egov_class_instructor as i ON t.instructorid = i.instructorid ";
        sSQL +=      " LEFT OUTER JOIN egov_class_pointofcontact as poc ON C.pocid = poc.pocid ";
        sSQL +=      " LEFT OUTER JOIN egov_registration_option as o ON C.optionid = o.optionid ";
        sSQL +=      " LEFT OUTER JOIN egov_class_seasons AS S ON C.classseasonid = S.classseasonid AND S.showpublic = 1 ";
        //sSQL += " WHERE t.iscanceled = 0 ";
        //sSQL += " AND S.showpublic = 1 ";
        sSQL += " WHERE C.classid = " + iClassID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (Convert.ToInt32(myReader["optionid"]) == 1)
            {

                if (Convert.ToBoolean(myReader["allowearlyregistration"]))
                {
                    sEarlyClassLabel = classes.getEarlyClassLabel(iClassID.ToString());

                    Response.Write("      <tr>");
                    Response.Write("          <td class=\"classdetails_label\">Early Registration Starts:</td>");
                    Response.Write("          <td>" + string.Format("{0:dddd, MMMM dd, yyyy}", Convert.ToDateTime(myReader["earlyregistrationdate"])) + "</td>");
                    Response.Write("      </tr>");
		    if (Convert.ToDateTime(myReader["earlyregistrationdate"]) <= sEarlistRegistrationStart)
		    {
		    	sEarlistRegistrationStart = Convert.ToDateTime(myReader["earlyregistrationdate"]);
		    }
                }

	    }

            //sClassName         = myReader["classname"].ToString().Trim();
            //sClassDescription  = myReader["classdescription"].ToString().Trim();
            //utf8EncodedString = myReader["classdescription"].ToString().Trim();
            //sClassDescription = System.Text.Encoding.UTF8.GetString(utf8EncodedString.GetNonUnicodeBytes());
            sClassName        = common.decodeUTFString(myReader["classname"].ToString().Trim());
            sClassDescription = common.decodeUTFString(myReader["classdescription"].ToString().Trim());

            sCanPurchase       = Convert.ToBoolean(myReader["canpurchase"]);
            sPublicCanOnlyView = Convert.ToBoolean(myReader["publiccanonlyview"]);
            sRegistrationLine  = classes.displayRegistrationLine(iClassID.ToString());
            sIsClassFilled     = classes.classIsFilled(iClassID);
            
            sDateLine = classes.buildDateLine(Convert.ToDateTime(myReader["startdate"]),
                                              Convert.ToDateTime(myReader["enddate"]),
                                              iClassID,
                                              Convert.ToInt32(myReader["classtypeid"]),
                                              Convert.ToBoolean(myReader["isparent"]));

            sClassFeeLine = classes.buildClassFeeLine(iClassID,
                                                      sClassFeesIncludeContainerDIV,
                                                      sClassFeesIncludeContainerTABLE);

            sBreadCrumbsLine = classes.buildBreadCrumbsCategories(iOrgID,
                                                                  sBreadCrumbsReturnLocation,
                                                                  iCategoryID,
                                                                  iClassID);

            if ((myReader["imgurl"].ToString().Trim() != "EMPTY") && (myReader["imgurl"] != null))
            {
                sImgURL = myReader["imgurl"].ToString().Trim();
                sImgALT = myReader["imgalttag"].ToString().Trim();
            }

            if ((myReader["externalurl"].ToString().Trim() != "") && (myReader["externalurl"] != null))
            {
                sExternalURL      = myReader["externalurl"].ToString().Trim();
                sExternalLinkText = myReader["externallinktext"].ToString().Trim();
            }

            Response.Write(sBreadCrumbsLine);
            Response.Write("<fieldset class=\"classdetails_fieldset\">");
            Response.Write("  <legend>" + sClassName + "</legend>");
            Response.Write("  <div class=\"categoryDateLine\">" + sDateLine + "</div>");
            Response.Write("  <div class=\"categoryInfo\">");

            //-- Class Image --//
            if ((sImgURL != "EMPTY") && (sImgURL != null) && (sImgURL != ""))
            {
                Response.Write("<img src=\"" + sImgURL.Replace("http://www.egovlink","https://www.egovlink") + "\" class=\"categoryimage\" border=\"0\" align=\"left\" title=\"" + sImgALT + "\" />");
            }

            //-- Class Description --//
            Response.Write(sClassDescription);

            //-- External URL --//
            if (sExternalURL != "")
            {
                Response.Write("<div id=\"classdetails_externalurl\">");
                Response.Write("  <a href=\"" + sExternalURL + "\">" + sExternalLinkText + "</a>");
                Response.Write("</div>");
            }

            //-- Class Fees --//
            Response.Write(sClassFeeLine);

            //-- Registration Closed Message --//
            if ((string.Format("{0:MM/dd/yyyy}", Convert.ToDateTime(myReader["registrationenddate"])) != "01/01/1900") &&
                (Convert.ToDateTime(string.Format("{0:MM/dd/yyyy}", DateTime.Now.ToString())) > Convert.ToDateTime(string.Format("{0:MM/dd/yyyy}", Convert.ToDateTime(myReader["registrationenddate"])))))
            {
                sRegistrationEnded = true;

                displayClassDetailsMessage("Registration for this event closed on " + string.Format("{0:MM/dd/yyyy}", Convert.ToDateTime(myReader["registrationenddate"])) + ".");
            }

            //Response.Write(sRegistrationLine);
            Response.Write("  </div>");
            Response.Write("</fieldset>");


            //-- Continue Registration --//
            if (sCanPurchase)
            {
                if (Convert.ToDateTime(string.Format("{0:MM/dd/yyyy}", sEarlistRegistrationStart)) <= Convert.ToDateTime(string.Format("{0:MM/dd/yyyy}", DateTime.Now.ToString())) && !sRegistrationEnded)
                {
                    if (!sIsClassFilled)
                    {
                        if (!sPublicCanOnlyView)
                        {
                            displayContinueRegButton("TOP",
                                                     iClassID,
                                                     iCategoryID);
                        }
                    }
                    else
                    {
                        displayClassDetailsMessage("This recreation activity is filled.  Registration is NOT available.");
                    }
                }
            }
            else
            {
                //Cannot purchase this class/event
                displayClassDetailsMessage(myReader["optionname"].ToString() + " - " + myReader["optiondescription"].ToString());
            }

            Response.Write("<div class=\"classdetails\">");

            //-- Availibility --//
            Response.Write("  <fieldset id=\"classdetails_availability\" class=\"classdetails_fieldset\">");
            Response.Write("    <legend>Availability</legend>");
                                displayClassAvailability(iClassID, 
                                                         sTimeID);
            Response.Write("  </fieldset>");

            //-- Details --//
            Response.Write("  <fieldset id=\"classdetails_details\" class=\"classdetails_fieldset\">");
            Response.Write("    <legend>Details</legend>");
            Response.Write("    <table class=\"classdetails_table\">");

            if (Convert.ToInt32(myReader["optionid"]) == 1)
            {
                sRegistrationDate_Start = classes.getRegistrationStarts(iClassID);
                sRegistrationDate_End   = string.Format("{0:dddd, MMMM dd, yyyy}", Convert.ToDateTime(myReader["registrationenddate"]));

                if (Convert.ToBoolean(myReader["allowearlyregistration"]))
                {
                    sEarlyClassLabel = classes.getEarlyClassLabel(iClassID.ToString());

                    Response.Write("      <tr>");
                    Response.Write("          <td class=\"classdetails_label\">Early Registration Starts:</td>");
                    Response.Write("          <td>" + string.Format("{0:dddd, MMMM dd, yyyy}", Convert.ToDateTime(myReader["earlyregistrationdate"])) + "</td>");
                    Response.Write("      </tr>");
                }

                Response.Write("      <tr valign=\"top\">");
                Response.Write("          <td class=\"classdetails_label\">Registration Starts:</td>");
                Response.Write("          <td>" + sRegistrationDate_Start + "</td>");
                Response.Write("          <td>&nbsp;</td>");
                Response.Write("      </tr>");

                if (string.Format("{0:MM/dd/yyyy}", Convert.ToDateTime(myReader["registrationenddate"])) != "01/01/1900")
                {
                    Response.Write("      <tr>");
                    Response.Write("          <td class=\"classdetails_label\">Registration Ends:</td>");
                    Response.Write("          <td>" + sRegistrationDate_End + "</td>");
                    Response.Write("      </tr>");
                }
            }

            if ((myReader["startdate"].ToString().Trim() != "") && (myReader["startdate"] != null))
            {
                Response.Write("      <tr>");
                Response.Write("          <td class=\"classdetails_label\">Start Date:</td>");
                Response.Write("          <td>" + string.Format("{0:dddd, MMMM dd, yyyy}", Convert.ToDateTime(myReader["startdate"])) + "</td>");
                Response.Write("      </tr>");
            }

            if ((myReader["enddate"].ToString().Trim() != "") && (myReader["enddate"] != null))
            {
                Response.Write("      <tr>");
                Response.Write("          <td class=\"classdetails_label\">End Date:</td>");
                Response.Write("          <td>" + string.Format("{0:dddd, MMMM dd, yyyy}", Convert.ToDateTime(myReader["enddate"])) + "</td>");
                Response.Write("      </tr>");
            }

            if ((myReader["alternatedate"].ToString().Trim() != "") && (myReader["alternatedate"] != null))
            {
                Response.Write("      <tr>");
                Response.Write("          <td class=\"classdetails_label\">Make Up Date:</td>");
                Response.Write("          <td>" + string.Format("{0:dddd, MMMM dd, yyyy}", Convert.ToDateTime(myReader["alternatedate"])) + "</td>");
                Response.Write("      </tr>");
            }

            if ((myReader["minage"].ToString().Trim() != "") && (myReader["minage"] != null))
            {
                try
                {
                    sMinage = Convert.ToDouble(myReader["minage"]);
                }
                catch
                {
                    sMinage = 0;
                }

                try
                {
                    sMaxage = Convert.ToDouble(myReader["maxage"]);
                }
                catch
                {
                    sMaxage = 0;
                }

                if (sMinage > 0)
                {
                    sDisplayMinage = Convert.ToString(sMinage);
                }

                if (sMaxage > 0)
                {
                    sDisplayMaxage = Convert.ToString(sMaxage);
                }

                Response.Write("      <tr>");
                Response.Write("          <td class=\"classdetails_label\">Minimum Age:</td>");
                Response.Write("          <td>" + sDisplayMinage + "</td>");
                Response.Write("      </tr>");
            }

            if ((myReader["maxage"].ToString().Trim() != "") && (myReader["maxage"] != null))
            {
                Response.Write("      <tr>");
                Response.Write("          <td class=\"classdetails_label\">Maximum Age:</td>");
                Response.Write("          <td>" + sDisplayMaxage + "</td>");
                Response.Write("      </tr>");
            }

            //-- Gender Restriction --//
            if (lcl_orghasfeature_gender_restriction)
            {
                if ((myReader["genderrestrictionid"].ToString().Trim() != "") && (myReader["genderrestrictionid"] != null))
                {
                    sGenderNotRequiredID = classes.getGenderNotRequiredID();

                    if (sGenderNotRequiredID != Convert.ToInt32(myReader["genderrestrictionid"]))
                    {
                        sGenderRestriction = classes.getGenderRestrictionText(Convert.ToInt32(myReader["genderrestrictionid"]));

                        Response.Write("      <tr>");
                        Response.Write("          <td class=\"classdetails_label\">Gender Restriction:</td>");
                        Response.Write("          <td>" + sGenderRestriction + "</td>");
                        Response.Write("      </tr>");
                    }
                }
            }

            //-- Instructor Name(s) --//
            displayInstructors(iClassID);

            //-- Point of Contact Info --//
            if ((myReader["name"].ToString().Trim() != "")                 && (myReader["name"] != null) ||
                (myReader["pointofcontact_phone"].ToString().Trim() != "") && (myReader["pointofcontact_phone"] != null) ||
                (myReader["pointofcontact_email"].ToString().Trim() != "") && (myReader["pointofcontact_email"] != null))
            {
                Response.Write("      <tr valign=\"top\">");
                Response.Write("          <td class=\"classdetails_label\">Point of Contact Information:</td>");
                Response.Write("          <td>");
                Response.Write("            <table class=\"pointofcontact_info_table\">");

                //-- Point of Contact Info: Name --//
                if ((myReader["name"].ToString().Trim() != "") && (myReader["name"] != null))
                {
                    sPOCName = common.decodeUTFString(myReader["name"].ToString().Trim());

                    Response.Write("              <tr>");
                    Response.Write("                  <td class=\"pointofcontact_info_label\">Name:</td>");
                    Response.Write("                  <td>" + sPOCName + "</td>");
                    Response.Write("              </tr>");
                }

                //-- Point of Contact Info: Phone --//
                if ((myReader["pointofcontact_phone"].ToString().Trim() != "") && (myReader["pointofcontact_phone"] != null))
                {
                    sPhoneNumber = common.formatPhoneNumber(myReader["pointofcontact_phone"].ToString().Trim());
                    sPhoneNumber = "<a href=\"tel:" + myReader["pointofcontact_phone"].ToString().Trim() + "\">" + sPhoneNumber + "</a>";

                    Response.Write("              <tr>");
                    Response.Write("                  <td class=\"pointofcontact_info_label\">Phone:</td>");
                    Response.Write("                  <td>" + sPhoneNumber + "</td>");
                    Response.Write("              </tr>");
                }

                //-- Point of Contact Info: Email --//
                if ((myReader["pointofcontact_email"].ToString().Trim() != "") && (myReader["pointofcontact_email"] != null))
                {
                    sPOCEmail = myReader["pointofcontact_email"].ToString().Trim();
                    sPOCEmail = "<a href=\"mailto:" + myReader["pointofcontact_email"].ToString().Trim() + "\">" + sPOCEmail + "</a>";
                    Response.Write("              <tr>");
                    Response.Write("                  <td class=\"pointofcontact_info_label\">Email:</td>");
                    Response.Write("                  <td>" + sPOCEmail + "</td>");
                    Response.Write("              </tr>");
                }

                Response.Write("            </table>");
                Response.Write("          </td>");
                Response.Write("      </tr>");
            }

            //Response.Write("<tr><td colspan=\"2\">LEFT OFF HERE line 279 class_details.aspx</td></tr>");

            Response.Write("    </table>");
            Response.Write("  </fieldset>");

            //-- Location --//
            if ((myReader["locationid"].ToString().Trim() != "") && (myReader["locationid"] != null))
            {
                sLocationID = Convert.ToInt32(myReader["locationid"]);
            }

            Response.Write("  <fieldset id=\"classdetails_location\" class=\"classdetails_fieldset\">");
            Response.Write("    <legend>Location</legend>");
                                displayLocationInfo(iOrgID,
                                                    sLocationID,
                                                    sDisplayDirections);
            Response.Write("  </fieldset>");

            //-- Continue Registration --//
            if (sCanPurchase)
            {
                if (Convert.ToDateTime(string.Format("{0:MM/dd/yyyy}", sEarlistRegistrationStart)) <= Convert.ToDateTime(string.Format("{0:MM/dd/yyyy}", DateTime.Now.ToString())) && !sRegistrationEnded)
                {
                    if (!sIsClassFilled)
                    {
                        if (!sPublicCanOnlyView)
                        {
                            displayContinueRegButton("BOTTOM",
                                                     iClassID,
                                                     iCategoryID);
                        }
                    }
                }
            }
            else
            {
                //Cannot purchase this class/event
                displayClassDetailsMessage("No information for the class or event could be found.");
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
    }

    public void displayContinueRegButton(string iLocation,
                                         Int32 iClassID,
                                         Int32 iCategoryID)
    {
        string sCategoryTitle = classes.getCategoryTitle(iCategoryID);

        Response.Write("<div class=\"continueRegistrationDiv\">");
        //Response.Write("  <input type=\"button\" name=\"continueRegistrationButton" + iLocation + "\" id=\"continueRegistrationButton" + iLocation + "\" value=\"Continue Registration\" class=\"continueRegistrationButton\" onclick=\"goToSignUp('" + iClassID.ToString() + "','" + iCategoryID.ToString() + "','" + sCategoryTitle + "');\" />");
        Response.Write("  <input type=\"button\" name=\"continueRegistrationButton" + iLocation + "\" id=\"continueRegistrationButton" + iLocation + "\" value=\"Continue Registration\" class=\"continueRegistrationButton\" onclick=\"goToSignUp('" + iClassID.ToString() + "','" + iCategoryID.ToString() + "');\" />");
        Response.Write("</div>");
    }

    public void displayClassAvailability(Int32 iClassID,
                                         Int32 iTimeID)
    {
        string sSQL            = "";
        string sActivityNo     = "&nbsp;";
        string sOldActivityNo  = "&nbsp;";
        string sInstructorName = "&nbsp;";
        string sClassSizeMin   = "&nbsp;";
        string sClassSizeMax   = "&nbsp;";
        string sEnrollmentSize = "&nbsp;";
        string sWaitListSize   = "&nbsp;";
        string sTotalEnrolled  = "&nbsp;";
        string sSunday         = "&nbsp;";
        string sMonday         = "&nbsp;";
        string sTuesday        = "&nbsp;";
        string sWednesday      = "&nbsp;";
        string sThursday       = "&nbsp;";
        string sFriday         = "&nbsp;";
        string sSaturday       = "&nbsp;";
        string sStartTime      = "&nbsp;";
        string sEndTime        = "&nbsp;";
        string sBGColor        = "#eeeeee";
        string sRowClass       = "";

        Boolean sIsClassSeries = false;

        Int32 sSeriesEnrollment = 0;

        sSQL  = "SELECT t.timeid, ";
        sSQL += " activityno, ";
        sSQL += " min, ";
        sSQL += " max, ";
        sSQL += " waitlistmax, ";
        sSQL += " isnull(instructorid,0) as instructorid, ";
        sSQL += " enrollmentsize, ";
        sSQL += " waitlistsize, ";
        sSQL += " sunday, ";
        sSQL += " monday, ";
        sSQL += " tuesday, ";
        sSQL += " wednesday, ";
        sSQL += " thursday, ";
        sSQL += " friday, ";
        sSQL += " saturday, ";
        sSQL += " d.starttime, ";
        sSQL += " d.endtime ";
        sSQL += " FROM egov_class_time t, ";
        sSQL +=      " egov_class_time_days d ";
        sSQL += " WHERE t.timeid = d.timeid ";
        sSQL += " AND t.iscanceled = 0 ";
        sSQL += " AND classid = " + iClassID.ToString();

        if (iTimeID != 0)
        {
            sSQL += " AND t.timeid = " + iTimeID.ToString();
        }

        sSQL += " ORDER BY activityno, timedayid ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            Response.Write("<table cellspacing=\"0\" id=\"offeringActivities\">");
            Response.Write("  <thead>");
            Response.Write("  <tr>");
            Response.Write("      <td>Activity No</td>");
            Response.Write("      <td>Instructor</td>");
            Response.Write("      <td>Min</td>");
            Response.Write("      <td>Max</td>");
            Response.Write("      <td>Enrld</td>");
            Response.Write("      <td>Su</td>");
            Response.Write("      <td>Mo</td>");
            Response.Write("      <td>Tu</td>");
            Response.Write("      <td>We</td>");
            Response.Write("      <td>Th</td>");
            Response.Write("      <td>Fr</td>");
            Response.Write("      <td>Sa</td>");
            Response.Write("      <td>Starts</td>");
            Response.Write("      <td>Ends</td>");
            Response.Write("  </tr>");
            Response.Write("  </thead>");

            while (myReader.Read())
            {
                sInstructorName = classes.getInstructorName("LASTNAME", Convert.ToInt32(myReader["instructorid"]));
                sInstructorName = common.decodeUTFString(sInstructorName);

                sBGColor        = common.changeBGColor(sBGColor, "#ffffff", "#eeeeee");
                sRowClass       = "offeringActivities_row_light";
                sActivityNo     = "";
                sClassSizeMin   = "";
                sClassSizeMax   = "";
                sEnrollmentSize = "";
                sWaitListSize   = "";
                sSunday         = "";
                sMonday         = "";
                sTuesday        = "";
                sWednesday      = "";
                sThursday       = "";
                sFriday         = "";
                sSaturday       = "";
                sStartTime      = "";
                sEndTime        = "";

                if (sBGColor == "#eeeeee")
                {
                    sRowClass = "offeringActivities_row_dark";
                }

                if (myReader["activityno"].ToString() != null)
                {
                    sActivityNo = myReader["activityno"].ToString();
                    sActivityNo = common.decodeUTFString(sActivityNo);
                }

                if (myReader["min"].ToString() != null)
                {
                    sClassSizeMin = myReader["min"].ToString();
                }

                if (myReader["max"].ToString() != null)
                {
                    sClassSizeMax = myReader["max"].ToString();
                }

                if (myReader["enrollmentsize"].ToString() != null)
                {
                    sEnrollmentSize = myReader["enrollmentsize"].ToString();
                }

                if (myReader["waitlistsize"].ToString() != null)
                {
                    sWaitListSize = myReader["waitlistsize"].ToString();
                }

                if (Convert.ToBoolean(myReader["sunday"]))
                {
                    sSunday = "Su";
                }

                if (Convert.ToBoolean(myReader["monday"]))
                {
                    sMonday = "Mo";
                }

                if (Convert.ToBoolean(myReader["tuesday"]))
                {
                    sTuesday = "Tu";
                }

                if (Convert.ToBoolean(myReader["wednesday"]))
                {
                    sWednesday = "We";
                }

                if (Convert.ToBoolean(myReader["thursday"]))
                {
                    sThursday = "Th";
                }

                if (Convert.ToBoolean(myReader["friday"]))
                {
                    sFriday = "Fr";
                }

                if (Convert.ToBoolean(myReader["saturday"]))
                {
                    sSaturday = "Sa";
                }

                if (myReader["starttime"].ToString() != null)
                {
                    sStartTime = myReader["starttime"].ToString();
                }

                if (myReader["endtime"].ToString() != null)
                {
                    sEndTime = myReader["endtime"].ToString();
                }

                Response.Write("  <tbody>");
                Response.Write("  <tr class=\"" + sRowClass + "\">");

                if (sActivityNo != sOldActivityNo)
                {
                    sOldActivityNo = sActivityNo;
                    sIsClassSeries = classes.classIsSeries(iClassID);

                    //Find the total enrollment size;
                    if (sIsClassSeries)
                    {
                        sSeriesEnrollment = classes.getSeriesEnrollment(iClassID);
                        sTotalEnrolled    = sSeriesEnrollment.ToString();
                    }
                    else
                    {
                        sTotalEnrolled = Convert.ToString(Convert.ToInt32(sEnrollmentSize) + Convert.ToInt32(sWaitListSize));
                    }

                    Response.Write("      <td>" + sActivityNo     + "</td>");
                    Response.Write("      <td>" + sInstructorName + "</td>");
                    Response.Write("      <td>" + sClassSizeMin   + "</td>");
                    Response.Write("      <td>" + sClassSizeMax   + "</td>");
                    Response.Write("      <td>" + sTotalEnrolled  + "</td>");
                }
                else
                {
                    Response.Write("      <td colspan=\"5\">&nbsp;</td>");
                }

                Response.Write("      <td>" + sSunday         + "</td>");
                Response.Write("      <td>" + sMonday         + "</td>");
                Response.Write("      <td>" + sTuesday        + "</td>");
                Response.Write("      <td>" + sWednesday      + "</td>");
                Response.Write("      <td>" + sThursday       + "</td>");
                Response.Write("      <td>" + sFriday         + "</td>");
                Response.Write("      <td>" + sSaturday       + "</td>");
                Response.Write("      <td>" + sStartTime      + "</td>");
                Response.Write("      <td>" + sEndTime        + "</td>");
                Response.Write("  </tr>");
                Response.Write("  </tbody>");
            }

            Response.Write("</table>");
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
    }

    public void displayInstructors(Int32 iClassID)
    {
        Int32 iRowCount = 0;

        string sSQL             = "";
        string sInstructorLabel = "";
        string sInstructorName  = "";

        sSQL  = "SELECT i.instructorid, ";
        sSQL += " i.firstname, ";
        sSQL += " i.lastname, ";
        sSQL += " ltrim(rtrim(i.firstname + ' ' + i.lastname)) as instructorname ";
        sSQL += " FROM egov_class_to_instructor ci, ";
        sSQL +=      " egov_class_instructor i ";
        sSQL += " WHERE ci.instructorid = i.instructorid ";
        sSQL += " AND ci.classid = " + iClassID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                iRowCount        = iRowCount + 1;
                sInstructorLabel = "&nbsp;";
                sInstructorName  = "&nbsp;";

                if (iRowCount == 1)
                {
                    sInstructorLabel = "Instructor(s):";
                }

                if ((myReader["instructorname"].ToString().Trim() != "") && (myReader["instructorname"] != null))
                {
                    sInstructorName = myReader["instructorname"].ToString().Trim();
                }

                Response.Write("      <tr>");
                Response.Write("          <td class=\"classdetails_label\">" + sInstructorLabel + "</td>");
                Response.Write("          <td>" + sInstructorName + "</td>");
                Response.Write("      </tr>");

            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

    }

    public void displayLocationInfo(Int32 iOrgID,
                                    Int32 iLocationID,
                                    Boolean iDisplayDirections)
    {
        string sSQL          = "";
        string sFullAddress  = "";
        string sLocationName = "";

        sSQL  = "SELECT name, ";
        sSQL += " address1, ";
        sSQL += " address2, ";
        sSQL += " city, ";
        sSQL += " state, ";
        sSQL += " zip, ";
        sSQL += " directions ";
        sSQL += " FROM egov_class_location ";
        sSQL += " WHERE orgid = " + iOrgID.ToString();
        sSQL += " AND locationid = " + iLocationID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            Response.Write("<div id=\"classdetails_locationDirections\">");

            //-- Name --//
            if ((myReader["name"].ToString().Trim() != "") && (myReader["name"] != null))
            {
                sLocationName = common.decodeUTFString(myReader["name"].ToString().Trim());

                Response.Write("<strong>" + sLocationName + "</strong><br />");
            }

            //-- Address: Line 1 --//
            if ((myReader["address1"].ToString().Trim() != "") && (myReader["address1"] != null))
            {
                sFullAddress += " " + myReader["address1"].ToString().Trim();

                Response.Write(myReader["address1"].ToString().Trim() + "<br />");
            }

            //-- Address: Line 2 --//
            if ((myReader["address2"].ToString().Trim() != "") && (myReader["address2"] != null))
            {
                sFullAddress += " " + myReader["address2"].ToString().Trim();

                Response.Write(myReader["address2"].ToString().Trim() + "<br />");
            }

            //-- City, State, Zip --//
            if ((myReader["city"].ToString().Trim() != "") && (myReader["city"] != null))
            {
                sFullAddress += ", " + myReader["city"].ToString().Trim() + ", " + myReader["state"].ToString().Trim() + ", " + myReader["zip"].ToString().Trim();

                Response.Write(myReader["city"].ToString().Trim() + ", " + myReader["state"].ToString().Trim() + " " + myReader["zip"].ToString().Trim());
            }

            sFullAddress = common.decodeUTFString(sFullAddress);

            Response.Write("  <input type=\"hidden\" name=\"location_name\" id=\"location_name\" value=\"" + myReader["name"].ToString().Trim() + "\" />");
            Response.Write("  <input type=\"hidden\" name=\"location_fulladdress\" id=\"location_fulladdress\" value=\"" + sFullAddress.Trim() + "\" />");
            Response.Write("  <input type=\"hidden\" name=\"location_latitude\" id=\"location_latitude\" value=\"\" />");
            Response.Write("  <input type=\"hidden\" name=\"location_longitude\" id=\"location_longitude\" value=\"\" />");
            Response.Write("</div>");
            Response.Write("<div id=\"map_canvas\"></div>");
            Response.Write("<input type=\"button\" name=\"location_getDirectionsButton\" id=\"location_getDirectionsButton\" value=\"Get Directions\" />");
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
    }

    public void displayClassDetailsMessage(string iMessage)
    {
        if (iMessage != null && iMessage != "")
        {
            Response.Write("<div class=\"classdetails_message\">" + iMessage + "</div>");
        }
    }

    public void displayRegattaItem(Int32 iOrgID,
                                   Int32 iClassID,
                                   Int32 iCategoryID)
    {
        Boolean sRegistrationStarted            = false;
        Boolean sRegistrationEnded              = false;
        Boolean sClassHasWaivers                = false;
        Boolean sClassIsRegattaTeamSignup       = false;
        Boolean sClassFeesIncludeContainerDIV   = true;
        Boolean sClassFeesIncludeContainerTABLE = true;
        Boolean sShowWaiverText = false;
        Boolean sShowWaiverName = false;
        Boolean sShowWaiverDesc = false;
        Boolean sShowWaiverLink = true;

        string sSQL                       = "";
        string sDateLine                  = "";
        string sClassFeeLine              = "";
        string sBreadCrumbsLine           = "";
        string sBreadCrumbsReturnLocation = "CLASSDETAILS";

        string sClassName              = "";
        string sClassDescription       = "";
        string sRegistrationDate_Start = "";
        string sRegistrationDate_End   = "";
        string sImgURL                 = "";
        string sImgALT                 = "";
        string sExternalURL            = "";
        string sExternalLinkText       = "";
        string sDisplayWaiverList      = "";
        string sButtonFunction         = "";

        sSQL  = "SELECT classid, ";
        sSQL += " classname, ";
        sSQL += " classdescription, ";
        sSQL += " startdate, ";
        sSQL += " enddate, ";
        sSQL += " classtypeid, ";
        sSQL += " isparent, ";
        sSQL += " imgurl, ";
        sSQL += " imgalttag, ";
        sSQL += " externalurl, ";
        sSQL += " externallinktext, ";
        sSQL += " registrationenddate, ";
        sSQL += " optionid ";
        sSQL += " FROM egov_class ";
        sSQL += " WHERE classid = " + iClassID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            sClassName                = myReader["classname"].ToString().Trim();
            sClassDescription         = myReader["classdescription"].ToString().Trim();
            sRegistrationStarted      = classes.isRegistrationStarted(iClassID);
            sClassIsRegattaTeamSignup = classes.classIsRegattaTeamSignup(iClassID);

            sClassFeeLine = classes.buildClassFeeLine(iClassID,
                                                      sClassFeesIncludeContainerDIV,
                                                      sClassFeesIncludeContainerTABLE);

            sDateLine = classes.buildDateLine(Convert.ToDateTime(myReader["startdate"]),
                                              Convert.ToDateTime(myReader["enddate"]),
                                              iClassID,
                                              Convert.ToInt32(myReader["classtypeid"]),
                                              Convert.ToBoolean(myReader["isparent"]));

            sBreadCrumbsLine = classes.buildBreadCrumbsCategories(iOrgID,
                                                                  sBreadCrumbsReturnLocation,
                                                                  iCategoryID,
                                                                  iClassID);

            if ((myReader["imgurl"].ToString().Trim() != "EMPTY") && (myReader["imgurl"] != null))
            {
                sImgURL = myReader["imgurl"].ToString().Trim();
                sImgALT = myReader["imgalttag"].ToString().Trim();
            }

            if ((myReader["externalurl"].ToString().Trim() != "") && (myReader["externalurl"] != null))
            {
                sExternalURL = myReader["externalurl"].ToString().Trim();
                sExternalLinkText = myReader["externallinktext"].ToString().Trim();
            }

            Response.Write(sBreadCrumbsLine);
            Response.Write("<fieldset class=\"classdetails_fieldset\">");
            Response.Write("  <legend>" + sClassName + "</legend>");
            Response.Write("  <div class=\"categoryDateLine\">" + sDateLine + "</div>");
            Response.Write("  <div class=\"categoryInfo\">");

            //-- Class Image --//
            if ((sImgURL != "EMPTY") && (sImgURL != null) && (sImgURL != ""))
            {
                Response.Write("<img src=\"" + sImgURL + "\" class=\"categoryimage\" border=\"0\" align=\"left\" title=\"" + sImgALT + "\" />");
            }

            //-- Class Description --//
            Response.Write(sClassDescription);

            //-- External URL --//
            if (sExternalURL != "")
            {
                Response.Write("<div id=\"classdetails_externalurl\">");
                Response.Write("  <a href=\"" + sExternalURL + "\">" + sExternalLinkText + "</a>");
                Response.Write("</div>");
            }

            //-- Class Fees --//
            Response.Write(sClassFeeLine);

            //-- Registration Closed Message --//
            if ((string.Format("{0:MM/dd/yyyy}", Convert.ToDateTime(myReader["registrationenddate"])) != "01/01/1900") &&
                (Convert.ToDateTime(string.Format("{0:MM/dd/yyyy}", DateTime.Now.ToString())) > Convert.ToDateTime(string.Format("{0:MM/dd/yyyy}", Convert.ToDateTime(myReader["registrationenddate"])))))
            {
                sRegistrationEnded = true;

                displayClassDetailsMessage("Registration for this event closed on " + string.Format("{0:MM/dd/yyyy}", Convert.ToDateTime(myReader["registrationenddate"])) + ".");
            }

            //Response.Write(sRegistrationLine);
            Response.Write("  </div>");
            Response.Write("</fieldset>");

            Response.Write("<div class=\"classdetails\">");

            //-- Details --//
            Response.Write("  <fieldset id=\"classdetails_details\" class=\"classdetails_fieldset\">");
            Response.Write("    <legend>Details</legend>");
            Response.Write("    <table class=\"classdetails_table\">");

            if (Convert.ToInt32(myReader["optionid"]) == 1)
            {
                sRegistrationDate_Start = classes.getRegistrationStarts(iClassID);
                sRegistrationDate_End = string.Format("{0:dddd, MMMM dd, yyyy}", Convert.ToDateTime(myReader["registrationenddate"]));

                Response.Write("      <tr valign=\"top\">");
                Response.Write("          <td class=\"classdetails_label\">Registration Starts:</td>");
                Response.Write("          <td>" + sRegistrationDate_Start + "</td>");
                Response.Write("          <td>&nbsp;</td>");
                Response.Write("      </tr>");

                if (string.Format("{0:MM/dd/yyyy}", Convert.ToDateTime(myReader["registrationenddate"])) != "01/01/1900")
                {
                    Response.Write("      <tr>");
                    Response.Write("          <td class=\"classdetails_label\">Registration Ends:</td>");
                    Response.Write("          <td>" + sRegistrationDate_End + "</td>");
                    Response.Write("      </tr>");
                }
            }

            if ((myReader["startdate"].ToString().Trim() != "") && (myReader["startdate"] != null))
            {
                Response.Write("      <tr>");
                Response.Write("          <td class=\"classdetails_label\">Event Date:</td>");
                Response.Write("          <td>" + string.Format("{0:dddd, MMMM dd, yyyy}", Convert.ToDateTime(myReader["startdate"])) + "</td>");
                Response.Write("      </tr>");
            }

            Response.Write("    </table>");
            Response.Write("  </fieldset>");

            //-- Waivers --//
            sClassHasWaivers = classes.classHasWaivers(iClassID);

            if (sClassHasWaivers)
            {
                sDisplayWaiverList = classes.showWaiverList(iOrgID,
                                                            iClassID,
                                                            sShowWaiverText,
                                                            sShowWaiverName,
                                                            sShowWaiverDesc,
                                                            sShowWaiverLink);

                Response.Write("  <fieldset id=\"classdetails_details\" class=\"classdetails_fieldset\">");
                Response.Write("    <legend>Required Waivers</legend>");
                Response.Write("    <table class=\"classdetails_table\">");
                Response.Write("      <tr>");
                Response.Write("          <td>" + sDisplayWaiverList + "</td>");
                Response.Write("      </tr>");
                Response.Write("    </table>");
                Response.Write("  </fieldset>");
            }

            //-- Continue Registration --//
            if (sRegistrationStarted && (sRegistrationEnded == false))
            {
                sButtonFunction = "MEMBER";

                if (sClassIsRegattaTeamSignup)
                {
                    sButtonFunction = "TEAM";
                }

                Response.Write("<div class=\"continueRegistrationDiv\">");
                Response.Write("  <input type=\"button\" name=\"signup\" id=\"signup\" value=\"Continue Registration\" class=\"continueRegistrationButton\" onclick=\"goToRegattaMaint('" + sButtonFunction + "', '" + iClassID.ToString() + "');\" />");
                Response.Write("</div>");
            }

            Response.Write("</div>");
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
    }
}

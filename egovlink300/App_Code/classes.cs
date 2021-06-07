using System;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
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
/// Summary description for classes
/// </summary>
public class classes
{
    public static Int32 getFirstCategory(string iOrgID)
    {
        Int32 lcl_return = 0;
        Int32 sOrgID = 0;

        string sSQL = "";

        if (iOrgID != null)
        {
            try
            {
                sOrgID = Convert.ToInt32(iOrgID);
            }
            catch
            {
                sOrgID = 0;
            }
        }

        sSQL = "SELECT categoryid ";
        sSQL += " FROM egov_class_categories ";
        sSQL += " WHERE orgid = " + sOrgID;
        sSQL += " AND isroot = 1";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToInt32(myReader["categoryid"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Int32 getFirstCategoryID_byClassID(Int32 iOrgID,
                                                     Int32 iClassID)
    {
        Int32 lcl_return = 0;

        string sSQL = "";

        sSQL  = "SELECT c.categoryid ";
        sSQL += " FROM egov_class_categories c, ";
        sSQL +=      " egov_class_category_to_class t ";
        sSQL += " WHERE c.categoryid = t.categoryid ";
        sSQL += " AND t.classid = " + iClassID.ToString();
        sSQL += " AND orgid = " + iOrgID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToInt32(myReader["categoryid"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string buildDateLine(DateTime iStartDate,
                                       DateTime iEndDate,
                                       Int32 iClassID,
                                       Int32 iClassTypeID,
                                       Boolean isParent)
    {
        string lcl_return = "";
        string sDaySuffix = "";
        string sDates = "";
        string sDatesAndDaysOfWeek = "";
        string sDatesAndSingleDayOfWeek = "";

        TimeSpan ts = iEndDate - iStartDate;

        Int32 sDateDiff = ts.Days;
        Int32 sActivityCount = 0;

        if (iStartDate == iEndDate)
        {
            sDates = string.Format("{0:MMMM d}", iStartDate);
        }
        else
        {
            sDates = string.Format("{0:MMMM d}", iStartDate);
            sDates += " &ndash; ";
            sDates += string.Format("{0:MMMM d}", iEndDate);

            //still need to fix the date range!!!!???

            if (sDateDiff > 7)
            {
                sDaySuffix = "s";
            }
        }

        sActivityCount = getClassActivityCount(iClassID);

        if (sActivityCount < 2)
        {
            //Draw Date Range based on ClassType
            sDatesAndDaysOfWeek = "<div>" + sDates + getClassDaysOfWeek(iClassID, sDaySuffix) + "</div>";
            sDatesAndSingleDayOfWeek = "<div>" + sDates + "<div class=\"classes_daysofweek\">(" + string.Format("{0:dddd}", iStartDate) + ")</div></div>";

            switch (iClassTypeID)
            {
                case 1:
                    lcl_return = sDatesAndDaysOfWeek;
                    break;
                case 2:
                    if (isParent)
                    {
                        lcl_return = "<div class=\"classes_daysofweek\">YEAR-ROUND</div>";
                    }
                    else
                    {
                        lcl_return = sDatesAndDaysOfWeek;
                    }
                    break;
                case 3:
                    lcl_return = sDatesAndDaysOfWeek;
                    break;
                case 4:
                    lcl_return = sDatesAndSingleDayOfWeek;
                    break;
                case 5:
                    lcl_return = sDatesAndSingleDayOfWeek;
                    break;
                default:
                    lcl_return = sDatesAndDaysOfWeek;
                    break;
            }
        }
        else
        {
            lcl_return = "<div>" + sDates + "<div class=\"classes_daysofweek\">(" + sActivityCount + " Activity Sessions Available)</div></div>";
        }

        return lcl_return;
    }

    public static string buildClassFeeLine(Int32 iClassID,
                                           Boolean iIncludeContainerDIV,
                                           Boolean iIncludeContainerTABLE)
    {
        Boolean sIsClassFilled = isClassFilled(iClassID);

        string lcl_return = "";

        if (sIsClassFilled)
        {
            lcl_return = "<div class=\"maxEnrollementMsg\">Maximum enrollment has been reached.</div>";
        }
        else
        {
            Boolean sClassHasDiscount = classHasDiscount(iClassID);

            string sSQL = "";
            string sDisplayPrice = "";
            string sDisplayMembership = "";

            double sFees = classes.getFeeTotal(iClassID);
            double sTotalPrice = 0.00;
            double sBasePrice = 0.00;

            sSQL = " SELECT pt.pricetypename, ";
            sSQL += " pt.basepricetypeid, ";
            sSQL += " isnull(cpp.amount,0.0) as amount, ";
            sSQL += " pt.isdropin, ";
            sSQL += " pt.ismember ";
            sSQL += " FROM egov_class_pricetype_price cpp, ";
            sSQL += " egov_price_types pt ";
            sSQL += " WHERE cpp.pricetypeid = pt.pricetypeid ";
            sSQL += " AND pt.isactiveforclasses = 1 ";
            sSQL += " AND pt.isfee = 0 ";
            sSQL += " AND pt.isdropin = 0 ";
            sSQL += " AND cpp.classid = " + iClassID.ToString();
            sSQL += " ORDER BY pt.displayorder ";

            SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
            sqlConn.Open();

            SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
            SqlDataReader myReader;
            myReader = myCommand.ExecuteReader();

            if (myReader.HasRows)
            {
                if (iIncludeContainerDIV)
                {
                    lcl_return = "<div id=\"classdetails_classfeeline\">";
                }

                if (iIncludeContainerTABLE)
                {
                    lcl_return += "  <table id=\"classdetails_classfee_table\">";
                }

                lcl_return += "    <tr valign=\"top\">";
                lcl_return += "        <td><strong>Fee: </strong></td>";
                lcl_return += "        <td>";
                lcl_return += "            <table id=\"classdetails_classfee_list\">";

                while (myReader.Read())
                {
                    sDisplayPrice = "";
                    sTotalPrice = sFees + Convert.ToDouble(myReader["amount"]);
                    sDisplayMembership = "";

                    if (myReader["basepricetypeid"] != null && myReader["basepricetypeid"].ToString() != "")
                    {
                        //add the fees: this price and its base price
                        sBasePrice = getBasePrice(Convert.ToInt32(myReader["basepricetypeid"]), iClassID);
                        sTotalPrice = sTotalPrice + sBasePrice;
                    }

                    if (Convert.ToBoolean(myReader["ismember"]))
                    {
                        sDisplayMembership = showMembership(iClassID);
                    }

                    sDisplayPrice = string.Format("{0:C}", sTotalPrice);
                    sDisplayPrice += sDisplayMembership;

                    lcl_return += "              <tr>";
                    lcl_return += "                  <td class=\"classdetails_label\">" + myReader["pricetypename"].ToString() + ":</td>";
                    lcl_return += "                  <td>" + sDisplayPrice + "</td>";
                    lcl_return += "              </tr>";
                }

                lcl_return += "            </table>";

                if (sClassHasDiscount)
                {
                    lcl_return += showDiscountPhrase(iClassID);
                }


                lcl_return += "        </td>";
                lcl_return += "    </tr>";

                if (iIncludeContainerTABLE)
                {
                    lcl_return += "  </table>";
                }

                if (iIncludeContainerDIV)
                {
                    lcl_return += "</div>";
                }
            }

            myReader.Close();
            sqlConn.Close();
            myReader.Dispose();
            sqlConn.Dispose();
        }

        return lcl_return;
    }

    public static Boolean classHasDiscount(Int32 iClassID)
    {
        Boolean lcl_return = false;

        Int32 sPriceDiscountID = 0;

        string sSQL = "";

        sSQL  = "SELECT isnull(pricediscountid,0) AS pricediscountid ";
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

            sPriceDiscountID = Convert.ToInt32(myReader["pricediscountid"]);

            if (sPriceDiscountID > 0)
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

    public static string showDiscountPhrase(Int32 iClassID)
    {
        string lcl_return = "";
        string sSQL       = "";

        sSQL  = "SELECT discountamount, ";
        sSQL += " discountdescription ";
        sSQL += " FROM egov_price_discount p, ";
        sSQL +=      " egov_class c ";
        sSQL += " WHERE c.pricediscountid = p.pricediscountid ";
        sSQL += " AND c.classid = " + iClassID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = "<div class=\"classes_discountdesc\">" + myReader["discountdescription"].ToString() + "</div>";
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static double getFeeTotal(Int32 iClassID)
    {
        double lcl_return = 0.0;

        string sSQL = "";

        sSQL  = "SELECT SUM(amount) as amount ";
        sSQL += " FROM egov_class_pricetype_price cpp, ";
        sSQL +=      " egov_price_types pt ";
        sSQL += " WHERE cpp.pricetypeid = pt.pricetypeid ";
        sSQL += " AND isactiveforclasses = 1 ";
        sSQL += " AND isfee = 1 ";
        sSQL += " AND classid = " + iClassID.ToString();
        sSQL += " GROUP BY classid ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (myReader["amount"] != null)
            {
                lcl_return = Convert.ToDouble(myReader["amount"]);
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static double getBasePrice(Int32 iBasePriceTypeID, Int32 iClassID)
    {
        double lcl_return = 0.00;

        string sSQL = "";

        sSQL  = "SELECT isnull(amount,0.00) as amount ";
        sSQL += " FROM egov_class_pricetype_price ";
        sSQL += " WHERE pricetypeid = " + iBasePriceTypeID.ToString();
        sSQL += " AND classid = " + iClassID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToDouble(myReader["amount"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string showMembership(Int32 iClassID)
    {
        string lcl_return = "";
        string sSQL       = "";

        sSQL  = "SELECT membershipdesc ";
        sSQL += " FROM egov_memberships m, ";
        sSQL +=      " egov_class c ";
        sSQL += " WHERE c.classid = " + iClassID.ToString();
        sSQL += " AND c.membershipid = m.membershipid ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = "&nbsp;<span class=\"classdetails_membershiprequired\">(" + myReader["membershipdesc"].ToString() + " Membership Required)</span>";
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Int32 getClassActivityCount(Int32 iClassID)
    {
        string sSQL = "";

        Int32 lcl_return = 0;

        sSQL = "SELECT count(timeid) as hits ";
        sSQL += " FROM egov_class_time ";
        sSQL += " WHERE iscanceled = 0 ";
        sSQL += " AND classid = " + iClassID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToInt32(myReader["hits"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getClassDaysOfWeek(Int32 iClassID,
                                            string iDaySuffix)
    {
        string lcl_return = "";
        string sEnabledDays         = "";
        string sDayEnabledSunday    = "";
        string sDayEnabledMonday    = "";
        string sDayEnabledTuesday   = "";
        string sDayEnabledWednesday = "";
        string sDayEnabledThursday  = "";
        string sDayEnabledFriday    = "";
        string sDayEnabledSaturday  = "";

        Int32 sEnabledDaysLength  = 0;
        Int32 sCommaCount         = 0;
        Int32 sFirstCommaPosition = 0;
        Int32 sLastCommaPosition  = 0;

        sDayEnabledSunday    = isDayEnabled(iClassID, iDaySuffix, "Sunday");
        sDayEnabledMonday    = isDayEnabled(iClassID, iDaySuffix, "Monday");
        sDayEnabledTuesday   = isDayEnabled(iClassID, iDaySuffix, "Tuesday");
        sDayEnabledWednesday = isDayEnabled(iClassID, iDaySuffix, "Wednesday");
        sDayEnabledThursday  = isDayEnabled(iClassID, iDaySuffix, "Thursday");
        sDayEnabledFriday    = isDayEnabled(iClassID, iDaySuffix, "Friday");
        sDayEnabledSaturday  = isDayEnabled(iClassID, iDaySuffix, "Saturday");

        sEnabledDays = sDayEnabledSunday;
        sEnabledDays += sDayEnabledMonday;
        sEnabledDays += sDayEnabledTuesday;
        sEnabledDays += sDayEnabledWednesday;
        sEnabledDays += sDayEnabledThursday;
        sEnabledDays += sDayEnabledFriday;
        sEnabledDays += sDayEnabledSaturday;

        sEnabledDaysLength = sEnabledDays.Length;

        for (int i = 0; i < sEnabledDaysLength; i++)
        {
            if (sEnabledDays[i].ToString() == ",")
            {
                sCommaCount++;
            }
        }

        //If there is more than comma then replace the last one with "AND"
        if (sCommaCount > 1)
        {
            sLastCommaPosition = sEnabledDays.LastIndexOf(",");
            sEnabledDays       = sEnabledDays.Remove(sLastCommaPosition, 1).Insert(sLastCommaPosition, " and");
        }

        //Replace the first comma with the proper formatting.
        if (sCommaCount > 0)
        {
            sFirstCommaPosition = sEnabledDays.IndexOf(",");
            sEnabledDays        = sEnabledDays.Remove(sFirstCommaPosition, 2).Insert(sFirstCommaPosition, "<div class=\"classes_daysofweek\">(");
            sEnabledDays        = sEnabledDays + ")</div>";
        }

        lcl_return = sEnabledDays;

        return lcl_return;
    }

    public static string isDayEnabled(Int32 iClassID,
                                  string iDaySuffix,
                                  string iDayOfWeek)
    {
        string lcl_return = "";
        string sSQL = "";

        sSQL = "SELECT timedayid ";
        sSQL += " FROM egov_class_time t, ";
        sSQL += " egov_class_time_days d ";
        sSQL += " WHERE t.timeid = d.timeid ";
        sSQL += " AND t.iscanceled = 0 ";
        sSQL += " AND t.classid = " + iClassID.ToString();
        sSQL += " AND d." + iDayOfWeek + " = 1 ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = ", " + iDayOfWeek + iDaySuffix;
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Boolean classIsRegattaEvent(Int32 iClassID)
    {
        Boolean lcl_return = false;

        string sSQL = "";

        sSQL  = "SELECT isregatta ";
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

            if (Convert.ToBoolean(myReader["isregatta"]))
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

    public static string displayRegistrationLine(string iClassID)
    {
        string lcl_return = "";
        string sSQL = "";
        string sAllowEarlyRegistration = "false";
        string sEarlyClassLabel = "";
        string sRegistrationStartDates = "";

        DateTime sEarlyRegistrationDate;

        Int32 sClassID = 0;
        Int32 sOptionID = 0;

        if (iClassID != null)
        {
            try
            {
                sClassID = Convert.ToInt32(iClassID);
            }
            catch
            {
                sClassID = 0;
            }
        }

        sSQL = "SELECT optionid, ";
        sSQL += " isnull(registrationenddate,0) as registrationenddate, ";
        sSQL += " allowearlyregistration, ";
        sSQL += " earlyregistrationdate, ";
        sSQL += " earlyregistrationclassid ";
        sSQL += " FROM egov_class ";
        sSQL += " WHERE egov_class.classid = " + sClassID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            sOptionID = Convert.ToInt32(myReader["optionid"]);

            if (sOptionID == 1)
            {
                sAllowEarlyRegistration = myReader["allowearlyregistration"].ToString();
                sRegistrationStartDates = getRegistrationStartDates(sClassID.ToString());

                lcl_return = "<div class=\"registrationLine\">";
                lcl_return += "<table class=\"registrationLineTable\">";

                //-- Early Registration --
                if (sAllowEarlyRegistration.ToUpper() == "TRUE")
                {
                    sEarlyClassLabel = getEarlyClassLabel(sClassID.ToString());
                    sEarlyRegistrationDate = new DateTime();

                    if (myReader["earlyregistrationdate"].ToString() != null)
                    {
                        sEarlyRegistrationDate = Convert.ToDateTime(myReader["earlyregistrationdate"]);
                    }

                    lcl_return += "  <tr>";
                    lcl_return += "      <td><strong>Early Registration Starts: </strong></td>";
                    lcl_return += "      <td>" + sEarlyRegistrationDate.ToString("dddd, MMMM dd, yyyy") + "</td>";
                    lcl_return += "  </tr>";
                    lcl_return += "  <tr>";
                    lcl_return += "      <td colspan=\"2\">";
                    lcl_return += "          <table class=\"enrolledInClasses\">";
                    lcl_return += "            <tr valign=\"top\">";
                    lcl_return += "                <td class=\"enrolledInLabel\">For those enrolled in:</td>";
                    lcl_return += "                <td>" + sEarlyClassLabel + "</td>";
                    lcl_return += "            </tr>";
                    lcl_return += "          </table>";
                    lcl_return += "      </td>";
                    lcl_return += "  </tr>";
                }

                //-- Registration Start Dates --
                lcl_return += sRegistrationStartDates;

                //-- Registration End Dates --
                if (string.Format("{0:MM/dd/yyyy}", Convert.ToDateTime(myReader["registrationenddate"])) != "01/01/1900")
                {
                    lcl_return += "  <tr>";
                    lcl_return += "      <td><strong>Registration Ends: </strong></td>";
                    lcl_return += "      <td>" + string.Format("{0:dddd, MMMM dd, yyyy}", Convert.ToDateTime(myReader["registrationenddate"])) + "</td>";
                    lcl_return += "  </tr>";
                }

                lcl_return += "</table>";
                lcl_return += "</div>";
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getRegistrationStartDates(string iClassID)
    {
        //-------------------------------------------------------
        //This function is used to build the registration start
        //date list on class_list.aspx
        //-------------------------------------------------------
        string lcl_return = "";
        string sSQL = "";
        string sRegStartDateLabel = "";

        DateTime sRegistrationStartDate;

        Int32 sClassID = 0;

        if (iClassID != null)
        {
            try
            {
                sClassID = Convert.ToInt32(iClassID);
            }
            catch
            {
                sClassID = 0;
            }
        }

        sSQL = "SELECT registrationstartdate, ";
        sSQL += " pricetypename ";
        sSQL += " FROM egov_class_pricetype_price c, ";
        sSQL += " egov_price_types p ";
        sSQL += " WHERE classid = " + sClassID.ToString();
        sSQL += " AND c.pricetypeid = p.pricetypeid ";
        sSQL += " AND registrationstartdate IS NOT NULL ";
        sSQL += " ORDER BY displayorder ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                sRegStartDateLabel = "Registration Starts";
                sRegistrationStartDate = Convert.ToDateTime(myReader["registrationstartdate"]);

                if (myReader["pricetypename"].ToString().ToUpper() != "EVERYONE")
                {
                    sRegStartDateLabel = myReader["pricetypename"].ToString() + " " + sRegStartDateLabel;
                }

                lcl_return += "  <tr>";
                lcl_return += "      <td><strong>" + sRegStartDateLabel + ": </strong></td>";
                lcl_return += "      <td>" + sRegistrationStartDate.ToString("dddd, MMMM dd, yyyy") + "</td>";
                lcl_return += "  </tr>";
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getRegistrationStarts(Int32 iClassID)
    {
        //-------------------------------------------------------
        //This function is used to build the registration start
        //date list on class_details.aspx
        //-------------------------------------------------------
        string lcl_return = "";
        string sSQL = "";
        string sRegistrationStartDate = "";

        Boolean sPriceTypeNamesExist = false;

        sSQL = "SELECT registrationstartdate, ";
        sSQL += " pricetypename ";
        sSQL += " FROM egov_class_pricetype_price c, ";
        sSQL += " egov_price_types p ";
        sSQL += " WHERE classid = " + iClassID.ToString();
        sSQL += " AND c.pricetypeid = p.pricetypeid ";
        sSQL += " AND registrationstartdate is not null ";
        sSQL += " ORDER BY displayorder ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                sRegistrationStartDate = "&nbsp;";

                if (myReader["pricetypename"].ToString().Trim() != null)
                {
                    if (myReader["registrationstartdate"].ToString() != null)
                    {
                        sRegistrationStartDate = string.Format("{0:dddd, MMMM dd, yyyy}", Convert.ToDateTime(myReader["registrationstartdate"]));
                    }

                    //If atleast one PriceTypeName exists that is NOT EQUAL to "EVERYONE"
                    //then display the results in a nested table.  Otherwise, simply return
                    //the RegistrationStartDate.
                    if (myReader["pricetypename"].ToString().ToUpper() != "EVERYONE")
                    {
                        lcl_return += "<tr valign=\"top\">";
                        lcl_return +=     "<td class=\"classes_regstart_label\">" + myReader["pricetypename"].ToString().Trim() + ":";
                        lcl_return +=     "<td>" + sRegistrationStartDate + "</td>";
                        lcl_return += "</tr>";

                        sPriceTypeNamesExist = true;
                    }
                    else
                    {
                        lcl_return = sRegistrationStartDate;
                    }
                }
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        if (lcl_return != null && sPriceTypeNamesExist == true)
        {
            lcl_return = "<table class=\"classes_regstart_table\">" + lcl_return + "</table>";
        }

        return lcl_return;
    }

    public static string getEarlyClassLabel(string iClassID)
    {
        string lcl_return = "";
        string sSQL = "";

        Int32 sClassID = 0;

        if (iClassID != null)
        {
            try
            {
                sClassID = Convert.ToInt32(iClassID);
            }
            catch
            {
                sClassID = 0;
            }
        }

        sSQL = "SELECT c.classname, ";
        sSQL += " s.seasonname ";
        sSQL += " FROM egov_class c, ";
        sSQL += " egov_class_seasons s, ";
        sSQL += " egov_class_earlyregistrations e ";
        sSQL += " WHERE c.classseasonid = s.classseasonid ";
        sSQL += " AND e.earlyregistrationclassid = c.classid ";
        sSQL += " AND e.classid = " + sClassID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                lcl_return += "<div>";
                lcl_return += myReader["classname"].ToString();
                lcl_return += " (" + myReader["seasonname"].ToString() + ")";
                lcl_return += "</div>";
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static DateTime getEarliestRegistrationStart(Int32 iClassID)
    {
        string sSQL = "";

        DateTime lcl_return = DateTime.Now;

        sSQL  = "SELECT registrationstartdate, ";
        sSQL += " pricetypename ";
        sSQL += " FROM egov_class_pricetype_price c, ";
        sSQL +=      " egov_price_types p ";
        sSQL += " WHERE classid = " + iClassID.ToString();
        sSQL += " AND c.pricetypeid = p.pricetypeid ";
        sSQL += " AND registrationstartdate IS NOT NULL ";
        sSQL += " ORDER BY registrationstartdate ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToDateTime(myReader["registrationstartdate"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getInstructorName(string iNamePart, Int32 iInstructorID)
    {
        string lcl_return = "";
        string sSQL       = "";
        string sNamePart  = "LASTNAME";

        if (iNamePart != null)
        {
            sNamePart = iNamePart.ToUpper();
        }

        sSQL  = "SELECT isnull(firstname, '') as firstname, ";
        sSQL += " isnull(lastname, '') as lastname ";
        sSQL += " FROM egov_class_instructor ";
        sSQL += " WHERE instructorid = " + iInstructorID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if(sNamePart == "FULLNAME")
            {
                lcl_return = myReader["firstname]"].ToString() + ' ' + myReader["lastname"].ToString();
                lcl_return = lcl_return.Trim();
            }
            else if (sNamePart == "FIRSTNAME")
            {
                lcl_return = myReader["firstname"].ToString().Trim();
            }
            else
            {
                lcl_return = myReader["lastname"].ToString().Trim();
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Boolean classIsSeries(Int32 iClassID)
    {
        Boolean lcl_return = false;

        string sSQL = "";

        sSQL  = "SELECT isParent ";
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

            if (Convert.ToBoolean(myReader["isParent"]) == true)
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

    public static Int32 getSeriesEnrollment(Int32 iClassID)
    {
        Int32 lcl_return = 0;

        string sSQL = "";

        sSQL  = "SELECT (t.enrollmentsize + t.waitlistsize) as enrolled ";
        sSQL += " FROM egov_class_time t, ";
        sSQL +=      " egov_class c ";
        sSQL += " WHERE t.classid = c.classid ";
        sSQL += " AND c.parentclassid = " + iClassID.ToString(); 
        sSQL += " ORDER BY 1 desc ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToInt32(myReader["enrolled"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getGenderRestrictionText(Int32 iGenderRestrictionID)
    {
        string lcl_return = "";
        string sSQL       = "";

        sSQL  = "SELECT genderrestrictiontext ";
        sSQL += " FROM egov_class_genderrestrictions ";
        sSQL += " WHERE genderrestrictionid = " + iGenderRestrictionID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = myReader["genderrestrictiontext"].ToString();
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getGenderRestriction(Int32 iGenderRestrictionID)
    {
        string lcl_return = "N";
        string sSQL       = "";

        sSQL  = "SELECT genderrestriction ";
        sSQL += " FROM egov_class_genderrestrictions ";
        sSQL += " WHERE genderrestrictionid = " + iGenderRestrictionID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = myReader["genderrestriction"].ToString();
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Int32 getGenderNotRequiredID()
    {
        Int32 lcl_return = 1;

        string sSQL = "";

        sSQL  = "SELECT genderrestrictionid ";
        sSQL += " FROM egov_class_genderrestrictions ";
        sSQL += " WHERE isgendernotrequired = 1 ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if ((myReader["genderrestrictionid"].ToString().Trim() != "") && (myReader["genderrestrictionid"] != null))
            {
                lcl_return = Convert.ToInt32(myReader["genderrestrictionid"]);
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string buildBreadCrumbsCategories(Int32 iOrgID,
                                                    string iLocation,
                                                    Int32 iCategoryID,
                                                    Int32 iClassID)
    {
        Boolean sShowCrumb_classList    = false;
        Boolean sShowCrumb_classDetails = false;

        string lcl_return        = "";
        string sLocation         = "";
        string sFeatureName      = "";
        string sCategoryTitle    = "";
        string sClassName        = "";
        string sURL_classlist    = "";
        string sURL_classdetails = "";

        if (iLocation != null)
        {
            sLocation = iLocation.Trim().ToUpper();
        }

        //Determine which crumb(s) to show based on the "location" passed in.
        //The design behind this is to NOT show the bread crumb of the current screen, but
        //to show the breadcrumb(s) of screen(s) that the user has already visited.
        //It is to give them a way back to those screen WITHOUT including a link to
        //the screen they are current sitting on!
        if (sLocation == "CLASSLIST")
        {
            sShowCrumb_classList = true;
        }
        else if(sLocation == "CLASSDETAILS")
        {
            sShowCrumb_classList    = true;
            sShowCrumb_classDetails = true;
        }

        //-- Build the BreadCrumbs List --
        sFeatureName = common.getFeatureName(iOrgID.ToString(),
                                             "activities");
        
        lcl_return += "<div class=\"classes_breadcrumbs\">";
        lcl_return += "  <table>";
        lcl_return += "    <tr valign=\"top\">";
        lcl_return += "        <td><strong>Return to: </strong></td>";
        lcl_return += "        <td>";
        lcl_return += "<a href=\"class_categories.aspx\">" + sFeatureName + ": Categories</a>";

        // -- Class List -- //
        if (sShowCrumb_classList)
        {
            sCategoryTitle = getCategoryTitle(iCategoryID);
            sURL_classlist = sCategoryTitle + ": Class/Event List";

            if (sLocation != "CLASSLIST")
            {
                sURL_classlist = "<a href=\"class_list.aspx?categoryid=" + iCategoryID.ToString() + "\">" + sURL_classlist + "</a>";
            }

            lcl_return += " >> " + sURL_classlist;
        }

        // -- Class Details -- //
        if (sShowCrumb_classDetails)
        {
            sClassName        = getClassName(iClassID);
            sURL_classdetails = sClassName + ": Details";

            if (sLocation != "CLASSDETAILS")
            {
                sURL_classdetails = "<a href=\"class_details.aspx?classid=" + iClassID.ToString() + "&categoryid=" + iCategoryID.ToString() + "\">" + sURL_classdetails + "</a>";
                sURL_classdetails += " >> <a href=\"#\">" + sURL_classdetails + "</a>";
            }

            lcl_return += " >> " + sURL_classdetails;
        }

        lcl_return += "        </td>";
        lcl_return += "    </tr>";
        lcl_return += "  </table>";
        lcl_return += "</div>";

        return lcl_return;
    }

    public static string getCategoryTitle(Int32 iCategoryID)
    {
        string lcl_return = "";
        string sSQL       = "";

        sSQL  = "SELECT categorytitle ";
        sSQL += " FROM egov_class_categories ";
        sSQL += " WHERE categoryid = " + iCategoryID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = myReader["categorytitle"].ToString().Trim();
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
        
        return lcl_return;
    }

    public static string getClassName(Int32 iClassID)
    {
        string lcl_return = "";
        string sSQL       = "";

        sSQL  = "SELECT classname ";
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

            lcl_return = myReader["classname"].ToString().Trim();
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Boolean isClassFilled(Int32 iClassID)
    {
        //The difference between the "classIsFilled" function and the "isClassFilled" function is that
        //the "classIsFilled" function checks to see if the class is filled, but the waitlist has NOT 
        // reached the max waitlistsize.  We want to allow them to get in and register for the waitlist.
        //The "isClassFilled" function purpose is to block the actual purchasing of a class, even if the enrollment
        //size is less than the max allowed in the class IF there is at least one person on the waitlist.  This gives
        //the organization to give the person(s) on the waitlist a chance to purchase the class before opening it to
        //the public.

        Boolean lcl_return = true;

        string sSQL = "";

        sSQL  = "SELECT timeid, ";
        sSQL += " activityno, ";
        sSQL += " ISNULL([max],0) AS max, ";
        sSQL += " ISNULL(enrollmentsize,0) as enrollmentsize, ";
        sSQL += " ISNULL(waitlistsize,0) as waitlistsize ";
        sSQL += " FROM egov_class_time ";
        sSQL += " WHERE classid = " + iClassID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        while (myReader.Read())
        {
            //We assume that the class is filled and look for any availability in all activities offered.
            //If we find one, then filled is FALSE and we are done.
            if (Convert.ToInt32(myReader["waitlistsize"]) == Convert.ToInt32(0) && Convert.ToInt32(myReader["enrollmentsize"]) < Convert.ToInt32(myReader["max"]))
            {
                lcl_return = false;
                break;
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;

    }

    public static Boolean classIsFilled(Int32 iClassID)
    {
        //The difference between the "classIsFilled" function and the "isClassFilled" function is that
        //the "classIsFilled" function checks to see if the class is filled, but the waitlist has NOT 
        // reached the max waitlistsize.  We want to allow them to get in and register for the waitlist.
        //The "isClassFilled" function purpose is to block the actual purchasing of a class, even if the enrollment
        //size is less than the max allowed in the class IF there is at least one person on the waitlist.  This gives
        //the organization to give the person(s) on the waitlist a chance to purchase the class before opening it to
        //the public.
        Boolean lcl_return = false;

        string sSQL = "";

        sSQL  = "SELECT SUM(enrollmentsize) as enrollmentsize, ";
        sSQL += " sum(isnull(max,9999)) as max, ";
        sSQL += " sum(waitlistsize) as waitlistsize, ";
        sSQL += " sum(isnull(waitlistmax,9999)) as waitlistmax ";
        sSQL += " FROM egov_class_time ";
        sSQL += " WHERE classid = " + iClassID.ToString();
        sSQL += " GROUP BY classid ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (Convert.ToInt32(myReader["enrollmentsize"]) >= Convert.ToInt32(myReader["max"]) &&
                Convert.ToInt32(myReader["waitlistsize"]) >= Convert.ToInt32(myReader["waitlistmax"]))
            {
                lcl_return = true;
            }
            else
            {
                lcl_return = false;
            }
        }
        else
        {
            lcl_return = true;  //something is wrong
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Boolean classHasWaivers(Int32 iClassID)
    {
        Boolean lcl_return = false;

        string sSQL = "";

        sSQL  = "SELECT count(rowid) AS hits ";
        sSQL += " FROM egov_class_to_waivers ";
        sSQL += " WHERE classid = " + iClassID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (Convert.ToInt32(myReader["hits"]) > Convert.ToInt32(0))
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

    public static string showWaiverList(Int32 iOrgID,
                                        Int32 iClassID,
                                        Boolean iShowWaiverText,
                                        Boolean iShowWaiverName,
                                        Boolean iShowWaiverDesc,
                                        Boolean iShowWaiverLink)
    {
        string lcl_return         = "";
        string sSQL               = "";
        string sWaiverText        = "";
        string sWaiverName        = "";
        string sWaiverDesc        = "";
        string sDisplayWaiverName = "";
        string sDisplayWaiverDesc = "";

        sSQL  = "SELECT cw.waiverid, ";
        sSQL += " cw.waivername, ";
        sSQL += " cw.waiverdescription, ";
        sSQL += " cw.waiverbody, ";
        sSQL += " cw.waiverurl, ";
        sSQL += " cw.isRequired, ";
        sSQL += " cw.orgid, ";
        sSQL += " cw.waivertype ";
        sSQL += " FROM egov_class_waivers cw ";
        sSQL +=      " INNER JOIN egov_class_to_waivers ctw ON cw.waiverid = ctw.waiverid ";
        sSQL += " WHERE cw.orgid = " + iOrgID.ToString();
        sSQL += " AND ctw.classid = " + iClassID.ToString();
        sSQL += " ORDER BY cw.waivername ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            lcl_return  = "<div class=\"classdetails_waiver_adobe\">";
            lcl_return += "<a href=\"http://www.adobe.com/products/acrobat/readstep2.html\" target=\"_blank\" title=\"Get Adobe Acrobat Reader Plug-in Here\">";
            lcl_return += "<img src=\"../images/adreader.gif\" hspace=\"10\" />Get Adobe Reader</a>";
            lcl_return += "</div>";

            if (iShowWaiverText)
            {
                sWaiverText = common.getOrgInfo(Convert.ToString(iOrgID), "orgwaivertext");

                lcl_return += "<div class=\"classdetails_waiverText\">" + sWaiverText + "</div>";
            }

            while (myReader.Read())
            {
                sDisplayWaiverName = "";
                sDisplayWaiverDesc = "";

                if (iShowWaiverName)
                {
                    if (Convert.ToString(myReader["waivername"]) != "")
                    {
                        sWaiverName = Convert.ToString(myReader["waivername"]).Trim();
                        sWaiverName = common.decodeUTFString(sWaiverName);

                        sDisplayWaiverName = "<legend>" + sWaiverName + "</legend>";
                    }
                }

                if (iShowWaiverDesc)
                {
                    if (Convert.ToString(myReader["waiverdescription"]) != "")
                    {
                        sWaiverDesc = Convert.ToString(myReader["waiverdescription"]).Trim();
                        sWaiverDesc = common.decodeUTFString(sWaiverDesc);

                        sDisplayWaiverDesc = "<div>" + sWaiverDesc + "</div>";
                    }
                }

                if (iShowWaiverLink)
                {
                    lcl_return += "<fieldset class=\"class_signup_fieldset\">";
                    lcl_return += sDisplayWaiverName + sDisplayWaiverDesc;
                    lcl_return += "  <div class=\"classdetails_waiver\">";
                    lcl_return += "    <a href=\"" + Convert.ToString(myReader["waiverurl"]) + "\" target=\"_NEW\" class=\"waiverlink\">Click here to download " + myReader["waivername"].ToString().ToUpper() + " waiver.</a>";
                    lcl_return += "  </div>";
                    lcl_return += "</fieldset>";


                }
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Boolean isRegistrationStarted(Int32 iClassID)
    {
        Boolean lcl_return = false;

        string sRegistrationStartDate = "";
        string sSQL                   = "";

        sSQL  = "SELECT registrationstartdate ";
        sSQL += " FROM egov_class_pricetype_price ";
        sSQL += " WHERE classid = " + iClassID.ToString();
        sSQL += " AND registrationstartdate IS NOT NULL ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            sRegistrationStartDate = Convert.ToString(myReader["registrationstartdate"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        if(sRegistrationStartDate != "" && sRegistrationStartDate != null) {
            //if (Convert.ToDateTime(string.Format("{0:MM/dd/yyyy}", sRegistrationStartDate)) < Convert.ToDateTime(string.Format("{0:MM/dd/yyyy}", DateTime.Now.ToString())))
            if (Convert.ToDateTime(string.Format("{0:MM/dd/yyyy}", sRegistrationStartDate)) < Convert.ToDateTime(string.Format("{0:MM/dd/yyyy}", DateTime.Now)))
           {
               lcl_return = true;
           }
        }

        return lcl_return;
    }

    public static Boolean classIsRegattaTeamSignup(Int32 iClassID)
    {
        Boolean lcl_return = false;

        string sSQL = "";

        sSQL  = "SELECT t.isteamsignup ";
        sSQL += " FROM egov_regattasignuptype t, ";
        sSQL +=      " egov_class c ";
        sSQL += " WHERE c.classid = " + iClassID.ToString();
        sSQL += " AND c.regattasignuptypeid = t.regattasignuptypeid ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToBoolean(myReader["isteamsignup"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Boolean UserIsMissingAddress(Int32 iUserID)
    {
        Boolean lcl_return = true;

        string sSQL = "";

        sSQL  = "SELECT isnull(useraddress,'') as useraddress, ";
        sSQL += " isnull(usercity,'') as usercity, ";
        sSQL += " isnull(userstate,'') as userstate, ";
        sSQL += " isnull(userzip,'') as userzip ";
        sSQL += " FROM egov_users ";
        sSQL += " WHERE userid = " + iUserID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (myReader["useraddress"].ToString() != "" &&
                myReader["usercity"].ToString()    != "" &&
                myReader["userstate"].ToString()   != "" &&
                myReader["userzip"].ToString()     != "")
            {
                lcl_return = false;
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Boolean getMemberInformation(Int32 iUserID,
                                               Int32 iMembershipID,
                                               Int32 lcl_membercount,
                                               out Int32 iMemberCount)
    {
        Boolean lcl_return = false;

        Int32 iCount       = 0;
        Int32 sMemberCount = lcl_membercount;

        string sSQL            = "";
        string sMembershipType = "";

        sSQL  = "SELECT familymemberid ";
        sSQL += " FROM egov_familymembers ";
        sSQL += " WHERE belongstouserid = " + iUserID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                sMembershipType = determineMembership(Convert.ToInt32(myReader["familymemberid"]),
                                                      iUserID,
                                                      iMembershipID);

                //Count the number of family members that are "Members (M)".
                if (sMembershipType == "M")
                {
                    sMemberCount = sMemberCount + 1;
                }

                iCount = iCount + 1;
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        if (sMemberCount == iCount)
        {
            lcl_return = true;
        }

        iMemberCount = sMemberCount;

        return lcl_return;
    }

    public static string determineMembership(Int32 iFamilyMemberID,
                                             Int32 iUserID,
                                             Int32 iMembershipID)
    {
        DateTime sExpirationDate = DateTime.Now;

        string lcl_return = "O";  //Not a member
        string sSQL       = "";

        sSQL  = "SELECT paymentdate, ";
        sSQL += " mp.is_seasonal, ";
        sSQL += " mp.period_interval, ";
        sSQL += " isnull(mp.period_qty,0) as period_qty ";
        sSQL += " FROM egov_poolpasspurchases p, ";
        sSQL +=      " egov_poolpassmembers m, ";
        sSQL +=      " egov_poolpassrates r, ";
        sSQL +=      " egov_membership_periods mp ";
        sSQL += " WHERE m.poolpassid = p.poolpassid ";
        sSQL += " AND p.rateid = r.rateid ";
        sSQL += " AND r.periodid = mp.periodid ";
        sSQL += " AND (p.paymentresult = 'Paid' OR p.paymentresult = 'APPROVED') ";
        sSQL += " AND m.familymemberid = " + iFamilyMemberID.ToString();
        sSQL += " AND p.userid = " + iUserID.ToString();
        sSQL += " AND r.membershipid = " + iMembershipID.ToString();
        sSQL += " ORDER BY paymentdate desc ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (Convert.ToBoolean(myReader["is_seasonal"]))
            {
                //if (string.Format("{0:yyyy}", Convert.ToDateTime(myReader["paymentdate"])) == string.Format("{0:yyyy}", DateTime.Now.ToString()))
                if (string.Format("{0:yyyy}", Convert.ToDateTime(myReader["paymentdate"])) == string.Format("{0:yyyy}", DateTime.Now))
                {
                    lcl_return = "M";  //A member
                }
            }
            else
            {
                sExpirationDate = getExpirationDate(Convert.ToString(myReader["period_interval"]),
                                                    Convert.ToInt32(myReader["period_qty"]),
                                                    Convert.ToDateTime(myReader["paymentdate"]));

                //if (Convert.ToDateTime(string.Format("{0:MM/dd/yyyy}", DateTime.Now.ToString())) >= sExpirationDate)
                if (Convert.ToDateTime(string.Format("{0:MM/dd/yyyy}", DateTime.Now)) >= sExpirationDate)
                {
                    lcl_return = "O";  //Expired Membership
                } else
                {
                    lcl_return = "M";  //Active membership
                }
            }
         }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static DateTime getExpirationDate(string iPeriodInterval,
                                             Int32 iPeriodQty,
                                             DateTime iPaymentDate)
    {
        DateTime lcl_return = new DateTime();

        string sPeriodInterval = "";

        if (iPaymentDate != null)
        {
            lcl_return = iPaymentDate;
        }

        if (iPeriodInterval != "")
        {
            sPeriodInterval = iPeriodInterval.ToUpper();
        }

        if (sPeriodInterval == "D")
        {
            lcl_return = lcl_return.AddDays(iPeriodQty);
        }
        else if (sPeriodInterval == "WW")
        {
            //Since there isn't an "AddWeeks" method, we will use the "AddDays".
            //Simply take the periodquantity passed in (in weeks) and muliple it by 7 days.
            lcl_return = lcl_return.AddDays((iPeriodQty * 7));
        }
        else if (sPeriodInterval == "M")
        {
            lcl_return = lcl_return.AddMonths(iPeriodQty);
        }
        else if (sPeriodInterval == "YYYY")
        {
            lcl_return = lcl_return.AddYears(iPeriodQty);
        }

        return lcl_return;
    }

    public static string getRegistrationEndDate(Int32 iClassID)
    {
        string lcl_return = "";
        string sSQL       = "";

        sSQL  = "SELECT isnull(registrationenddate,'') as registrationenddate ";
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

            lcl_return = myReader["registrationenddate"].ToString();
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Boolean registrationStarted(Int32 iClassID,
                                              Int32 iUserID,
                                              out string sRegStartDate,
                                              out string sPriceType)
    {
        Boolean lcl_return = false;

        DateTime sLatestRegStartDate = getLatestRegistrationStartDate(iClassID);

        string sSQL          = "";
        string sResidentType = "";
        string sMemberType   = "";

        sRegStartDate = "";
        sPriceType    = "";

	int sTimeZone = Int32.Parse(common.getOrgInfo(common.getOrgId(),"orgid, t.GMTOffset"));

	DateTime currDateTime = DateTime.Now.AddHours(5).AddHours(sTimeZone);


        if (sLatestRegStartDate <= currDateTime)
        {
            lcl_return = true;
        }
        else
        {
            sSQL  = "SELECT isnull(c.registrationstartdate,'') as registrationstartdate, ";
            sSQL += " p.pricetype, ";
            sSQL += " p.pricetypename, ";
            sSQL += " p.checkresidency, ";
            sSQL += " p.isresident, ";
            sSQL += " p.checkmembership, ";
            sSQL += " p.ismember ";
            sSQL += " FROM egov_class_pricetype_price c, ";
            sSQL +=      " egov_price_types p ";
            sSQL += " WHERE c.pricetypeid = p.pricetypeid ";
            sSQL += " AND p.isactiveforclasses = 1 ";
            sSQL += " AND c.registrationstartdate is not null ";
            sSQL += " AND c.classid = " + iClassID.ToString();

            SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
            sqlConn.Open();

            SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
            SqlDataReader myReader;
            myReader = myCommand.ExecuteReader();

            if (myReader.HasRows)
            {
                while (myReader.Read())
                {
                    if (myReader["pricetype"].ToString() == "E")
                    {
                        //For EVERYONE pricing just check the date.
                        if (myReader["registrationstartdate"].ToString() != "")
                        {
                            if (Convert.ToDateTime(myReader["registrationstartdate"]) <= currDateTime)
                            {
                                lcl_return = true;
                            }
                            else
                            {
                                sRegStartDate = Convert.ToString(Convert.ToDateTime(myReader["registrationstartdate"]));
                                sPriceType    = Convert.ToString(myReader["pricetypename"]);
                            }

                            break;
                        }
                    }
                    else
                    {
                        if (Convert.ToBoolean(myReader["checkresidency"]))
                        {
                            sResidentType = getUserResidentType(iUserID);

                            if (sResidentType == myReader["pricetype"].ToString())
                            {
                                //Matches R or N
                                if (Convert.ToDateTime(myReader["registrationstartdate"]) <= currDateTime)
                                {
                                    lcl_return = true;
                                }
                                else
                                {
                                    sRegStartDate = Convert.ToString(Convert.ToDateTime(myReader["registrationstartdate"]));
                                    sPriceType    = Convert.ToString(myReader["pricetypename"]) + "s";
                                }

                                break;
                            }
                        }
                        else
                        {
                            //Check Membership
                            sMemberType = getMembershipCode(iClassID,
                                                            iUserID);

                            if (sMemberType == myReader["pricetype"].ToString())
                            {
                                //Matches "M" or "O"
                                if (Convert.ToDateTime(myReader["registrationstartdate"]) <= currDateTime)
                                {
                                    lcl_return = true;
                                }
                                else
                                {
                                    sRegStartDate = Convert.ToString(Convert.ToDateTime(myReader["registrationstartdate"]));
                                    sPriceType    = Convert.ToString(myReader["pricetypename"]) + "s";
                                }

                                break;
                            }
                        }
                    }
                }
            }

            myReader.Close();
            sqlConn.Close();
            myReader.Dispose();
            sqlConn.Dispose();

            //This only set member pricing not anything else and the user is not a member
            if (! lcl_return)
            {
                if (sMemberType == "O" && sPriceType == "" && sRegStartDate == "")
                {
                    sPriceType = "Non Members";
                }
            }
        }

        return lcl_return;
    }

    public static DateTime getLatestRegistrationStartDate(Int32 iClassID)
    {
        DateTime lcl_return = DateTime.Now;

        string sSQL = "";

        sSQL  = "SELECT isnull(c.registrationstartdate,'') as registrationstartdate ";
        sSQL += " FROM egov_class_pricetype_price c, ";
        sSQL +=      " egov_price_types p ";
        sSQL += " WHERE c.pricetypeid = p.pricetypeid ";
        sSQL += " AND registrationstartdate is not null ";
        sSQL += " AND c.classid = " + iClassID.ToString();
        sSQL += " ORDER BY c.registrationstartdate desc";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (myReader["registrationstartdate"].ToString() != "")
            {
                lcl_return = Convert.ToDateTime(myReader["registrationstartdate"]);
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getUserResidentType(Int32 iUserID)
    {
        string lcl_return = "";
        string sSQL       = "";

        if (iUserID.ToString() != "")
        {
            sSQL  = "SELECT isnull(residenttype,'N') as residenttype ";
            sSQL += " FROM egov_users ";
            sSQL += " WHERE userid = " + iUserID.ToString();

            SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
            sqlConn.Open();

            SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
            SqlDataReader myReader;
            myReader = myCommand.ExecuteReader();

            if (myReader.HasRows)
            {
                myReader.Read();

                lcl_return = myReader["residenttype"].ToString();
            }

            myReader.Close();
            sqlConn.Close();
            myReader.Dispose();
            sqlConn.Dispose();  
        }

        if (lcl_return == null || lcl_return == "")
        {
            lcl_return = "N";
        }

        return lcl_return;
    }

    public static string getResidentTypeByAddress(Int32 iUserID,
                                                  Int32 iOrgID)
    {
        string lcl_return = "N";
        string sSQL       = "";

        sSQL  = "SELECT count(r.residentaddressid) as hits ";
        sSQL += " FROM egov_residentaddresses r, ";
        sSQL +=      " egov_users u ";
        sSQL += " WHERE r.orgid = u.orgid ";
        sSQL += " AND r.residentstreetnumber + ' ' + r.residentstreetname = u.useraddress ";
        sSQL += " AND r.residenttype = 'R' ";
        sSQL += " AND r.orgid = " + iOrgID.ToString();
        sSQL += " AND u.userid = " + iUserID.ToString();

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
                lcl_return = "R";
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();  

        return lcl_return;
    }

    public static string getMembershipCode(Int32 iClassID,
                                           Int32 iUserID)
    {
        Int32 iFamilyMemberID = getCitizenFamilyID(iUserID);
        Int32 iMembershipID   = getClassMembershipID(iClassID);

        string lcl_return = determineMembership(iFamilyMemberID,
                                                iUserID,
                                                iMembershipID);

        return lcl_return;
    }

    public static Int32 getCitizenFamilyID(Int32 iUserID)
    {
        Int32 lcl_return = 0;

        string sSQL = "";

        sSQL  = "SELECT familymemberid ";
        sSQL += " FROM egov_familymembers ";
        sSQL += " WHERE userid = " + iUserID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToInt32(myReader["familymemberid"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Int32 getAttendeeUserID(Int32 iFamilyMemberID)
    {
        Int32 lcl_return = 0;

        string sSQL = "";

        sSQL  = "SELECT userid ";
        sSQL += " FROM egov_familymembers ";
        sSQL += " WHERE familymemberid = " + iFamilyMemberID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToInt32(myReader["userid"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Int32 getClassMembershipID(Int32 iClassID)
    {
        Int32 lcl_return = 0;

        string sSQL = "";

        sSQL  = "SELECT isnull(membershipid,0) as membershipid ";
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

            lcl_return = Convert.ToInt32(myReader["membershipid"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Boolean userCanRegisterEarly(Int32 iClassID,
                                               Int32 iUserID)
    {
        Boolean lcl_return = false;

        string sSQL = "";

        sSQL  = "SELECT count(p.paymentid) as hits ";
        sSQL += " FROM egov_class_payment p, ";
        sSQL +=      " egov_class_list l, ";
        sSQL +=      " egov_class_earlyregistrations e ";
        sSQL += " WHERE e.earlyregistrationclassid = l.classid ";
        sSQL += " AND p.paymentid = l.paymentid ";
        sSQL += " AND l.status = 'ACTIVE' ";
        sSQL += " AND e.classid = " + iClassID.ToString();
        sSQL += " AND l.userid = " + iUserID.ToString();

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

    public static Boolean userNotBlocked(Boolean iOrgHasFeature_registrationBlocking,
                                         Int32 iUserID)
    {
        Boolean lcl_return = true;

        string sSQL = "";

        if (iOrgHasFeature_registrationBlocking)
        {
            sSQL = "SELECT registrationblocked ";
            sSQL += " FROM egov_users ";
            sSQL += " WHERE userid = " + iUserID.ToString();

            SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
            sqlConn.Open();

            SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
            SqlDataReader myReader;
            myReader = myCommand.ExecuteReader();

            if (myReader.HasRows)
            {
                myReader.Read();

                lcl_return = !Convert.ToBoolean(myReader["registrationblocked"]);  //NOTE: the "!" makes it return the opposite value
            }

            myReader.Close();
            sqlConn.Close();
            myReader.Dispose();
            sqlConn.Dispose();
        }

        return lcl_return;
    }

    public static Boolean checkMemberRequirement(Int32 iClassID, Int32 iOrgID, Int32 iMemberCount, string iUserType, Boolean iAllMembers, out string sPriceType)
    {
        Boolean lcl_return = false;

        Int32 iCount = 0;

        string sSQL = "";
        
        sPriceType = "";

        sSQL  = "SELECT pricetype ";
        sSQL += " FROM egov_price_types t, ";
        sSQL +=      " egov_class_pricetype_price p ";
        sSQL += " WHERE t.pricetypeid = p.pricetypeid ";
        sSQL += " AND orgid = " + iOrgID.ToString();
        sSQL += " AND classid = " + iClassID.ToString();
        sSQL += " ORDER BY p.pricetypeid ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                //If they fail it is probably from having only one restrictive pricetype like 'M' or 'R'
                sPriceType = myReader["pricetype"].ToString();

                //For Resident/Non-Resident Choices ONLY
                if (sPriceType == "R" || sPriceType == "N")
                {
                    if (sPriceType == iUserType)
                    {
                        iCount = iCount + 1;
                    }
                }
                else
                {
                    //If it is for members and they have some members
                    if (sPriceType == "M")
                    {
                        if (iMemberCount > 0)
                        {
                            iCount = iCount + 1;
                        }
                    }
                    else
                    {
                        //For non-members, only if some are not members
                        if (sPriceType == "O")
                        {
                            if (!iAllMembers)
                            {
                                iCount = iCount + 1;
                            }
                        }
                        else
                        {
                            //That leaves "E" for Everyone
                            iCount = iCount + 1;
                        }
                    }
                }
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        if (iCount > 0)
        {
            lcl_return = true;
        }

        return lcl_return;
    }

    public static Boolean ageRequirementsMet(Int32 iUserID, Int32 iClassID)
    {
        Boolean lcl_return                        = false;
        Boolean sHasWholeYearPrecision_min        = true;
        Boolean sHasWholeYearPrecision_max        = true;
        Boolean sIsAgeToMin_greaterOrEqual_MinAge = false;
        Boolean sIsAgeToMax_lessOrEqual_MaxAge    = false;

        double sMinAge = 0.0;
        double sMaxAge = 0.0;
        double sAge    = 22.0;

        Int32 sMinAgePrecisionID = 1;
        Int32 sMaxAgePrecisionID = 1;

        string sSQL            = "";
        //string sBirthDate      = "";
        string sAgeCompareDate = Convert.ToString(DateTime.Now);

        getClassValues(iClassID,
                       out sAgeCompareDate,
                       out sMinAge,
                       out sMaxAge,
                       out sMinAgePrecisionID,
                       out sMaxAgePrecisionID);

        sSQL  = "SELECT familymemberid, ";
        sSQL += " firstname, ";
        sSQL += " lastname, ";
        sSQL += " isnull(birthdate,'') as birthdate ";
        sSQL += " FROM egov_familymembers ";
        sSQL += " WHERE belongstouserid = " + iUserID.ToString();
        sSQL += " AND isdeleted = 0 ";
        sSQL += " ORDER BY birthdate ";
        
        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                
                //sBirthDate = getBirthDate(Convert.ToInt32(myReader["familymemberid"]));
                if (Convert.ToString(myReader["birthdate"]) != "")
                {
                    if (string.Format("{0:MM/dd/yyyy}", Convert.ToDateTime(myReader["birthdate"])) != "01/01/1900")
                    {
                        sAge = getAgeOnDate(Convert.ToDateTime(myReader["birthdate"]),
                                            Convert.ToDateTime(sAgeCompareDate));
                    }
                }
                // we want the age to have only one decimal place so we are comparing it to a one decimal age restriction. Otherwise some ages will fall into an age gap and not be eligible for a class. 
                // Example: a child who is 8.91 being too old for a class that has a max age of 8.9. The intention is for any child under 9 to be valid.
                sAge = Math.Truncate( sAge * 10 ) / 10;

                sHasWholeYearPrecision_min = hasWholeYearPrecision(sMinAgePrecisionID);
                sHasWholeYearPrecision_max = hasWholeYearPrecision(sMaxAgePrecisionID);

                sIsAgeToMin_greaterOrEqual_MinAge = false;
                sIsAgeToMax_lessOrEqual_MaxAge    = false;

                if (sHasWholeYearPrecision_min)
                {
                    if(Convert.ToInt32(Math.Truncate(sAge)) >= Convert.ToInt32(sMinAge))
                    {
                        sIsAgeToMin_greaterOrEqual_MinAge = true;
                    }
                }
                else
                {
                    if (sAge >= sMinAge)
                    {
                        sIsAgeToMin_greaterOrEqual_MinAge = true;
                    }
                }

                if (sHasWholeYearPrecision_max)
                {
                    if (Convert.ToInt32(Math.Truncate(sAge)) <= Convert.ToInt32(sMaxAge))
                    {
                        sIsAgeToMax_lessOrEqual_MaxAge = true;
                    }
                }
                else
                {
                    if (sMaxAge > 0)
                    {
                        if (sAge <= sMaxAge)
                        {
                            sIsAgeToMax_lessOrEqual_MaxAge = true;
                        }
                    }
                    else
                    {
                        if (sAge > 0)
                        {
                            sIsAgeToMax_lessOrEqual_MaxAge = true;
                        }
                    }
                }

                //Determine, based on the precision level above if the age is:
                //  1. The age is GREATER THAN or EQUAL TO the minimum age limit.
                //  2. The age is GREATER THAN zero (0).
                //  3. The age is LESS THAN or EQUAL TO the max age limit.
                //If ALL criteria is met then the age has met all requirements.
                if (sIsAgeToMin_greaterOrEqual_MinAge)
                {
                    if (sMaxAge > 0)
                    {
                        if (sIsAgeToMax_lessOrEqual_MaxAge)
                        {
                            lcl_return = true;
                            break;
                        }
                    }
                    else
                    {
                        lcl_return = true;
                        break;
                    }
                }
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static void getClassValues(Int32 iClassID,
                               out string sAgeCompareDate,
                               out double sMinAge,
                               out double sMaxAge,
                               out Int32 sMinAgePrecisionID,
                               out Int32 sMaxAgePrecisionID)
    {
        string sSQL = "";

        sAgeCompareDate    = "";
        sMinAge            = 0;
        sMaxAge            = 0;
        sMinAgePrecisionID = 0;
        sMaxAgePrecisionID = 0;

        sSQL  = "SELECT isnull(agecomparedate, '" + DateTime.Now.ToString() + "') as agecomparedate, ";
        sSQL += " isnull(minage,0.0) as minage, ";
        sSQL += " isnull(maxage,0.0) as maxage, ";
        sSQL += " isnull(minageprecisionid,0) as minageprecisionid, ";
        sSQL += " isnull(maxageprecisionid,0) as maxageprecisionid ";
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

            if (myReader["agecomparedate"].ToString() != "")
            {
                sAgeCompareDate    = Convert.ToString(myReader["agecomparedate"]);
                sMinAge            = Convert.ToDouble(myReader["minage"]);
                sMaxAge            = Convert.ToDouble(myReader["maxage"]);
                sMinAgePrecisionID = Convert.ToInt32(myReader["minageprecisionid"]);
                sMaxAgePrecisionID = Convert.ToInt32(myReader["maxageprecisionid"]);
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
    }

    public static string getFamilyMemberName(Int32 iFamilyMemberID)
    {
        string lcl_return = "";
        string sSQL       = "";

        sSQL  = "SELECT firstname, ";
        sSQL += " lastname ";
        sSQL += " FROM egov_familymembers ";
        sSQL += " WHERE familymemberid = " + iFamilyMemberID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return  = Convert.ToString(myReader["firstname"]);
            lcl_return += " ";
            lcl_return += Convert.ToString(myReader["lastname"]);
            lcl_return = lcl_return.Trim();
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getBirthDate(Int32 iFamilyMemberID)
    {
        DateTime sBirthDate = DateTime.Now.AddYears(-22);

        string lcl_return = "";
        string sSQL       = "";

        sSQL  = "SELECT birthdate ";
        sSQL += " FROM egov_familymembers ";
        sSQL += " WHERE familymemberid = " + iFamilyMemberID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (Convert.ToString(myReader["birthdate"]) != "")
            {
                sBirthDate = Convert.ToDateTime(myReader["birthdate"]);
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        lcl_return = Convert.ToString(sBirthDate);

        return lcl_return;
    }

    public static double getAgeOnDate(DateTime iBirthDate, DateTime iCompareDate)
    {
        double lcl_return = 0;
        
        if (iCompareDate.Month == iBirthDate.Month && iCompareDate.Day == iBirthDate.Day)
        {
            // if today is your birthday, then just subtract the years
            int ageInYears = iCompareDate.Year - iBirthDate.Year;
            lcl_return = double.Parse( ageInYears.ToString() );
        }
        else 
        {
            System.TimeSpan lcl_diff = iCompareDate.Subtract(iBirthDate);
            lcl_return = lcl_diff.Days / 365.25;
        }

        return lcl_return;
    }
    
    public static Int32 getWholeYearAgeOnDate(DateTime iBirthDate, DateTime iCompareDate)
    {
        Int32 age = 0;
        
        age = iCompareDate.Year - iBirthDate.Year;
        if (iCompareDate.Month < iBirthDate.Month || (iCompareDate.Month == iBirthDate.Month && iCompareDate.Day < iBirthDate.Day)) age--;
        
        return age;
    }

    public static Boolean hasWholeYearPrecision(Int32 iAgePrecisionID)
    {
        Boolean lcl_return = false;

        string sSQL = "";

        sSQL  = "SELECT iswholeyear ";
        sSQL += " FROM egov_class_ageprecisions ";
        sSQL += " WHERE precisionid = " + iAgePrecisionID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (Convert.ToBoolean(myReader["iswholeyear"]))
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

    public static Boolean showFamilyMembers(Int32 iUserID, Int32 iOrgID, Int32 sMemberCount, double iMinAge, double iMaxAge, Int32 iMembershipID, string iAgeCompareDate, 
        Int32 iMinAgePrecisionID, Int32 iMaxAgePrecisionID, string iGenderRestriction, out Int32 iMemberCount, out Int32 iSelectedFamilyMemberID, out string sDropDownList_familyMembers)
    {
        Boolean lcl_return                        = false;
        Boolean sHasWholeYearPrecision_min        = true;
        Boolean sHasWholeYearPrecision_max        = true;
        Boolean sIsAgeToMin_greaterOrEqual_MinAge = false;
        Boolean sIsAgeToMax_lessOrEqual_MaxAge    = false;
        Boolean sOkToAdd                          = false;
        sDropDownList_familyMembers = "";

        double sMinAge = iMinAge;
        //sDropDownList_familyMembers += "<!-- min age: " + Convert.ToString( sMinAge ) + " -->";
        double sMaxAge = iMaxAge;
        //sDropDownList_familyMembers += "<!-- max age: " + Convert.ToString( sMaxAge ) + " -->";
        double sAge    = 0;

        Int32 sMinAgePrecisionID = iMinAgePrecisionID;
        Int32 sMaxAgePrecisionID = iMaxAgePrecisionID;
        Int32 sCount             = 0;
        Int32 sChildsAge         = 0;

        string sSQL                    = "";
        string sMember                 = "O";
        string sFamilyMemberGender     = "N";
        //string sBirthDate              = "";
        string sAgeCompareDate         = iAgeCompareDate;
        string sFamilyMemberSelected   = "";
        string sFamilyMemberOptionText = "";

        iMemberCount                = sMemberCount;
        iSelectedFamilyMemberID     = 0;
        

        sSQL  = "SELECT familymemberid, ";
        sSQL += " firstname, ";
        sSQL += " lastname, ";
        sSQL += " relationship, ";
        sSQL += " isnull(birthdate,'') as birthdate, ";
        sSQL += " userid ";
        sSQL += " FROM egov_familymembers ";
        sSQL += " WHERE isdeleted = 0 ";
        sSQL += " AND belongstouserid = " + iUserID.ToString();
        sSQL += " ORDER BY birthdate ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            sDropDownList_familyMembers += "<div id=\"class_signup_selectfamilymembers\">";
            sDropDownList_familyMembers += "  <select name=\"familymemberid\" id=\"familymemberid\">";

            while (myReader.Read())
            {
                if (Convert.ToString(myReader["birthdate"]) == "" || string.Format("{0:MM/dd/yyyy}", Convert.ToDateTime(myReader["birthdate"])) == "01/01/1900")
                {
                    sAge = 22.0;
                }
                else
                {
                    if (iAgeCompareDate == "")
                    {
                        sAgeCompareDate = Convert.ToString(DateTime.Now);
                    }

                    //sBirthDate = classes.getBirthDate(Convert.ToInt32(myReader["familymemberid"]));
                    sAge = classes.getAgeOnDate(Convert.ToDateTime(myReader["birthdate"]), Convert.ToDateTime(sAgeCompareDate));
                }
                // we want the age to have only one decimal place so we are comparing it to a one decimal age restriction. Otherwise some ages will fall into an age gap and not be eligible for a class. 
                // Example: a child who is 8.91 being too old for a class that has a max age of 8.9. The intention is for any child under 9 to be valid.
                sAge = Math.Truncate(sAge * 10) / 10;
                //sDropDownList_familyMembers += "<!-- " + Convert.ToString( myReader["firstname"] ) + ": " + Convert.ToString( sAge ) + " -->";


                sHasWholeYearPrecision_min = hasWholeYearPrecision(sMinAgePrecisionID);
                sHasWholeYearPrecision_max = hasWholeYearPrecision(sMaxAgePrecisionID);

                sIsAgeToMin_greaterOrEqual_MinAge = false;
                sIsAgeToMax_lessOrEqual_MaxAge    = false;

                //sDropDownList_familyMembers += "<option>" + myReader["firstname"].ToString() + " " + myReader["lastname"].ToString() + ": [" + sMinAgePrecisionID.ToString() + "][" + sHasWholeYearPrecision_min.ToString() + "] - [" + sMaxAgePrecisionID.ToString() + "][" + sHasWholeYearPrecision_min.ToString() + "] [" + sAge.ToString() + "] - [" + sMinAge.ToString() + "] [" + sMaxAge.ToString() + "]</option>";

                if (sHasWholeYearPrecision_min)
                {
                    // truncate the age down for the comparison on the whole year precision
                    if (Convert.ToInt32(Math.Truncate(sAge)) >= Convert.ToInt32(sMinAge))
                    {
                        sIsAgeToMin_greaterOrEqual_MinAge = true;
                    }
                }
                else
                {
                    if (sAge >= sMinAge)
                    {
                        sIsAgeToMin_greaterOrEqual_MinAge = true;
                    }
                }

                if (sHasWholeYearPrecision_max)
                {
                    if (Convert.ToInt32(Math.Truncate(sAge)) <= Convert.ToInt32(sMaxAge))
                    {
                        sIsAgeToMax_lessOrEqual_MaxAge = true;
                    }
                }
                else
                {
                    if (sAge <= sMaxAge)
                    {
                        sIsAgeToMax_lessOrEqual_MaxAge = true;
                    }
                }
                //sDropDownList_familyMembers += "<option>" + myReader["firstname"].ToString() + " " + myReader["lastname"].ToString() + ": [" + iMembershipID.ToString() + "] [" + iGenderRestriction.ToString() + "] - [" + sIsAgeToMin_greaterOrEqual_MinAge.ToString() + "] - [" + sIsAgeToMax_lessOrEqual_MaxAge.ToString() + "]</option>";

                //if (sIsAgeToMin_greaterOrEqual_MinAge && sIsAgeToMax_lessOrEqual_MaxAge)
                if (sIsAgeToMin_greaterOrEqual_MinAge)
                {
                    if ((sMaxAge > 0 && sIsAgeToMax_lessOrEqual_MaxAge) || (sMaxAge <= 0 && !sIsAgeToMax_lessOrEqual_MaxAge))
                    {
                        if (iMembershipID > 0)
                        {
                            sMember = classes.determineMembership(Convert.ToInt32(myReader["familymemberid"]),
                                                                  iUserID,
                                                                  iMembershipID);
                        }

                        if (iGenderRestriction == "N")
                        {
                            sOkToAdd = true;
                        }
                        else
                        {
                            sFamilyMemberGender = classes.getFamilyMemberGender(Convert.ToInt32(myReader["familymemberid"]));

                            if (sFamilyMemberGender == iGenderRestriction)
                            {
                                sOkToAdd = true;
                            }
                            else
                            {
                                sOkToAdd = false;
                            }
                        }

                        if (sOkToAdd)
                        {
                            sFamilyMemberSelected = "";

                            sFamilyMemberOptionText = Convert.ToString(myReader["firstname"]);
                            sFamilyMemberOptionText += " ";
                            sFamilyMemberOptionText += Convert.ToString(myReader["lastname"]);

                            if (sCount == 0)
                            {
                                sFamilyMemberSelected = " selected=\"selected\"";
                                iSelectedFamilyMemberID = Convert.ToInt32(myReader["familymemberid"]);
                            }

                            if (Convert.ToString(myReader["relationship"]).ToUpper() == "CHILD")
                            {
                                if (Convert.ToString(myReader["birthdate"]) != "")
                                {
                                    sChildsAge = classes.getChildAge(Convert.ToDateTime(myReader["birthdate"]));

                                    sFamilyMemberOptionText += " - Age: " + sChildsAge.ToString() + " yrs";
                                }
                            }

                            if (sMember == "M")
                            {
                                sFamilyMemberOptionText += " - Member";

                                iMemberCount = iMemberCount + 1;
                            }

                            sDropDownList_familyMembers += "    <option value=\"" + Convert.ToString(myReader["familymemberid"]) + "\"" + sFamilyMemberSelected + ">" + sFamilyMemberOptionText + "</option>";

                            sCount = sCount + 1;
                        }
                    }
                }
            }

            sDropDownList_familyMembers += "  </select>";
            sDropDownList_familyMembers += "&nbsp;&nbsp<input type=\"button\" name=\"updateFamilyMembersButton\" id=\"updateFamilyMembersButton\" value=\"Update Family Members\" onclick=\"updateFamily('" + iUserID.ToString() + "')\" />";
            sDropDownList_familyMembers += "<div id=\"notinlist\">* Family members will not appear in the list if they do not meet the restrictions.</div>";
            sDropDownList_familyMembers += "</div>";
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        if (iMemberCount == sCount)
        {
            lcl_return = true;
        }

        return lcl_return;
    }

    public static string getFamilyMemberGender(Int32 iFamilyMemberID)
    {
        string lcl_return = "N";
        string sSQL       = "";

        sSQL  = "SELECT isnull(u.gender,'N') as gender ";
        sSQL += " FROM egov_users u, ";
        sSQL +=      " egov_familymembers f ";
        sSQL += " WHERE f.userid = u.userid ";
        sSQL += " AND f.familymemberid = " + iFamilyMemberID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToString(myReader["gender"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Int32 getChildAge(DateTime iBirthDate)
    {
        double sAge = 0;

        Int32 lcl_return = 0;

        TimeSpan sDaysDiff = DateTime.Now.Subtract(iBirthDate);

        sAge = sDaysDiff.Days / 365.25;

        // we do not want to round up ever for ages, so truncate the decimal portion. ToInt32 rounds up and down.
        lcl_return = Convert.ToInt32( Math.Truncate( sAge ) );

        return lcl_return;
    }

    public static string showCostOptions(Int32 iClassID,
                                         string iUserType,
                                         Int32 iOrgID,
                                         Boolean iAllMembers,
                                         Int32 iMemberCount,
                                         Int32 iPriceDiscountID)
    {
        Int32 iCount = 0;

        string lcl_return         = "";
        string sSQL               = "";
        string sPriceType         = "";
        string sDiscountPhrase    = classes.getDiscountPhrase(iPriceDiscountID);
        string sCheckboxPricePick = "";

        sSQL  = "SELECT p.pricetypeid, ";
        sSQL += " pricetypename, ";
        sSQL += " amount, ";
        sSQL += " pricetype, ";
        sSQL += " ismember, ";
        sSQL += " isnull(basepricetypeid,0) as basepricetypeid ";
        sSQL += " FROM egov_price_types t, ";
        sSQL +=      " egov_class_pricetype_price p ";
        sSQL += " WHERE t.pricetypeid = p.pricetypeid ";
        sSQL += " AND t.isdropin = 0 ";
        sSQL += " AND t.isfee = 0 ";
        sSQL += " AND orgid = " + iOrgID.ToString();
        sSQL += " AND classid = " + iClassID.ToString();
        sSQL += " ORDER BY p.pricetypeid ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {

            lcl_return = "<table id=\"classFees\">";
            lcl_return += "  <tr valign=\"top\">";
            lcl_return += "      <td>";
            lcl_return += "          <strong>Your Fee:</strong>";
            lcl_return += "      </td>";
            lcl_return += "      <td>";
            lcl_return += "          <table id=\"classFeesPrices\">";

            while (myReader.Read())
            {
                sPriceType = Convert.ToString(myReader["pricetype"]).ToUpper();

                //For resident/non-resident choices only show the one they get.
                if (sPriceType == "R" || sPriceType == "N")
                {
                    if (sPriceType == iUserType)
                    {
                        sCheckboxPricePick = classes.includePricePick(Convert.ToInt32(myReader["pricetypeid"]),
                                                                      Convert.ToString(myReader["pricetypename"]),
                                                                      Convert.ToDouble(myReader["amount"]),
                                                                      iCount,
                                                                      sDiscountPhrase,
                                                                      Convert.ToBoolean(myReader["ismember"]),
                                                                      iClassID,
                                                                      Convert.ToInt32(myReader["basepricetypeid"]),
                                                                      out iCount);

                        lcl_return += sCheckboxPricePick;
                    }
                }
                else
                {
                    //If it is for members and they have some members then show it.
                    if (sPriceType == "M")
                    {
                        if (iMemberCount > 0)
                        {
                            sCheckboxPricePick = classes.includePricePick(Convert.ToInt32(myReader["pricetypeid"]),
                                                                          Convert.ToString(myReader["pricetypename"]),
                                                                          Convert.ToDouble(myReader["amount"]),
                                                                          iCount,
                                                                          sDiscountPhrase,
                                                                          Convert.ToBoolean(myReader["ismember"]),
                                                                          iClassID,
                                                                          Convert.ToInt32(myReader["basepricetypeid"]),
                                                                          out iCount);

                            lcl_return += sCheckboxPricePick;
                        }
                    }
                    else
                    {
                        if (sPriceType == "O")
                        {
                            //For non-members only show if some are not members.
                            if (!iAllMembers)
                            {
                                //Include Pick
                                sCheckboxPricePick = classes.includePricePick(Convert.ToInt32(myReader["pricetypeid"]),
                                                                              Convert.ToString(myReader["pricetypename"]),
                                                                              Convert.ToDouble(myReader["amount"]),
                                                                              iCount,
                                                                              sDiscountPhrase,
                                                                              Convert.ToBoolean(myReader["ismember"]),
                                                                              iClassID,
                                                                              Convert.ToInt32(myReader["basepricetypeid"]),
                                                                              out iCount);

                                lcl_return += sCheckboxPricePick;
                            }
                        }
                        else
                        {
                            //That leaves "E" for everyone.
                            sCheckboxPricePick = classes.includePricePick(Convert.ToInt32(myReader["pricetypeid"]),
                                                                          Convert.ToString(myReader["pricetypename"]),
                                                                          Convert.ToDouble(myReader["amount"]),
                                                                          iCount,
                                                                          sDiscountPhrase,
                                                                          Convert.ToBoolean(myReader["ismember"]),
                                                                          iClassID,
                                                                          Convert.ToInt32(myReader["basepricetypeid"]),
                                                                          out iCount);

                            lcl_return += sCheckboxPricePick;
                        }
                    }
                }

            }

            lcl_return += "          </table>";
            lcl_return += "      </td>";
            lcl_return += "  </tr>";
            lcl_return += "</table>";
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        if (iCount == 0)
        {
            lcl_return = "<div class=\"noPricingOptions\">No Pricing Options</div>";
        }

        return lcl_return;
    }

    public static string showEmergencyInfo(Int32 iOrgID,
                                           Int32 iSelectedFamilyMemberID)
    {
        //NOTE:
        //  1. The name/id for the input field for the "emergencycontact" has been changed and moved.
        //  2. The hidden field for the "emergencyphone" has been moved.
        //
        //The reason for these changes are due to how we are building the screen.  On the initial screen
        //build we call this function and the screen displays this section properly.  This function is
        //also called in an "onChange" event on the FamilyMember dropdown list via a jQuery - Ajax call.
        //By doing this the form cannot "find" these fields when it attempts to POST to the class_addtocart.aspx.
        //Therefore, by moving these to fields to the same screen as the <form> and updating them via the 
        //"validateform" javascript function call as well as during the "onChange" on the FamilyMember
        //dropdown list, we can get the proper values into these fields and the POST knows where to find
        //the fields.

        Boolean sOrgHasFeatureEmergencyInfoRequired = common.orgHasFeature(Convert.ToString(iOrgID), "emergency info required");

        string lcl_return              = "";
        string sUserID                 = "0";
        string sEmergencyContact       = "";
        string sEmergencyPhone         = "";
        string sEmergencyPhoneAreaCode = "";
        string sEmergencyPhoneExchange = "";
        string sEmergencyPhoneLine     = "";

        if (sOrgHasFeatureEmergencyInfoRequired)
        {
            sUserID = getFamilyMemberUserId(Convert.ToString(iSelectedFamilyMemberID));

            sEmergencyContact = getUserContactInfo(Convert.ToInt32(sUserID), "emergencycontact");
            sEmergencyPhone   = getUserContactInfo(Convert.ToInt32(sUserID), "emergencyphone");

            if (sEmergencyPhone != "" && sEmergencyPhone.Length == 10)
            {
                //sEmergencyPhone = common.formatPhoneNumber(sEmergencyPhone);
                sEmergencyPhoneAreaCode = sEmergencyPhone.Substring(0, 3);
                sEmergencyPhoneExchange = sEmergencyPhone.Substring(3, 3);
                sEmergencyPhoneLine     = sEmergencyPhone.Substring(6);
            }

            lcl_return  = "<fieldset id=\"classEmergencyInfo\" class=\"fieldset\">";
            lcl_return += "  <legend>Emergency Contact Info</legend>";
            lcl_return += "  <div id=\"emergencyInfoMsg\">";
            lcl_return += "    We require participants to provide an emergency contact. Please review what we have on file for <strong>"; //sAttendeeName
	        lcl_return += "    </strong> and make updates as necessary. All fields are required.";
            lcl_return += "  </div>";
            lcl_return += "  <table id=\"classEmergencyInfoTable\" class=\"respTable\">";
            lcl_return += "    <tr valign=\"top\">";
            lcl_return += "        <td class=\"classEmergencyInfoLabel\"><span class=\"requiredField\">*</span>Emergency Contact:</td>";
            //lcl_return += "        <td><input type=\"text\" name=\"emergencycontact\" id=\"emergencycontact\" value=\"" + sEmergencyContact + "\" size=\"30\" maxlength=\"100\" /></td>";
            lcl_return += "        <td><input type=\"text\" name=\"emergencycontactmaint\" id=\"emergencycontactmaint\" value=\"" + sEmergencyContact + "\" size=\"30\" maxlength=\"100\" /></td>";
            lcl_return += "    </tr>";
            lcl_return += "    <tr valign=\"top\">";
            lcl_return += "        <td class=\"classEmergencyInfoLabel\"><span class=\"requiredField\">*</span>Emergency Phone:</td>";
            lcl_return += "        <td>";
            //lcl_return += "            <input type=\"hidden\" name=\"emergencyphone\" id=\"emergencyphone\" value=\"" + sEmergencyPhone + "\" size=\"30\" />";
            lcl_return += "           (<input type=\"text\" name=\"emergencyphone_areacode\" id=\"emergencyphone_areacode\" value=\"" + sEmergencyPhoneAreaCode + "\" size=\"3\" maxlength=\"3\" onKeyUp=\"return autoTab(this, 3, event);\" />)&nbsp;";
            lcl_return += "            <input type=\"text\" name=\"emergencyphone_exchange\" id=\"emergencyphone_exchange\" value=\"" + sEmergencyPhoneExchange + "\" size=\"3\" maxlength=\"3\" onKeyUp=\"return autoTab(this, 3, event);\" />&nbsp;&ndash;";
            lcl_return += "            <input type=\"text\" name=\"emergencyphone_line\" id=\"emergencyphone_line\" value=\"" + sEmergencyPhoneLine + "\" size=\"4\" maxlength=\"4\" onKeyUp=\"return autoTab(this, 4, event);\" />";
            lcl_return += "        </td>";
            lcl_return += "    </tr>";
            lcl_return += "  </table>";
            lcl_return += "</fieldset>";
        }

        return lcl_return;
    }

    public static void updateEmergencyInfo(Int32 iOrgID,
                                           Int32 iFamilyMemberID,
                                           string iEmergencyContact,
                                           string iEmergencyPhone)
    {
        Int32 sUserID = 0;

        string sSQL = "";
        string sEmergencyContact = "";
        string sEmergencyPhone = "";

        if (iFamilyMemberID != null)
        {
            sUserID = Convert.ToInt32(classes.getFamilyMemberUserId(Convert.ToString(iFamilyMemberID)));
        }
        
        if (iEmergencyContact != null)
        {
            sEmergencyContact = iEmergencyContact;
            //sEmergencyContact = common.dbSafe(sEmergencyContact);
            sEmergencyContact = sEmergencyContact.Replace("'", "''");
            sEmergencyContact = sEmergencyContact.Replace("<", "&lt;");
        }

        if (iEmergencyPhone != null)
        {
            sEmergencyPhone = common.dbSafe(iEmergencyPhone);
        }

        sEmergencyContact = "'" + sEmergencyContact + "'";
        sEmergencyPhone   = "'" + sEmergencyPhone   + "'";

        sSQL  = "UPDATE egov_users SET ";
        sSQL += " emergencycontact = " + sEmergencyContact + ", ";
        sSQL += " emergencyphone = " + sEmergencyPhone;
        sSQL += " WHERE userid = " + sUserID.ToString();
        sSQL += " AND orgid = " + iOrgID.ToString();

        common.RunSQLStatement(sSQL);
    }

    public static string getDiscountPhrase(Int32 iPriceDiscountID)
    {
        string lcl_return = "";
        string sSQL       = "";

        sSQL  = "SELECT discountdescription ";
        sSQL += " FROM egov_price_discount ";
        sSQL += " WHERE pricediscountid = " + iPriceDiscountID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToString(myReader["discountdescription"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string includePricePick(Int32 iPriceTypeID,
                                          string iPriceTypeName,
                                          double iAmount,
                                          Int32 sCount,
                                          string iDiscountPhrase,
                                          Boolean iIsMember,
                                          Int32 iClassID,
                                          Int32 iBasePriceTypeID,
                                          out Int32 iCount)
    {
        Boolean sIsClassFilled = isClassFilled(iClassID);

        double sFees      = classes.getFeeTotal(iClassID);
        double sPrice     = 0.00;
        double sBasePrice = 0.00;
        
        string lcl_return        = "";
        string sCheckedPricePick = "";
        string sDisplayPrice     = "";

        iCount = sCount;

        if (iCount == 0)
        {
            sCheckedPricePick = " checked=\"checked\"";
        }

        if (!sIsClassFilled)
        {
            sPrice = sFees + iAmount;

            if (iBasePriceTypeID != 0)
            {
                sBasePrice = classes.getBasePrice(iBasePriceTypeID, iClassID);
                sPrice = sPrice + sBasePrice;
            }

            sDisplayPrice = string.Format("{0:C}", sPrice);

            if (iIsMember)
            {
                sDisplayPrice += "&nbsp;";
                sDisplayPrice += classes.showMembership(iClassID);
            }

            if (iDiscountPhrase != "")
            {
                sDisplayPrice += "&nbsp;(" + iDiscountPhrase + ")";
            }

            lcl_return += "  <tr>";
            lcl_return += "      <td><input type=\"radio\" name=\"pricetypeid\" value=\"" + iPriceTypeID.ToString() + "\"" + sCheckedPricePick + " /></td>";
            lcl_return += "      <td><strong>" + iPriceTypeName + "</strong></td>";
            lcl_return += "      <td>" + sDisplayPrice + "</td>";
            lcl_return += "  </tr>";
        }
        else
        {
            lcl_return += "<input type=\"hidden\" name=\"pricetypeid\" value=\"" + iPriceTypeID.ToString() + "\" />";
        }

        iCount = iCount + 1;

        return lcl_return;
    }

    public static string buildTeamRosterAccessories(Int32 iOrgID,
                                                    Int32 iClassID,
                                                    string iAccessoryEnabled,
                                                    string iAccessoryInputType,
                                                    string iAccessoryType,
                                                    string iAccessoryName)
    {
        string lcl_return     = "";
        string sInputBoxSize  = "20";
        string sSQL           = "";
        string sAccessoryType = "TSHIRT";

        if (iAccessoryType != "")
        {
            sAccessoryType = iAccessoryType.ToUpper();
        }

        if (iAccessoryEnabled == "BOTH")
        {
            if (iAccessoryInputType != "")
            {
                if (iAccessoryInputType == "TEXT")
                {
                    if(sAccessoryType == "GRADE")
                    {
                        sInputBoxSize = "5";
                    }

                    lcl_return = "<input type=\"text\" name=\"" + iAccessoryName + "\" id=\"" + iAccessoryName + "\" size=\"" + sInputBoxSize + "\" maxlength=\"50\" onchange=\"clearMsg('" + iAccessoryName + "');\" />";
                }
                else
                {
                    sAccessoryType = common.dbSafe(sAccessoryType);
                    sAccessoryType = "'" + sAccessoryType + "'";

                    sSQL  = "SELECT atc.accessoryid, ";
                    sSQL += " a.accessoryname, ";
                    sSQL += " a.accessoryvalue ";
                    sSQL += " FROM egov_class_teamroster_accessories_to_class atc ";
                    sSQL +=      " INNER JOIN egov_class_teamroster_accessories a ON atc.accessoryid = a.accessoryid ";
                    sSQL += " WHERE a.orgid = " + iOrgID.ToString();
                    sSQL += " AND atc.classid = " + iClassID.ToString();
                    sSQL += " AND UPPER(a.accessorytype) = " + sAccessoryType;

                    SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
                    sqlConn.Open();

                    SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
                    SqlDataReader myReader;
                    myReader = myCommand.ExecuteReader();

                    if (myReader.HasRows)
                    {
                        lcl_return = "<select name=\"" + iAccessoryName + "\" id=\"" + iAccessoryName + "\" onchange=\"clearMsg('" + iAccessoryName + "');\">";

                        while (myReader.Read())
                        {
                            lcl_return += "  <option value=\"" + Convert.ToString(myReader["accessoryvalue"]) + "\">" + Convert.ToString(myReader["accessoryname"]) + "</option>";
                        }

                        lcl_return += "</select>";
                    }

                    myReader.Close();
                    sqlConn.Close();
                    myReader.Dispose();
                    sqlConn.Dispose();

                }
            }
        }

        return lcl_return;
    }

    public static string displayActivityTimes(Int32 iClassID,
                                              Boolean iIsParent)
    {
        Boolean sTimeHasAvailability = false;
        Boolean sNoWait              = false;
        Boolean sIsChecked           = false;
        Boolean sBuyOrWaitSet        = false;

        Int32 sLineCount             = 0;
        Int32 sPickableCount         = 0;
        Int32 sTotalAvailability     = 0;
        Int32 sTotalSeriesEnrollment = 0;

        string lcl_return                = "";
        string sSQL                      = "";
        string sTimeID                   = "";
        string sActivityNo               = "&nbsp;";
        string sOldActivityNo            = "&nbsp;";
        string sClassSizeMin             = "&nbsp;";
        string sClassSizeMax             = "&nbsp;";
        string sEnrollmentSize           = "&nbsp;";
        string sWaitListSize             = "&nbsp;";
        string sSunday                   = "&nbsp;";
        string sMonday                   = "&nbsp;";
        string sTuesday                  = "&nbsp;";
        string sWednesday                = "&nbsp;";
        string sThursday                 = "&nbsp;";
        string sFriday                   = "&nbsp;";
        string sSaturday                 = "&nbsp;";
        string sStartTime                = "&nbsp;";
        string sEndTime                  = "&nbsp;";
        string sBGColor                  = "#eeeeee";
        string sRowClass                 = "";
        string sBuyOrWait                = "";
        string sFirstBuyOrWait           = "B";
        string sDisplayTimeOption        = "";
        string sDisplayAvailability      = "";
        string sTimeIDChecked            = "";
        string sTimeIDOnClick            = "";
        string sTimeIDDisabled           = "";
        string sAvailabilityHiddenFields = "";

        sSQL = "SELECT t.timeid, ";
        sSQL += " activityno, ";
        sSQL += " min, ";
        sSQL += " isnull(max,-999) as max, ";
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
        sSQL += " ORDER BY activityno, timedayid ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            lcl_return += "<table cellspacing=\"0\" id=\"offeringActivities\">";
            lcl_return += "  <thead>";
            lcl_return += "  <tr>";
            lcl_return += "      <td class=\"offeringActivities_activityno\">Activity No</td>";
            lcl_return += "      <td>Availability</td>";
            lcl_return += "      <td>Su</td>";
            lcl_return += "      <td>Mo</td>";
            lcl_return += "      <td>Tu</td>";
            lcl_return += "      <td>We</td>";
            lcl_return += "      <td>Th</td>";
            lcl_return += "      <td>Fr</td>";
            lcl_return += "      <td>Sa</td>";
            lcl_return += "      <td>Starts</td>";
            lcl_return += "      <td>Ends</td>";
            lcl_return += "  </tr>";
            lcl_return += "  </thead>";

            while (myReader.Read())
            {
                sLineCount               += 1;
                sBGColor                  = common.changeBGColor(sBGColor, "#ffffff", "#eeeeee");
                sRowClass                 = "offeringActivities_row_light";
                sTimeID                   = "";
                sActivityNo               = "";
                sClassSizeMin             = "";
                sClassSizeMax             = "";
                sEnrollmentSize           = "";
                sWaitListSize             = "";
                sSunday                   = "";
                sMonday                   = "";
                sTuesday                  = "";
                sWednesday                = "";
                sThursday                 = "";
                sFriday                   = "";
                sSaturday                 = "";
                sStartTime                = "";
                sEndTime                  = "";
                sDisplayTimeOption        = "";
                sAvailabilityHiddenFields = "";

                if (sBGColor == "#eeeeee")
                {
                    sRowClass = "offeringActivities_row_dark";
                }

                if(myReader["timeid"].ToString() != null)
                {
                    sTimeID = myReader["timeid"].ToString();
                }

                if (myReader["activityno"].ToString() != null)
                {
                    sActivityNo = myReader["activityno"].ToString().Trim();
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

                lcl_return += "  <tbody>";
                lcl_return += "  <tr class=\"" + sRowClass + "\">";

                if (sActivityNo != sOldActivityNo)
                {
                    sOldActivityNo       = sActivityNo;
                    sNoWait              = false;
                    sBuyOrWait           = "B";
                    sTimeHasAvailability = classes.timeHasAvailability(Convert.ToInt32(sTimeID));
                    sTimeIDChecked       = "";
                    sTimeIDOnClick       = "";
                    sTimeIDDisabled      = "";

                    //Setup Activity Number and Time Option (checkbox)
                    if (sTimeHasAvailability)
                    {
                        //"Check" ONLY the first option if there are multiple options
                        if (!sIsChecked)
                        {
                            sTimeIDChecked = " checked=\"checked\"";
                            sIsChecked     = true;
                        }

                        sTimeIDOnClick = " onclick=\"changebuyorwait('" + sTimeID + "');\"";
                        sPickableCount = sPickableCount + 1;
                    }
                    else
                    {
                        sNoWait         = true;
                        sTimeIDDisabled = " disabled=\"disabled\"";
                    }

                    sDisplayTimeOption = "<input type=\"radio\" name=\"timeid\" id=\"timeid" + sLineCount.ToString() + "\" value=\"" + sTimeID + "\"" + sTimeIDChecked + sTimeIDOnClick + sTimeIDDisabled + " />";
                    
                    //Setup Availability
                    if (!iIsParent)
                    {
                        if (Convert.ToInt32(myReader["max"]) != -999)
                        {
                            if (Convert.ToInt32(myReader["waitlistsize"]) > 0)
                            {
                                sTotalAvailability = 0;
                            }
                            else
                            {
                                //sTotalAvailability = Convert.ToInt32(myReader["max"]) - Convert.ToInt32(myReader["enrollmentsize"]) + Convert.ToInt32(myReader["waitlistsize"]);
                                sTotalAvailability   = Convert.ToInt32(myReader["max"]) - Convert.ToInt32(myReader["enrollmentsize"]);
                                sDisplayAvailability = sTotalAvailability.ToString();
                            }

                            if (sTotalAvailability < 1)
                            {
                                if (sNoWait)
                                {
                                    sDisplayAvailability = "0<br />Filled";
                                }
                                else
                                {
                                    sDisplayAvailability = "0<br />(Wait List Only)";
                                }

                                sBuyOrWait = "W";
                                sTotalAvailability = 0;
                            }
                        }
                        else
                        {
                            sDisplayAvailability = "No Limit";
                            sTotalAvailability = 9999;
                        }
                    }
                    else
                    {
                        //Max availability of parent is least available of any child
                        if (Convert.ToInt32(myReader["max"]) != -999)
                        {
                            sTotalSeriesEnrollment = classes.getSeriesEnrollment(iClassID);
                            sTotalAvailability     = Convert.ToInt32(myReader["max"]) - sTotalSeriesEnrollment;
                            sDisplayAvailability   = sTotalAvailability.ToString();

                            if (sTotalAvailability < 1)
                            {
                                if (sNoWait)
                                {
                                    sDisplayAvailability = "0<br />(Filled)";
                                }
                                else
                                {
                                    sDisplayAvailability = "0<br />(Wait List Only)";
                                }

                                sBuyOrWait = "W";
                                sTotalAvailability = 0;
                            }
                        }
                        else
                        {
                            sDisplayAvailability = "No Limit";
                            sTotalAvailability   = 9999;
                        }
                    }
                    
                    if (sPickableCount == 1 && !sBuyOrWaitSet)
                    {
                        sFirstBuyOrWait = sBuyOrWait;
                        sBuyOrWaitSet = true;
                    }

                    sAvailabilityHiddenFields  = "<input type=\"hidden\" name=\"firstbuyorwait" + sTimeID + "\" id=\"firstbuyorwait" + sTimeID + "\" value=\"" + sPickableCount.ToString() + " " + sFirstBuyOrWait.ToString() + "\" />";
                    sAvailabilityHiddenFields += "<input type=\"hidden\" name=\"buyorwait"      + sTimeID + "\" id=\"buyorwait"      + sTimeID + "\" value=\"" + sBuyOrWait + "\" />";
                    sAvailabilityHiddenFields += "<input type=\"hidden\" name=\"avail"          + sTimeID + "\" id=\"avail"          + sTimeID + "\" value=\"" + sTotalAvailability.ToString() + "\" />";

                    lcl_return += "      <td class=\"offeringActivities_activityno\">" + sDisplayTimeOption + sActivityNo + "</td>";
                    lcl_return += "      <td>" + sDisplayAvailability + sAvailabilityHiddenFields + "</td>";
                }
                else
                {
                    lcl_return += "      <td colspan=\"2\">&nbsp;</td>";
                }

                lcl_return += "      <td>" + sSunday + "</td>";
                lcl_return += "      <td>" + sMonday + "</td>";
                lcl_return += "      <td>" + sTuesday + "</td>";
                lcl_return += "      <td>" + sWednesday + "</td>";
                lcl_return += "      <td>" + sThursday + "</td>";
                lcl_return += "      <td>" + sFriday + "</td>";
                lcl_return += "      <td>" + sSaturday + "</td>";
                lcl_return += "      <td>" + sStartTime + "</td>";
                lcl_return += "      <td>" + sEndTime + "</td>";
                lcl_return += "  </tr>";
                lcl_return += "  </tbody>";
            }

            lcl_return += "</table>";
            lcl_return += "<input type=\"hidden\" name=\"buyorwait\" id=\"buyorwait\" value=\"" + sFirstBuyOrWait + "\" />";
            lcl_return += "<input type=\"hidden\" name=\"activityTimeTotalCount\" id=\"activityTimeTotalCount\" value=\"" + sLineCount.ToString() + "\" />";
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );

        return lcl_return;
    }

    public static Boolean timeHasAvailability(Int32 iTimeID)
    {
        Boolean lcl_return = false;

        string sSQL = "";

        sSQL  = "SELECT enrollmentsize, ";
        sSQL += " isnull(max,9999) as max, ";
        sSQL += " waitlistsize, ";
        sSQL += " isnull(waitlistmax,9999) as waitlistmax ";
        sSQL += " FROM egov_class_time ";
        sSQL += " WHERE timeid = " + iTimeID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (Convert.ToInt32(myReader["enrollmentsize"]) >= Convert.ToInt32(myReader["max"]) &&
                Convert.ToInt32(myReader["waitlistsize"]) >= Convert.ToInt32(myReader["waitlistmax"]))
            {
                lcl_return = false;
            }
            else
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

    public static string showTermsList(Int32 iOrgID)
    {
        string lcl_return  = "";
        string sSQL        = "";
        string sWaiverDesc = "";

        sSQL  = "SELECT waiverdescription ";
        sSQL += " FROM egov_class_waivers ";
        sSQL += " WHERE waivertype = 'TERM' ";
        sSQL += " AND orgid = " + iOrgID.ToString();
        sSQL += " ORDER BY waivername ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            sWaiverDesc = Convert.ToString(myReader["waiverdescription"]).Trim();
            sWaiverDesc = common.decodeUTFString(sWaiverDesc);

            lcl_return  = "<div class\"class_terms\">";
            lcl_return += "  <div class=\"class_waivername\">Waiver and Release</div>";
            //lcl_return += "  <textarea class=\"class_termstext\" readonly=\"readonly\">" + Convert.ToString(myReader["waiverdescription"]) + "</textarea>";
            lcl_return += "  <div class=\"class_termstext\">" + sWaiverDesc + "</div>";
            lcl_return += "  <div>";
            lcl_return += "    <input type=\"checkbox\" name=\"terms\" id=\"terms\" onclick=\"if(this.checked){clearMsg('terms')}\" />";
            lcl_return += "    I agree.  You must check here to indicate that you agree to the above terms and conditions before ";
            lcl_return += "    continuing registration.";
            lcl_return += "  </div>";
            lcl_return += "</div>";
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getActivityNo(Int32 iTimeID)
    {
        string lcl_return = "";
        string sSQL       = "";

        sSQL  = "SELECT isnull(activityno,'') as activityno ";
        sSQL += " FROM egov_class_time ";
        sSQL += " WHERE timeid = " + Convert.ToString(iTimeID);

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToString(myReader["activityno"]);
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );

        return lcl_return;
    }

    public static Int32 getClassPriceDiscountID(Int32 iClassID)
    {
        Int32 lcl_return = 0;

        string sSQL = "";

        sSQL  = "SELECT isnull(pricediscountid,0) as pricediscountid ";
        sSQL += " FROM egov_class ";
        sSQL += " WHERE classid = " + iClassID.ToString( );

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            myReader.Read( );

            lcl_return = Convert.ToInt32( myReader["pricediscountid"] );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );

        return lcl_return;
    }

    public static Int32 getCurrentAvailability(Int32 iClassTimeID)
    {
        Int32 lcl_return = 0;

        string sSQL = "";

        sSQL  = "SELECT isnull(max, 999999) as max, ";
        sSQL += " enrollmentsize ";
        sSQL += " FROM egov_class_time ";
        sSQL += " WHERE timeid = " + Convert.ToString( iClassTimeID );

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sSQL, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            myReader.Read( );

            lcl_return = Convert.ToInt32( myReader["max"] ) - Convert.ToInt32( myReader["enrollmentsize"] );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );

        if (lcl_return < 0)
        {
            lcl_return = 0;
        }

        return lcl_return;
    }

    public static Boolean classCartHasItems( string sessionId )
    {
        Boolean cartHasItems = false;

        string sSessionID = "";

        if (sessionId != "")
        {
            sSessionID = common.dbSafe(sessionId);
        }

        sSessionID = "'" + sSessionID + "'";

        string sql  = "SELECT CASE WHEN COUNT(cartid) = 0 THEN CAST(0 AS BIT) ELSE CAST(1 AS BIT) END AS CartHasItems ";
               sql += " FROM egov_class_cart ";
               sql += " WHERE sessionid_csharp = " + sSessionID;

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            myReader.Read( );
            Boolean.TryParse( myReader["CartHasItems"].ToString( ), out cartHasItems );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );

        return cartHasItems;
    }

    public static double getCartTotalAmount( string _SessionId )
    {
        double cartTotalAmount = 0;

        string sSessionID = "";

        if (_SessionId != "")
        {
            sSessionID = common.dbSafe(_SessionId);
        }

        sSessionID = "'" + sSessionID + "'";

        //string sql = "SELECT SUM(ISNULL(amount,0.00)) AS CartTotalAmount FROM egov_class_cart WHERE (isnull(sessionid_csharp,sessionid) = " + _SessionId;
        string sql  = "SELECT SUM(ISNULL(amount,0.00)) AS CartTotalAmount ";
               sql += " FROM egov_class_cart ";
               sql += " WHERE sessionid_csharp = " + sSessionID;

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            myReader.Read( );
            double.TryParse( myReader["CartTotalAmount"].ToString( ), out cartTotalAmount );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );

        return cartTotalAmount;
    }

    public static string makeALedgerEntry( string _PaymentId, string _OrgId, string _EntryType, string _AccountId, double _Amount, string _ItemTypeId, string _PlusMinus, string _ItemId, string _IsPaymentAccount, string _PaymentTypeId, string _PriorBalance, string _PriceTypeId )
    {
        string ledgerEntryId = "0";

        string sql = "INSERT INTO egov_accounts_ledger ( paymentid, orgid, entrytype, accountid, amount, itemtypeid, plusminus, ";
        sql += "itemid, ispaymentaccount, paymenttypeid, priorbalance, pricetypeid ) VALUES ( " + _PaymentId + ", " + _OrgId + ", '" + _EntryType + "', ";
        sql += _AccountId + ", " + _Amount.ToString( "F2" ) + ", " + _ItemTypeId + ", '" + _PlusMinus + "', " + _ItemId + ", " + _IsPaymentAccount + ", ";
        sql += _PaymentTypeId + ", " + _PriorBalance + ", " + _PriceTypeId + " )";

        ledgerEntryId = common.RunInsertStatement( sql );

        return ledgerEntryId;
    }


    public static string getFamilyMemberUserId( string _FamilyMemberId )
    {
        string userId = "0";

        string sql = "SELECT userid FROM egov_familymembers WHERE familymemberid = " + _FamilyMemberId;

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            myReader.Read( );
            userId = myReader["userid"].ToString( );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );

        return userId;
    }

    public static string getCartItemAccountId( string _CartId, string _PriceTypeId )
    {
        string accountId = "0";

        string sql = "SELECT ISNULL(P.accountid,0) AS accountid FROM egov_class_cart C, egov_class_pricetype_price P WHERE C.cartid = " + _CartId;
        sql += " AND C.classid = P.classid AND P.pricetypeid = " + _PriceTypeId;

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            myReader.Read( );
            accountId = myReader["accountid"].ToString( );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );

        return accountId;
    }


    public static void addClassLedgerEntries( string _OrgId, string _PaymentId, string _CartId, string _ItemTypeId, string _ClassListId, string _EntryType, string _PlusMinus )
    {
        string accountId = "NULL";
        string ledgerId;
        double amount = 0;
        string priceTypeId;

        // pull the price types associated to a cart item and the ledger entries
        string sql = "SELECT pricetypeid, ISNULL(amount,0) AS amount FROM egov_class_cart_price WHERE cartid = " + _CartId;

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        while (myReader.Read( ))
        {
            priceTypeId = myReader["pricetypeid"].ToString( );

            accountId = getCartItemAccountId( _CartId, priceTypeId );
            if (accountId == "0")
                accountId = "NULL";

            double.TryParse( myReader["amount"].ToString( ), out amount );

            ledgerId = makeALedgerEntry( _PaymentId, _OrgId, _EntryType, accountId, amount, _ItemTypeId, _PlusMinus, _ClassListId, "0", "NULL", "NULL", priceTypeId );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );
    }


    public static string getAssignedAdminEmail( string _OrgId )
    {
        string adminEmail = "";

        string sql = "";
        
        sql  = "SELECT ISNULL(assigned_email,'') AS assigned_email ";
        sql += " FROM dbo.egov_paymentservices ";
        sql += " WHERE paymentservicename = 'Classes and Events' ";
        sql += " AND orgid = " + _OrgId;

        SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
        sqlConn.Open( );

        SqlCommand myCommand = new SqlCommand( sql, sqlConn );
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader( );

        if (myReader.HasRows)
        {
            myReader.Read( );
            adminEmail = myReader["assigned_email"].ToString( );
        }

        myReader.Close( );
        sqlConn.Close( );
        myReader.Dispose( );
        sqlConn.Dispose( );

        return adminEmail;
    }

    public static double getAmount(Int32 iPriceTypeID,
                                   Int32 iClassID)
    {
        double lcl_return = 0.00;
        double sBaseFee   = 0.00;
        double sFees      = classes.getFeeTotal(iClassID);

        string sSQL = "";
        
        sSQL  = "SELECT p.pricetypeid, ";
        sSQL += " pricetypename, ";
        sSQL += " amount, ";
        sSQL += " pricetype, ";
        sSQL += " ismember, ";
        sSQL += " isnull(basepricetypeid,0) as basepricetypeid ";
        sSQL += " FROM egov_price_types t, ";
        sSQL +=      " egov_class_pricetype_price p ";
        sSQL += " WHERE t.pricetypeid = p.pricetypeid ";
        sSQL += " AND t.isdropin = 0 ";
        sSQL += " AND t.isfee = 0 ";
        sSQL += " AND classid = " + iClassID.ToString();
        sSQL += " AND t.pricetypeid = " + iPriceTypeID.ToString();
        
        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (Convert.ToInt32(myReader["basepricetypeid"]) > 0)
            {
                sBaseFee = classes.getBasePrice(Convert.ToInt32(myReader["basepricetypeid"]), 
                                                iClassID);
            }

            lcl_return = (sFees + Convert.ToDouble(myReader["amount"]) + sBaseFee);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static void addToCartPrice(Int32 iCartID,
                                      Int32 iPriceTypeID,
                                      double iAmount,
                                      Int32 iQuantity)
    {
        double sTotalPrice = 0.00;

        string sSQL = "";

        sTotalPrice = Convert.ToDouble(iQuantity * iAmount);

        sSQL = "INSERT INTO egov_class_cart_price (";
        sSQL += "cartid,";
        sSQL += "pricetypeid,";
        sSQL += "unitprice,";
        sSQL += "amount";
        sSQL += ") VALUES (";
        sSQL += iCartID + ", ";
        sSQL += iPriceTypeID + ", ";
        sSQL += iAmount + ", ";
        sSQL += sTotalPrice;
        sSQL += ")";

        common.RunSQLStatement(sSQL);

    }

    public static void updateClassTime(Int32 iTimeID,
                                       Int32 iQuantity,
                                       string iBuyOrWait)
    {
        Int32 sQty = 0;
        Int32 sColumnNewTotal = 0;

        string sSQL        = "";
        string sSQLu       = "";
        string sColumnName = "waitlistsize";

        if (iBuyOrWait == "B")
        {
            sColumnName = "enrollmentsize";
        }

        sSQL = "SELECT " + sColumnName;
        sSQL += " FROM egov_class_time ";
        sSQL += " WHERE timeid = " + iTimeID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            sQty = Convert.ToInt32(myReader[sColumnName]);
            sColumnNewTotal = sQty + iQuantity;

            sSQLu  = "UPDATE egov_class_time SET ";
            sSQLu += sColumnName + " = " + sColumnNewTotal.ToString();
            sSQLu += " WHERE timeid = " + iTimeID.ToString();

            common.RunSQLStatement(sSQLu);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
    }

    public static void updateClassTimeSeriesChildren(Int32 iClassID,
                                                     Int32 iQuantity,
                                                     string iBuyOrWait)
    {
        string sSQL = "";

        sSQL  = "SELECT c.classid, ";
        sSQL += " t.timeid ";
        sSQL += " FROM egov_class c, ";
        sSQL +=      " egov_class_time t ";
        sSQL += " WHERE c.classid = t.classid ";
        sSQL += " AND c.parentclassid = " + iClassID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                classes.updateClassTime(Convert.ToInt32(myReader["timeid"]),
                                        iQuantity,
                                        iBuyOrWait);
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
    }

    public static void resetCartPrices()
    {
        double sItemFees = 0.00;

        Int32 sCartID      = 0;
        Int32 sClassID     = 0;
        Int32 sPriceTypeID = 0;
        Int32 sQuantity    = 0;

        string sSQL = "";
        string sSessionID = HttpContext.Current.Session.SessionID;

        sSQL  = "SELECT cc.cartid, ";
        sSQL += " pt.amount, ";
        sSQL += " cc.quantity, ";
        sSQL += " cc.pricetypeid, ";
        sSQL += " cc.classid ";
        sSQL += " FROM egov_class_cart cc, ";
        sSQL +=      "egov_class_pricetype_price pt ";
        sSQL += " WHERE cc.classid = pt.classid ";
        sSQL += " AND cc.pricetypeid = pt.pricetypeid ";
        sSQL += " AND cc.buyorwait = 'B' ";
        //sSQL += " AND cc.sessionid = " + sSessionID;
        sSQL += " AND cc.sessionid_csharp = '" + sSessionID + "'";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                sCartID      = Convert.ToInt32(myReader["cartid"]);
                sClassID     = Convert.ToInt32(myReader["classid"]);
                sPriceTypeID = Convert.ToInt32(myReader["pricetypeid"]);
                sQuantity    = Convert.ToInt32(myReader["quantity"]);

                //Get the total price per item
                sItemFees = classes.getAmount(sPriceTypeID,
                                              sClassID);

                //Set the total price visible in the cart
                classes.setPriceInCart(sCartID,
                                       Convert.ToDouble(sQuantity * sItemFees));

                //Clear the class_cart_price table
                classes.deleteCartPrices(sCartID);

                //Create new class cart prices
                classes.createClassCartPrices(sCartID,
                                              sClassID,
                                              sQuantity,
                                              sPriceTypeID);
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

    }

    public static void setPriceInCart(Int32 iCartID,
                                      double iAmount)
    {
        string sSQL = "";

        sSQL  = "UPDATE egov_class_cart SET ";
        sSQL += " amount = " + iAmount.ToString();
        sSQL += " WHERE cartid = " + iCartID.ToString();

        common.RunSQLStatement(sSQL);
    }

    public static void updateCartPrice(Int32 iCartID,
                                       Int32 iPriceTypeID,
                                       double iAmount,
                                       Int32 iQuantity)
    {
        double sTotalPrice = Convert.ToDouble(iQuantity * iAmount);

        string sSQL = "";

        sSQL  = "UPDATE egov_class_cart_price SET ";
        sSQL += " unitprice = " + iAmount.ToString() + ", ";
        sSQL += " amount = " + sTotalPrice.ToString();
        sSQL += " WHERE cartid = " + iCartID.ToString();
        sSQL += " AND pricetypeid = " + iPriceTypeID.ToString();

        common.RunSQLStatement(sSQL);
    }

    public static void deleteCartPrices(Int32 iCartID)
    {
        string sSQL = "";

        sSQL  = "DELETE FROM egov_class_cart_price ";
        sSQL += " WHERE cartid = " + iCartID.ToString();

        common.RunSQLStatement(sSQL);
    }

    public static void createClassCartPrices(Int32 iCartID,
                                             Int32 iClassID,
                                             Int32 iQuantity,
                                             Int32 iPriceTypeID)
    {
        string sSQL = "";

        //Input the extra fees
        classes.addFeesToCartPrice(iCartID,
                                   iClassID,
                                   iQuantity);

        //Get the set of all pricetypes that apply to this class
        sSQL  = "SELECT p.pricetypeid, ";
        sSQL += " pricetypename, ";
        sSQL += " isnull(amount, 0.00) as amount, ";
        sSQL += " pricetype, ";
        sSQL += " ismember, ";
        sSQL += " isnull(basepricetypeid, 0) as basepricetypeid ";
        sSQL += " FROM egov_price_types t, ";
        sSQL +=      " egov_class_pricetype_price p ";
        sSQL += " WHERE t.pricetypeid = p.pricetypeid ";
        sSQL += " AND t.isdropin = 0 ";
        sSQL += " AND t.isfee = 0 ";
        sSQL += " AND classid = " + iClassID.ToString();
        sSQL += " AND p.pricetypeid = " + iPriceTypeID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            classes.addToCartPrice(iCartID,
                                   Convert.ToInt32(myReader["pricetypeid"]),
                                   Convert.ToDouble(myReader["amount"]),
                                   iQuantity);

            if (Convert.ToInt32(myReader["basepricetypeid"]) > 0)
            {
                classes.addBaseToCartPrice(Convert.ToInt32(myReader["basepricetypeid"]),
                                           iCartID,
                                           iClassID,
                                           iQuantity);
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

    }

    public static void addFeesToCartPrice(Int32 iCartID,
                                          Int32 iClassID,
                                          Int32 iQuantity)
    {
        string sSQL = "";

        sSQL  = "SELECT pt.pricetypeid, ";
        sSQL += " isnull(amount, 0.00) as amount ";
        sSQL += " FROM egov_class_pricetype_price cpp, ";
        sSQL +=      " egov_price_types pt ";
        sSQL += " WHERE cpp.pricetypeid = pt.pricetypeid ";
        sSQL += " AND isactiveforclasses = 1 ";
        sSQL += " AND isfee = 1 ";
        sSQL += " AND classid = " + iClassID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                classes.addToCartPrice(iCartID,
                                       Convert.ToInt32(myReader["pricetypeid"]),
                                       Convert.ToDouble(myReader["amount"]),
                                       iQuantity);
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
    }

    public static void addBaseToCartPrice(Int32 iPriceTypeID,
                                          Int32 iCartID,
                                          Int32 iClassID,
                                          Int32 iQuantity)
    {
        string sSQL = "";

        sSQL  = "SELECT isnull(amount, 0.00) as amount ";
        sSQL += " FROM egov_class_pricetype_price ";
        sSQL += " WHERE pricetypeid = " + iPriceTypeID.ToString();
        sSQL += " AND classid = " + iClassID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            classes.addToCartPrice(iCartID,
                                   iPriceTypeID,
                                   Convert.ToDouble(myReader["amount"]),
                                   iQuantity);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
    }

    public static Boolean hasCorrectDiscountQtyForModulus(Int32 iCartID,
                                                          Int32 iPriceDiscountID,
                                                          Int32 iOptionID,
                                                          Boolean iIsShared,
                                                          Int32 iClassID,
                                                          Int32 iQtyRequired)
    {
        Boolean lcl_return = false;

        string sSQL = "";
        string sSessionID = HttpContext.Current.Session.SessionID;

        sSQL  = "SELECT sum(cc.quantity) as qty ";
        sSQL += " FROM egov_class_cart cc, ";
        sSQL +=      " egov_class c ";
        sSQL += " WHERE cc.classid = c.classid ";
        sSQL += " AND cc.buyorwait = 'B' ";
        sSQL += " AND cc.sessionid_csharp = '" + sSessionID + "' ";
        sSQL += " AND c.pricediscountid = " + iPriceDiscountID.ToString();
        sSQL += " AND c.optionid = " + iOptionID.ToString();

        if (iIsShared)
        {
            sSQL += " AND cc.classid = " + iClassID.ToString();
        }

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if (Convert.ToInt32(myReader["qty"]) >= iQtyRequired)
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

    public static void determineDiscounts()
    {
        Boolean sIsShared = false;
        Boolean sHasCorrectDiscountQtyForModulus = false;

        double sAmount         = 0.00;
        double sDiscountAmount = 0.00;
        double sUseThisAmount  = 0.00;
        double sPrice          = 0.00;

        Int32 sOldDiscountID   = 0;
        Int32 sOldClassID      = 0;
        Int32 sCount           = 0;
        Int32 sPriceDiscountID = 0;
        Int32 sClassID         = 0;
        Int32 sQuantity        = 0;
        Int32 sOptionID        = 0;
        Int32 sQtyRequired     = 0;
        Int32 sCartID          = 0;
        Int32 sPriceTypeID     = 0;
        Int32 sFullPriceCount  = 0;
        Int32 sFullPriceQty    = 0;
        Int32 sDiscountQty     = 0;

        string sSQL = "";
        string sSessionID = HttpContext.Current.Session.SessionID;
        string sDiscountType = "";

        sSQL  = "SELECT cc.cartid, ";
        sSQL += " cc.classid, ";
        sSQL += " cc.familymemberid, ";
        sSQL += " cc.quantity, ";
        sSQL += " d.discountamount, ";
        sSQL += " cc.amount, ";
        sSQL += " d.pricediscountid, ";
        sSQL += " c.optionid, ";
        sSQL += " t.discounttype, ";
        sSQL += " d.isshared, ";
        sSQL += " d.qtyrequired, ";
        sSQL += " cc.pricetypeid ";
        sSQL += " FROM egov_class_cart cc, ";
        sSQL +=      " egov_class c, ";
        sSQL +=      " egov_price_discount d, ";
        sSQL +=      " egov_class_pricetype_price pt, ";
        sSQL +=      " egov_price_discount_types t ";
        sSQL += " WHERE cc.classid = c.classid ";
        sSQL += " AND c.pricediscountid = d.pricediscountid ";
        sSQL += " AND t.discounttypeid = d.discounttypeid ";
        sSQL += " AND cc.pricetypeid = pt.pricetypeid ";
        sSQL += " AND cc.classid = pt.classid ";
        sSQL += " AND cc.buyorwait = 'B' ";
        sSQL += " AND cc.sessionid_csharp = '" + sSessionID + "' ";
        
        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                sIsShared        = Convert.ToBoolean(myReader["isshared"]);

				sAmount          = Convert.ToDouble(myReader["amount"]);
                sDiscountAmount  = Convert.ToDouble(myReader["discountamount"]);

                sPriceDiscountID = Convert.ToInt32(myReader["pricediscountid"]);
                sClassID         = Convert.ToInt32(myReader["classid"]);
                sQuantity        = Convert.ToInt32(myReader["quantity"]);
                sOptionID        = Convert.ToInt32(myReader["optionid"]);
                sQtyRequired     = Convert.ToInt32(myReader["qtyrequired"]);
                sCartID          = Convert.ToInt32(myReader["cartid"]);
                sPriceTypeID     = Convert.ToInt32(myReader["pricetypeid"]);

                sDiscountType    = Convert.ToString(myReader["discounttype"]).ToUpper();

                if (sOldDiscountID != sPriceDiscountID)
                {
                    sOldDiscountID = sPriceDiscountID;
                    sOldClassID    = sClassID;
                    sCount         = sQuantity;
                }
                else
                {
                    //Registration
                    if (sOptionID == 1)
                    {
                        //Shared amoung classes
                        if (sIsShared)
                        {
                            sCount = sCount + sQuantity;
                        }
                        else
                        {
                            //Same Class
                            if (sOldClassID == sClassID)
                            {
                                sCount = sCount + sQuantity;
                            }
                            else
                            {
                                //Different classes, not shared
                                sCount = sQuantity;
                            }
                        }

                        sOldClassID = sClassID;
                    }
                    else
                    {
                        //Tickets = Always the cart row quantity
                        sCount = sQuantity;
                    }
                }
                
                if (sDiscountType == "THRESHOLD")
                {
                    if (sOptionID == 1)
                    {
                        //Registered attendees
                        if (sCount >= sQtyRequired)
                        {
                            //Apply the discount
                            sUseThisAmount = sDiscountAmount;
                        }
                        else
                        {
                            //Regular price
                            sUseThisAmount = sAmount;
                        }

                        classes.setPriceInCart(sCartID,
                                               sUseThisAmount);

                        classes.updateCartPrice(sCartID,
                                                sPriceTypeID,
                                                sUseThisAmount,
                                                sQuantity);
                    }
                    else
                    {
                        //Ticketed events
                        if (sCount > sQtyRequired)
                        {
                            sFullPriceCount = sQtyRequired = 1;
                            sPrice = (sFullPriceCount * sAmount) + ((sQuantity - sFullPriceCount) * sDiscountAmount);
                        }
                        else
                        {
                            sPrice = sQuantity * sAmount;
                        }

                        classes.setPriceInCart(sCartID,
                                               sPrice);

                        classes.updateCartPrice(sCartID,
                                                sPriceTypeID,
                                                sAmount,
                                                sQuantity);
                    }
                }
                else
                {
                    //Couples
                    if (sDiscountType == "COUPLES")
                    {
                        if (sOptionID == 1)
                        {
                            sHasCorrectDiscountQtyForModulus = classes.hasCorrectDiscountQtyForModulus(sCartID,
                                                                                                       sPriceDiscountID,
                                                                                                       sOptionID,                                                                                                       sIsShared,
                                                                                                       sClassID,
                                                                                                       sQtyRequired);

                            sUseThisAmount = sAmount;

                            if (sHasCorrectDiscountQtyForModulus)
                            {
                                if ((sCount % sQtyRequired) == 0)  //"%" in C# replaces "mod" in VBScript
                                {
                                    sUseThisAmount = sDiscountAmount;
                                }
                            }

                            classes.setPriceInCart(sCartID,
                                                   sUseThisAmount);

                            classes.updateCartPrice(sCartID,
                                                    sPriceTypeID,
                                                    sUseThisAmount,
                                                    sQuantity);
                        }
                        else
                        {
                            //Tickets
                            if (sCount >= sQtyRequired)
                            {
                                //Figure out how many are at full price
                                sFullPriceQty = Convert.ToInt32(Convert.ToDouble(sQuantity / sQtyRequired) + 0.5);
                                sDiscountQty = sQuantity - sFullPriceQty;
                                sPrice = sFullPriceQty * classes.getAmount(sPriceTypeID,
                                                                           sClassID);
                                sPrice = sPrice + (sDiscountQty * sDiscountAmount);
                            }
                            else
                            {
                                sPrice = sQuantity * classes.getAmount(sPriceTypeID,
                                                                       sClassID);
                            }

                            classes.setPriceInCart(sCartID,
                                                   sPrice);

                            classes.updateCartPrice(sCartID,
                                                    sPriceTypeID,
                                                    classes.getAmount(sPriceTypeID, sClassID),
                                                    sQuantity);
                        }
                    }
                }
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
    }

    public static void removeItemFromCart(Int32 iCartID,
                                          Int32 iTimeID,
                                          string iBuyOrWait)
    {
        Boolean sIsParent = false;

        Int32 sCartQty = 0;
        Int32 sClassID = 0;
        Int32 sClassTypeID = 0;
        Int32 sParentClassID = 0;

        string sSQL = "";
        string sSQL2 = "";
        string sBuyOrWait = "";

        sSQL  = "SELECT cc.classid, ";
        sSQL += " cc.quantity, ";
        sSQL += " cc.isparent, ";
        sSQL += " isnull(cc.classtypeid, 0) as classtypeid, ";
        sSQL += " cc.buyorwait, ";
        sSQL += " isnull(c.parentclassid, 0) as parentclassid ";
        sSQL += " FROM egov_class_cart cc, ";
        sSQL +=      " egov_class c ";
        sSQL += " WHERE c.classid = cc.classid ";
        sSQL += " AND cc.cartid = " + iCartID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            sCartQty       = -(Convert.ToInt32(myReader["quantity"]));
            sClassID       = Convert.ToInt32(myReader["classid"]);
            sBuyOrWait     = Convert.ToString(myReader["buyorwait"]);
            sIsParent      = Convert.ToBoolean(myReader["isparent"]);
            sClassTypeID   = Convert.ToInt32(myReader["classtypeid"]);
            sParentClassID = Convert.ToInt32(myReader["parentclassid"]);

            classes.updateClassTime(iTimeID,
                                    sCartQty,
                                    sBuyOrWait);

            sSQL2 = "DELETE FROM egov_class_cart_price WHERE cartid = " + iCartID.ToString();
            common.RunSQLStatement(sSQL2);

            sSQL2 = "DELETE FROM egov_class_cart WHERE cartid = " + iCartID.ToString();
            common.RunSQLStatement(sSQL2);

            sSQL2 = "DELETE FROM egov_class_cart_regattateams WHERE cartid = " + iCartID.ToString();
            common.RunSQLStatement(sSQL2);

            sSQL2 = "DELETE FROM egov_class_cart_regattateammembers WHERE cartid = " + iCartID.ToString();
            common.RunSQLStatement(sSQL2);

            classes.updateClassTimeSeriesChildren(sClassID,
                                                  sCartQty,
                                                  sBuyOrWait);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
    }

    public static string displayPublicBlockedNote(Int32 iUserID)
    {
        string lcl_return = "";
        string sSQL = "";

        sSQL  = "SELECT blockedexternalnote ";
        sSQL += " FROM egov_users ";
        sSQL += " WHERE userid = " + iUserID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToString(myReader["blockedexternalnote"]);

            if (lcl_return != "")
            {
                lcl_return = "<div>" + lcl_return + "</div>";
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static void getUserInfo(Int32 iUserID,
                                   out string sEmail,
                                   out string sAddress,
                                   out string sAddress2,
                                   out string sUserUnit,
                                   out string sCity,
                                   out string sState,
                                   out string sZip,
                                   out string sFirstName,
                                   out string sLastName,
                                   out string sName,
                                   out string sUserHomePhone)
    {
        string sSQL = "";

        sEmail = "";
        sAddress = "";
        sAddress2 = "";
        sUserUnit = "";
        sCity = "";
        sState = "";
        sZip = "";
        sFirstName = "";
        sLastName = "";
        sName = "";
        sUserHomePhone = "";

        sSQL  = "SELECT userfname, ";
        sSQL += " userlname, ";
        sSQL += " useraddress, ";
        sSQL += " useraddress2, ";
        sSQL += " userunit, ";
        sSQL += " usercity, ";
        sSQL += " userstate, ";
        sSQL += " userzip, ";
        sSQL += " useremail, ";
        sSQL += " userhomephone ";
        sSQL += " FROM egov_users ";
        sSQL += " WHERE userid = " + iUserID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            sEmail         = common.decodeUTFString(Convert.ToString(myReader["useremail"]));
            sAddress       = common.decodeUTFString(Convert.ToString(myReader["useraddress"]));
            sAddress2      = common.decodeUTFString(Convert.ToString(myReader["useraddress2"]));
            sUserUnit      = common.decodeUTFString(Convert.ToString(myReader["userunit"]));
            sCity          = common.decodeUTFString(Convert.ToString(myReader["usercity"]));
            sState         = common.decodeUTFString(Convert.ToString(myReader["userstate"]));
            sFirstName     = common.decodeUTFString(Convert.ToString(myReader["userfname"]));
            sLastName      = common.decodeUTFString(Convert.ToString(myReader["userlname"]));

            sZip           = Convert.ToString(myReader["userzip"]);
            sUserHomePhone = Convert.ToString(myReader["userhomephone"]);

            sName = sFirstName + " " + sLastName;
            sName = sName.Trim();
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
    }

    public static string getSerialNumber(Int32 iOrgID)
    {
        string lcl_return = "00000000";
        string sSQL = "";

        sSQL  = "SELECT serialnumber ";
        sSQL += " FROM egov_skipjackoptions ";
        sSQL += " WHERE orgid = " + iOrgID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToString(myReader["serialnumber"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getActivityNosForPayPalComment2(string iSessionID)
    {
        string lcl_return = "";
        string sSessionID = "";
        string sSQL       = "";

        if (iSessionID != "" && iSessionID != null)
        {
            sSessionID = common.dbSafe(iSessionID);
        }

        sSessionID = "'" + sSessionID + "'";

        sSQL  = "SELECT isnull(t.activityno,'') as activityno ";
        sSQL += " FROM egov_class_time t, ";
        sSQL +=      " egov_class_cart c ";
        sSQL += " WHERE t.timeid = c.classtimeid ";
        sSQL += " AND isregatta = 0 ";
        sSQL += " AND issinglefloater = 0 ";
        sSQL += " AND issalestax = 0 ";
        sSQL += " AND isshippingfee = 0 ";
        sSQL += " AND isnull(c.sessionid_csharp, c.sessionid) = " + sSessionID;
        sSQL += " ORDER BY t.activityno ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                if (Convert.ToString(myReader["activityno"]) != "")
                {
                    if (lcl_return != "")
                    {
                        lcl_return += ",";
                    }

                    lcl_return += Convert.ToString(myReader["activityno"]);
                }
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        if (lcl_return.Length > 128)
        {
            lcl_return = lcl_return.Substring(0, 128);  //PayPal allows 128 chars for this.
        }

        return lcl_return;
    }

    public static string showCartItems(string iSessionID)
    {
        string lcl_return        = "";
        string sSessionID        = "";
        string sSQL              = "";
        string sItemType         = "";
        string sActivityNo       = "";
        string sClassName        = "";
        string sStartDate        = "";
        string sFamilyMemberName = "";
        //string sShowMerchandiseItemsForVerisign = "";

        if (iSessionID != "")
        {
            sSessionID = common.dbSafe(iSessionID);
        }

        sSessionID = "'" + sSessionID + "'";

        sSQL  = "SELECT cc.cartid, ";
        sSQL += " cc.quantity, ";
        sSQL += " cc.optionid, ";
        sSQL += " c.classname, ";
        sSQL += " c.startdate, ";
        sSQL += " i.itemtype, ";
        sSQL += " isnull(cc.familymemberid,0) as familymemberid, ";
        sSQL += " isnull(cc.classtimeid,0) as classtimeid ";
        sSQL += " FROM egov_class_cart cc ";
        sSQL +=      " LEFT OUTER JOIN egov_class c ON cc.classid = c.classid ";
        sSQL +=      " LEFT OUTER JOIN egov_item_types i ON cc.itemtypeid = i.itemtypeid ";
        sSQL += " WHERE cc.sessionid_csharp = " + sSessionID;

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            lcl_return = "<ul>";

            while (myReader.Read())
            {
                sItemType         = Convert.ToString(myReader["itemtype"]).ToUpper();
                sClassName        = Convert.ToString(myReader["classname"]);
                sStartDate        = "";
                sActivityNo       = "";
                sFamilyMemberName = "";

                lcl_return += "<li>";

                switch (sItemType)
                {
                    case "RECREATION ACTIVITY":
                        sActivityNo = classes.getActivityNo(Convert.ToInt32(myReader["classtimeid"]));
                        sStartDate  = string.Format("{0:M/d/yyyy}", Convert.ToDateTime(myReader["startdate"]));

                        lcl_return += sClassName;
                        lcl_return += " (" + sActivityNo + ") ";
                        lcl_return += "<strong>on </strong>";
                        lcl_return += sStartDate;

                        if (Convert.ToInt32(myReader["optionid"]) == 2)
                        {
                            //Show quantity for ticket events
                            lcl_return += " <strong>Qty: </strong>" + Convert.ToString(myReader["quantity"]);
                        }
                        else
                        {
                            //Show family member name for registration events
                            sFamilyMemberName = classes.getFamilyMemberName(Convert.ToInt32(myReader["familymemberid"]));

                            lcl_return += "&nbsp;&nbsp;<strong>For: </strong>" + sFamilyMemberName;
                        }

                        break;
/*
                    case "MERCHANDISE":
                        sShowMerchandiseItemsForVerisign = classes.getMerchandiseItemsForVerisign(Convert.ToInt32(myReader["cartid"]));

                        lcl_return += "Merchandise";
                        lcl_return += sShowMerchandiseItemsForVerisign;

                        break;
*/
                }

                lcl_return += "</li>";

            }

            lcl_return += "</ul>";
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }
/*
    public static string getMerchandiseItemsForVerisign(Int32 iCartID)
    {
        string lcl_return = "";
        string sSQL = "";

        sSQL  = "SELECT m.merchandise, ";
        sSQL += " mc.merchandisecolor, ";
        sSQL += " mc.isnocolor, ";
        sSQL += " ms.merchandisesize, ";
        sSQL += " ms.isnosize, ";
        sSQL += " i.quantity, ";
        sSQL += " i.price ";
        sSQL += " FROM egov_class_cart_merchandiseitems i, ";
        sSQL +=      " egov_merchandisecatalog c, ";
        sSQL +=      " egov_merchandise m, ";
        sSQL +=      " egov_merchandisecolors mc, ";
        sSQL +=      " egov_merchandisesizes ms ";
        sSQL += " WHERE i.merchandisecatalogid = c.merchandisecatalogid ";
        sSQL += " AND c.merchandiseid = m.merchandiseid ";
        sSQL += " AND c.merchandisecolorid = mc.merchandisecolorid ";
        sSQL += " AND c.merchandisesizeid = ms.merchandisesizeid ";
        sSQL += " AND i.cartid = " + iCartID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                lcl_return += "<br />";
                lcl_return += Convert.ToString(myReader["quantity"]);
                lcl_return += "&nbsp;" + Convert.ToString(myReader["merchandise"]);

                if (!Convert.ToBoolean(myReader["isnocolor"]))
                {
                    lcl_return += ",&nbsp;" + Convert.ToString(myReader["merchandisecolor"]);
                }

                if (!Convert.ToBoolean(myReader["isnosize"]))
                {
                    lcl_return += ",&nbsp;" + Convert.ToString(myReader["merchandisesize"]);
                }
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }
*/
    public static string getUserContactInfo(Int32 iUserID,
                                            string iColumnName)
    {
        string lcl_return = "";
        string sSQL = "";
        string sColumnName = "";

        if(iColumnName != null) {
            sColumnName = common.dbSafe(iColumnName);
        }

        if (iUserID != null && sColumnName != "")
        {
            sSQL  = "SELECT isnull(" + sColumnName + ", '') as dbcolumn ";
            sSQL += " FROM egov_users ";
            sSQL += " WHERE userid = " + iUserID.ToString();

            SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
            sqlConn.Open();

            SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
            SqlDataReader myReader;
            myReader = myCommand.ExecuteReader();

            if (myReader.HasRows)
            {
                myReader.Read();

                lcl_return = Convert.ToString(myReader["dbcolumn"]);
            }

            myReader.Close();
            sqlConn.Close();
            myReader.Dispose();
            sqlConn.Dispose();
        }

        return lcl_return;
    }

    public static string getAdminName(Int32 iUserID)
    {
        string lcl_return = "";
        string sSQL = "";

        sSQL  = "SELECT firstname + ' ' + lastname as username ";
        sSQL += " FROM users ";
        sSQL += " WHERE userid = " + iUserID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToString(myReader["username"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getAdminLocation(Int32 iAdminLocationID)
    {
        string lcl_return = "";
        string sSQL = "";

        sSQL  = "SELECT name ";
        sSQL += " FROM egov_class_location ";
        sSQL += " WHERE locationid = " + iAdminLocationID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToString(myReader["name"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getResidentTypeDesc(string iUserType)
    {
        string lcl_return = "";
        string sSQL = "";
        string sUserType = "";

        if(iUserType != "")
        {
            sUserType = common.dbSafe(iUserType);
        }

        sUserType = "'" + sUserType + "'";

        sSQL  = "SELECT description ";
        sSQL += " FROM egov_poolpassresidenttypes ";
        sSQL += " WHERE resident_type = " + sUserType;

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToString(myReader["description"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getFamilyEmail(Int32 iUserID)
    {
        string lcl_return = "";
        string sSQL = "";

        sSQL = "SELECT useremail ";
        sSQL += " FROM egov_users ";
        sSQL += " WHERE useremail IS NOT NULL ";
        sSQL += " AND familyid = " + iUserID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToString(myReader["useremail"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getRefundName(Int32 iOrgID)
    {
        string lcl_return = "Refund Voucher";
        string sSQL = "";

        sSQL  = "SELECT t.paymenttypename ";
        sSQL += " FROM egov_paymenttypes t, ";
        sSQL += " egov_organizations_to_paymenttypes o ";
        sSQL += " WHERE t.isrefundmethod = 1 ";
        sSQL += " AND t.paymenttypeid = o.paymenttypeid ";
        sSQL += " AND o.orgid = " + iOrgID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToString(myReader["paymenttypename"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static double getLedgerAmount(Int32 iPaymentID,
                                         Int32 iPaymentTypeID)
    {
        double lcl_return = 0.00;

        string sSQL = "";

        sSQL  = "SELECT amount ";
        sSQL += " FROM egov_accounts_ledger ";
        sSQL += " WHERE ispaymentaccount = 1 ";
        sSQL += " AND paymentid = " + iPaymentID.ToString();
        sSQL += " AND paymenttypeid = " + iPaymentTypeID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToDouble(myReader["amount"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getCheckNo(Int32 iPaymentID,
                                    Int32 iPaymentTypeID)
    {
        string lcl_return = "";
        string sSQL = "";

        sSQL  = "SELECT checkno ";
        sSQL += " FROM egov_verisign_payment_information ";
        sSQL += " WHERE paymentid = " + iPaymentID.ToString();
        sSQL += " AND paymenttypeid = " + iPaymentTypeID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToString(myReader["checkno"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getAccountName(Int32 iPaymentID,
                                        Int32 iPaymentTypeID)
    {
        string lcl_return = "";
        string sSQL = "";

        sSQL  = "SELECT userfname, ";
        sSQL += " userlname ";
        sSQL += " FROM egov_verisign_payment_information, ";
        sSQL +=      " egov_users ";
        sSQL += " WHERE paymentid = " + iPaymentID.ToString();
        sSQL += " AND paymenttypeid = " + iPaymentTypeID.ToString();
        sSQL += " AND citizenuserid = userid";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToString(myReader["userfname"]) + " " + Convert.ToString(myReader["userlname"]);
            lcl_return = lcl_return.Trim();
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string showReceiptTransactions(Int32 iOrgID,
                                                 Int32 iPaymentID,
                                                 string iJournalEntryType,
                                                 Boolean iHasPaymentFee,
                                                 double iProcessingFee)
    {
        string lcl_return = "";
        string sEntryType = "";
        string sJournalEntryType = "";

        if(iJournalEntryType != "")
        {
            sJournalEntryType = common.dbSafe(iJournalEntryType);
            sJournalEntryType = sJournalEntryType.ToUpper();
        }

        switch (sJournalEntryType)
        {
            case "PURCHASE":  //Show purchase details
                sEntryType = "credit";
                break;

            case "REFUND":  //Show refund details
                sEntryType = "debit";
                break;

            case "TRANSFER":  //Show citizen account transfer
                break;

            case "DEPOSIT":  //Show citizen account deposit
                break;

            case "WITHDRAWAL":  //Show citizen account withdrawal
                break;
        }

        if (sEntryType != "")
        {
            lcl_return = showPurchaseDetails(iOrgID,
                                             iPaymentID,
                                             sEntryType,
                                             sJournalEntryType,
                                             iHasPaymentFee,
                                             iProcessingFee);
        }

        return lcl_return;
    }

    public static string showPurchaseDetails(Int32 iOrgID,
                                             Int32 iPaymentID,
                                             string iEntryType,
                                             string iJournalEntryType,
                                             Boolean iHasPaymentFee,
                                             double iProcessingFee)
    {
        double sTotal = 0.00;
        double sAmount = 0.00;

        Int32 sItemID     = 0;
        Int32 sItemTypeID = 0;

        string lcl_return = "";
        string sSQL       = "";
        string sEntryType = "";
        string sItemType  = "";
        string sRefundDetails = "";
        string sLabelTotal = "Total";

        if(iEntryType != "")
        {
            sEntryType = common.dbSafe(iEntryType);
            sEntryType = "'" + sEntryType + "'";
        }

        sSQL  = "SELECT t.cartdisplayorder, ";
        sSQL += " itemtype, ";
        sSQL += " itemid, ";
        sSQL += " L.itemtypeid, ";
        sSQL += " sum(amount) as amount ";
        sSQL += " FROM egov_accounts_ledger L, ";
        sSQL +=      " egov_item_types t ";
        sSQL += " WHERE L.itemtypeid = t.itemtypeid ";
        sSQL += " AND L.ispaymentaccount = 0 ";
        sSQL += " AND entrytype = " + sEntryType;
        sSQL += " AND L.paymentid = " + iPaymentID.ToString();
        sSQL += " GROUP BY t.cartdisplayorder, itemtype, itemid, L.itemtypeid ";
        sSQL += " ORDER BY t.cartdisplayorder, itemtype, itemid, L.itemtypeid ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                sAmount = Convert.ToDouble(myReader["amount"]);

                sItemID = Convert.ToInt32(myReader["itemid"]);
                sItemTypeID = Convert.ToInt32(myReader["itemtypeid"]);

                sItemType = Convert.ToString(myReader["itemtype"]);
                sItemType = sItemType.ToUpper();

                switch (sItemType)
                {
                    case "RECREATION ACTIVITY":
                        sTotal = sTotal + sAmount;

                        lcl_return += "<fieldset class=\"fieldset\">";
                        lcl_return += showActivityDetails(iOrgID,
                                                         sItemID,
                                                         sAmount,
                                                         iPaymentID,
                                                         sItemTypeID);
                        lcl_return += "</fieldset>";

                        break;
                }
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        //BEGIN: Refund/Processing Fees -----------------------------
        if (iJournalEntryType.ToUpper() == "REFUND")
        {
            //Refund Fees
            sTotal = sTotal - getRefundFee(iPaymentID,
                                           sTotal,
                                           out sRefundDetails);

            if (sRefundDetails != "")
            {
                lcl_return += sRefundDetails;
            }

            sLabelTotal = "Refund Total";
        }
        else
        {
            //Processing Fees
            if (iHasPaymentFee)
            {
                //Add the processing fee to the total charged.
                sTotal = sTotal + iProcessingFee;

                lcl_return += "<fieldset class=\"fieldset\">";
                lcl_return += "<table border=\"0\" width=\"100%\">";
                lcl_return += "  <tr>";
                lcl_return += "      <td class=\"receiptLabel\">Processing Fee</td>";
                lcl_return += "      <td class=\"receiptSubTotal\">" + string.Format("{0:C}", iProcessingFee) + "</td>";
                lcl_return += "  </tr>";
                lcl_return += "</table>";
                lcl_return += "</fieldset>";
            }
        }
        //END: Refund/Processing Fees -------------------------------

        //BEGIN: Total ----------------------------------------------
        lcl_return += "<fieldset class=\"fieldset\">";
        lcl_return += "<table border=\"0\" width=\"100%\">";
        lcl_return += "  <tr>";
        lcl_return += "      <td class=\"receiptTotalLabel\">" + sLabelTotal + "</td>";
        lcl_return += "      <td class=\"receiptTotal\">" + string.Format("{0:C}", sTotal) + "</td>";
        lcl_return += "  </tr>";
        lcl_return += "</table>";
        lcl_return += "</fieldset>";
        //END: Total ------------------------------------------------
        
        return lcl_return;
    }

    public static string showActivityDetails(Int32 iOrgID,
                                             Int32 iItemID,
                                             double iSubTotal,
                                             Int32 iPaymentID,
                                             Int32 iItemTypeID)
    {
        Boolean sIsSeriesChild = false;
        Boolean sIsParent  = false;
        Boolean sIsDropIn  = false;
        Boolean sSunday = false;
        Boolean sMonday = false;
        Boolean sTuesday = false;
        Boolean sWednesday = false;
        Boolean sThursday = false;
        Boolean sFriday = false;
        Boolean sSaturday = false;

        Int32 sQuantity = 0;
        Int32 sClassID = 0;

        string lcl_return = "";
        string sSQL = "";
        string sClassName = "";
        string sActivityNo = "";
        string sLocationName = "";
        string sAddress1 = "";
        string sDisplaySubTotal = "";
        string sUserFirstName = "";
        string sUserLastName = "";
        string sUserName = "";
        string sJournalItemStatus = "";
        string sStartDate = "";
        DateTime? dStartDate = null;
        string sStartTime = "";
        string sEndDate = "";
        string sEndTime = "";        
        string sDaysOfWeek = "";
        string sDropInDate = "";
        string sRosterGrade = "";
        string sRosterShirtSize = "";
        string sRosterPantsSize = "";
        string sRosterCoachType = "";
        string sRosterVolunteerCoachName = "";
        string sRosterVolunteerCoachDayPhone = "";
        string sRosterVolunteerCoachCellPhone = "";
        string sRosterVolunteerCoachEmail = "";
        string sNotes = "";
        string sShowSeriesChildrenDetails = "";
	string sQueryString = "";
	bool bNoRefunds = false;

        sSQL  = "SELECT u.userfname, ";
        sSQL += " u.userlname, ";
        sSQL += " status, ";
        sSQL += " quantity, ";
        sSQL += " isdropin, ";
        sSQL += " dropindate, ";
        sSQL += " c.classid, ";
        sSQL += " classname, ";
        sSQL += " c.isparent, ";
        sSQL += " cl.name, ";
        sSQL += " cl.address1, ";
        sSQL += " c.startdate, ";
        sSQL += " c.enddate, ";
        sSQL += " notes, ";
        sSQL += " t.activityno, ";
        sSQL += " sunday, ";
        sSQL += " monday, ";
        sSQL += " tuesday, ";
        sSQL += " wednesday, ";
        sSQL += " thursday, ";
        sSQL += " friday, ";
        sSQL += " saturday, ";
        sSQL += " td.starttime, ";
        sSQL += " td.endtime, ";
        sSQL += " l.rostergrade, ";
        sSQL += " l.rostershirtsize, ";
        sSQL += " l.rosterpantssize, ";
        sSQL += " l.rostercoachtype, ";
        sSQL += " l.rostervolunteercoachname, ";
        sSQL += " l.rostervolunteercoachdayphone, ";
        sSQL += " l.rostervolunteercoachcellphone, ";
        sSQL += " l.rostervolunteercoachemail, ";
        sSQL += " t.timeid, ";
        sSQL += " c.norefunds ";
        sSQL += " FROM egov_class_list l, ";
        sSQL += " egov_class c, ";
        sSQL += " egov_class_time t, ";
        sSQL += " egov_class_time_days td, ";
        sSQL += " egov_class_location cl, ";
        sSQL += " egov_users u ";
        sSQL += " WHERE c.classid = l.classid ";
        sSQL += " AND l.classtimeid = t.timeid ";
        sSQL += " AND l.attendeeuserid = u.userid ";
        sSQL += " AND t.timeid = td.timeid ";
        sSQL += " AND c.locationid = cl.locationid ";
        sSQL += " AND l.classlistid = " + iItemID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            sIsDropIn = Convert.ToBoolean(myReader["isdropin"]);
            sIsParent = Convert.ToBoolean(myReader["isparent"]);
            bNoRefunds = Convert.ToBoolean(myReader["norefunds"]);

            sQuantity = Convert.ToInt32(myReader["quantity"]);

	    sQueryString = "classid=" + myReader["classid"] + "&timeid=" + myReader["timeid"] + "&classlistid=" + iItemID.ToString() + "&paymentid=" + iPaymentID + "&iqty=1";
            
            sClassID       = Convert.ToInt32(myReader["classid"]);
            sClassName     = common.decodeUTFString(Convert.ToString(myReader["classname"]));
            sActivityNo    = common.decodeUTFString(Convert.ToString(myReader["activityno"]));
            sLocationName  = common.decodeUTFString(Convert.ToString(myReader["name"]));
            sAddress1      = common.decodeUTFString(Convert.ToString(myReader["address1"]));
            sUserFirstName = common.decodeUTFString(Convert.ToString(myReader["userfname"]));
            sUserLastName  = common.decodeUTFString(Convert.ToString(myReader["userlname"]));
            sNotes         = common.decodeUTFString(Convert.ToString(myReader["notes"]));

            sUserName = sUserFirstName + " " + sUserLastName;
            sUserName = sUserName.Trim();

            sDisplaySubTotal = string.Format("{0:C}", iSubTotal);

            sJournalItemStatus = classes.getJournalItemStatus(iPaymentID,
                                                              iItemTypeID,
                                                              iItemID);
            if (! sIsDropIn)
            {
                sSunday    = Convert.ToBoolean(myReader["sunday"]);
                sMonday    = Convert.ToBoolean(myReader["monday"]);
                sTuesday   = Convert.ToBoolean(myReader["tuesday"]);
                sWednesday = Convert.ToBoolean(myReader["wednesday"]);
                sThursday  = Convert.ToBoolean(myReader["thursday"]);
                sFriday    = Convert.ToBoolean(myReader["friday"]);
                sSaturday  = Convert.ToBoolean(myReader["saturday"]);

                sStartTime = Convert.ToString(myReader["starttime"]);
                sEndTime   = Convert.ToString(myReader["endtime"]);

                sDaysOfWeek = buildDaysOfWeek(sSunday,
                                              sMonday,
                                              sTuesday,
                                              sWednesday,
                                              sThursday,
                                              sFriday,
                                              sSaturday);

                if (sDaysOfWeek != "")
                {
		    dStartDate = Convert.ToDateTime(myReader["startdate"]);
                    sStartDate = string.Format("{0:MM/dd/yyyy}", dStartDate);
                    sEndDate   = string.Format("{0:MM/dd/yyyy}", Convert.ToDateTime(myReader["enddate"]));

                    sDaysOfWeek = "<span class=\"receiptLabel\">From: </span>" + sStartDate + "<span class=\"receiptLabel\">&nbsp;&nbsp;&nbsp;To: </span>" + sEndDate + "<br /><span class=\"receiptLabel\">Days: </span>" + sDaysOfWeek;
                    sDaysOfWeek += "&nbsp;&nbsp;-&nbsp;&nbsp;";
                    //sDaysOfWeek += "<br />";
                    sDaysOfWeek += "<span class=\"receiptLabel\">Times: </span>" + sStartTime + "<span class=\"receiptLabel\"> to </span>" + sEndTime + "<br />";
                }

                while (myReader.Read())
                {
                    sSunday    = Convert.ToBoolean(myReader["sunday"]);
                    sMonday    = Convert.ToBoolean(myReader["monday"]);
                    sTuesday   = Convert.ToBoolean(myReader["tuesday"]);
                    sWednesday = Convert.ToBoolean(myReader["wednesday"]);
                    sThursday  = Convert.ToBoolean(myReader["thursday"]);
                    sFriday    = Convert.ToBoolean(myReader["friday"]);
                    sSaturday  = Convert.ToBoolean(myReader["saturday"]);

                    sStartTime = Convert.ToString(myReader["starttime"]);
                    sEndTime   = Convert.ToString(myReader["endtime"]);

                    string sDaysOfWeekLoop = buildDaysOfWeek(sSunday,
                                                  sMonday,
                                                  sTuesday,
                                                  sWednesday,
                                                  sThursday,
                                                  sFriday,
                                                  sSaturday);

                    if (sDaysOfWeekLoop != "")
                    {
                        sDaysOfWeek += "<span class=\"receiptLabel\">Days: </span>" + sDaysOfWeekLoop;
                        sDaysOfWeek += "&nbsp;&nbsp;-&nbsp;&nbsp;";
                        //sDaysOfWeek += "<br />";
                    	sDaysOfWeek += "<span class=\"receiptLabel\">Times: </span>" + sStartTime + "<span class=\"receiptLabel\"> to </span>" + sEndTime + "<br />";
                    }
                }

            }
            else
            {
                sDropInDate = string.Format("{0:MM/dd/yyyy}", Convert.ToDateTime(myReader["dropindate"]));
                sDaysOfWeek = "<span class=\"receiptLabel\">Drop In Date: </span>" + sDropInDate;
            }

            if (sIsParent)
            {
                sShowSeriesChildrenDetails = classes.showSeriesChildren(iOrgID,
                                                                        sClassID);
            }

            lcl_return = showReceiptTransactionClassInfo(iOrgID,
                                                         sIsSeriesChild,
                                                         sClassName,
                                                         sActivityNo,
                                                         sLocationName,
                                                         sAddress1,
                                                         sDisplaySubTotal,
                                                         sUserName,
                                                         sJournalItemStatus,
                                                         Convert.ToString(sQuantity),
                                                         sDaysOfWeek,
                                                         sShowSeriesChildrenDetails,
                                                         sNotes,
                                                         sRosterGrade,
                                                         sRosterShirtSize,
                                                         sRosterPantsSize,
                                                         sRosterCoachType,
                                                         sRosterVolunteerCoachName,
                                                         sRosterVolunteerCoachDayPhone,
                                                         sRosterVolunteerCoachCellPhone,
                                                         sRosterVolunteerCoachEmail,
							 dStartDate,
							 sQueryString,
							 bNoRefunds, iItemID);
            /*
            lcl_return  = "<table border=\"1\" width=\"100%\">";
            lcl_return += "  <tr>";
            lcl_return += "      <td><span class=\"receiptLabel\">Activity: </span>" + sClassName + " (" + sActivityNo + ")</td>";
            lcl_return += "      <td><span class=\"receiptLabel\">Location: </span>" + sLocationName + "&nbsp;-&nbsp;" + sAddress1 + "</td>";
            lcl_return += "      <td align=\"right\">" + sDisplaySubTotal + "</td>";
            lcl_return += "  </tr>";
            lcl_return += "  <tr valign=\"top\">";
            lcl_return += "      <td>";
            lcl_return += "          <span class=\"receiptLabel\">Attendee: </span>" + sUserName + "&nbsp;-&nbsp;" + sJournalItemStatus + "<br />";
            lcl_return += "          <span class=\"receiptLabel\">Qty: </span>" + sQuantity.ToString();
            lcl_return += "      </td>";
            lcl_return += "      <td>" + sDaysOfWeek + "</td>";
            lcl_return += "      <td>&nbsp;</td>";
            lcl_return += "  </tr>";
            lcl_return += sShowSeriesChildrenDetails;
            lcl_return += "</table>";
            */
        }



        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string showSeriesChildren(Int32 iOrgID,
                                            Int32 iParentClassID)
    {
        Boolean sIsSeriesChild = true;
        Boolean sSunday = false;
        Boolean sMonday = false;
        Boolean sTuesday = false;
        Boolean sWednesday = false;
        Boolean sThursday = false;
        Boolean sFriday = false;
        Boolean sSaturday = false;

        Int32 sQuantity = 0;

        string lcl_return = "";
        string sSQL = "";
        string sClassName = "";
        string sActivityNo = "";
        string sLocationName = "";
        string sAddress1 = "";
        string sDisplaySubTotal = "";
        string sUserName = "";
        string sJournalItemStatus = "";
        string sDaysOfWeek = "";
        string sShowSeriesChildrenDetails = "";
        string sStartDate = "";
        string sStartTime = "";
        string sEndDate = "";
        string sEndTime = "";
        string sNotes = "";
        string sRosterGrade = "";
        string sRosterShirtSize = "";
        string sRosterPantsSize = "";
        string sRosterCoachType = "";
        string sRosterVolunteerCoachName = "";
        string sRosterVolunteerCoachDayPhone = "";
        string sRosterVolunteerCoachCellPhone = "";
        string sRosterVolunteerCoachEmail = "";

        sSQL  = "SELECT c.classid, ";
        sSQL += " classname, ";
        sSQL += " cl.name, ";
        sSQL += " cl.address1, ";
        sSQL += " c.startdate, ";
        sSQL += " c.enddate, ";
        sSQL += " notes, ";
        sSQL += " t.activityno, ";
        sSQL += " sunday, ";
        sSQL += " monday, ";
        sSQL += " tuesday, ";
        sSQL += " wednesday, ";
        sSQL += " thursday, ";
        sSQL += " friday, ";
        sSQL += " saturday, ";
        sSQL += " td.starttime, ";
        sSQL += " td.endtime ";
        sSQL += " FROM egov_class c, ";
        sSQL += " egov_class_time t, ";
        sSQL += " egov_class_time_days td, ";
        sSQL += " egov_class_location cl ";
        sSQL += " WHERE c.classid = t.classid ";
        sSQL += " AND t.timeid = td.timeid ";
        sSQL += " AND c.locationid = cl.locationid ";
        sSQL += " AND parentclassid = " + iParentClassID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {

            lcl_return = "<tr>";
            lcl_return += "    <td valign=\"top\" colspan=\"3\">";
            lcl_return += "        <fieldset class=\"fieldset\">";
            lcl_return += "          <legend>Series Includes:</legend>";
            lcl_return += "          <table>";

            while (myReader.Read())
            {
                sQuantity = Convert.ToInt32(myReader["quantity"]);

                sClassName    = common.decodeUTFString(Convert.ToString(myReader["classname"]));
                sActivityNo   = common.decodeUTFString(Convert.ToString(myReader["activityno"]));
                sLocationName = common.decodeUTFString(Convert.ToString(myReader["name"]));
                sAddress1     = common.decodeUTFString(Convert.ToString(myReader["address1"]));

                sSunday    = Convert.ToBoolean(myReader["sunday"]);
                sMonday    = Convert.ToBoolean(myReader["monday"]);
                sTuesday   = Convert.ToBoolean(myReader["tuesday"]);
                sWednesday = Convert.ToBoolean(myReader["wednesday"]);
                sThursday  = Convert.ToBoolean(myReader["thursday"]);
                sFriday    = Convert.ToBoolean(myReader["friday"]);
                sSaturday  = Convert.ToBoolean(myReader["saturday"]);

                sDaysOfWeek = buildDaysOfWeek(sSunday,
                                              sMonday,
                                              sTuesday,
                                              sWednesday,
                                              sThursday,
                                              sFriday,
                                              sSaturday);

                if (sDaysOfWeek != "")
                {
                    sStartDate = string.Format("{0:MM/dd/yyyy}", Convert.ToDateTime(myReader["startdate"]));
                    sEndDate   = string.Format("{0:MM/dd/yyyy}", Convert.ToDateTime(myReader["enddate"]));

                    sStartTime = Convert.ToString(myReader["starttime"]);
                    sEndTime   = Convert.ToString(myReader["endtime"]);

                    sDaysOfWeek  = "<span class=\"receiptLabel\">From: </span>" + sStartDate + "<span class=\"receiptLabel\">&nbsp;&nbsp;&nbsp;To: </span>" + sEndDate + "<br /><span class=\"receiptLabel\">Days: </span>" + sDaysOfWeek;
                    //sDaysOfWeek += "&nbsp;&nbsp;-&nbsp;&nbsp;";
                    sDaysOfWeek += "<br />";
                    sDaysOfWeek += "<span class=\"receiptLabel\">Times: </span>" + sStartTime + "<span class=\"receiptLabel\"> to </span>" + sEndTime + "<br />";
                }

                lcl_return += showReceiptTransactionClassInfo(iOrgID,
                                                              sIsSeriesChild,
                                                              sClassName,
                                                              sActivityNo,
                                                              sLocationName,
                                                              sAddress1,
                                                              sDisplaySubTotal,
                                                              sUserName,
                                                              sJournalItemStatus,
                                                              Convert.ToString(sQuantity),
                                                              sDaysOfWeek,
                                                              sShowSeriesChildrenDetails,
                                                              sNotes,
                                                              sRosterGrade,
                                                              sRosterShirtSize,
                                                              sRosterPantsSize,
                                                              sRosterCoachType,
                                                              sRosterVolunteerCoachName,
                                                              sRosterVolunteerCoachDayPhone,
                                                              sRosterVolunteerCoachCellPhone,
                                                              sRosterVolunteerCoachEmail,null,null,false,0);
                /*
                lcl_return += "  <tr>";
                lcl_return += "      <td><span class=\"receiptLabel\">Activity: </span>" + sClassName + " (" + sActivityNo + ")</td>";
                lcl_return += "      <td><span class=\"receiptLabel\">Location: </span>" + sLocationName + "&nbsp;-&nbsp;" + sAddress1 + "</td>";
                lcl_return += "<td>&nbsp;</td>";
                //lcl_return += "      <td align=\"right\">" + sDisplaySubTotal + "</td>";
                lcl_return += "  </tr>";
                lcl_return += "  <tr valign=\"top\">";
                //lcl_return += "      <td>";
                //lcl_return += "          <span class=\"receiptLabel\">Attendee: </span>" + sUserName + "&nbsp;-&nbsp;" + sJournalItemStatus + "<br />";
                //lcl_return += "          <span class=\"receiptLabel\">Qty: </span>" + sQuantity.ToString();
                //lcl_return += "      </td>";
                lcl_return += "      <td>" + sDaysOfWeek + "</td>";
                lcl_return += "      <td>&nbsp;</td>";
                lcl_return += "  </tr>";
                */
            }

            lcl_return += "          </table>";
            lcl_return += "        </fieldset>";
            lcl_return += "    </td>";
            lcl_return += "</tr>";
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string showReceiptTransactionClassInfo(Int32 iOrgID,
                                                         Boolean iIsSeriesChild,
                                                         string iClassName,
                                                         string iActivityNo,
                                                         string iLocationName,
                                                         string iAddress1,
                                                         string iDisplaySubTotal,
                                                         string iUserName,
                                                         string iJournalItemStatus,
                                                         string iQuantity,
                                                         string iDaysOfWeek,
                                                         string iShowSeriesChildrenDetails,
                                                         string iNotes,
                                                         string iRosterGrade,
                                                         string iRosterShirtSize,
                                                         string iRosterPantsSize,
                                                         string iRosterCoachType,
                                                         string iRosterVolunteerCoachName,
                                                         string iRosterVolunteerCoachDayPhone,
                                                         string iRosterVolunteerCoachCellPhone,
                                                         string iRosterVolunteerCoachEmail,
							 DateTime? dStartDate,
							 string sQueryString,
							 bool bNoRefunds,
							 int iClassListID)
    {
        Boolean sIsSeriesChild = false;
        Boolean sRosterInfoExists = false;
        Boolean sOrgHasFeatureCustomRegistrationCraigCO = common.orgHasFeature(Convert.ToString(iOrgID), "custom_registration_CraigCO");
        Boolean sOrgHasDisplayClassTeamRegistrationTshirtLabel = false;

        string lcl_return = "";
        string sLabelTshirt = "T-Shirt";
        string sDisplayRosterGrade = "";
        string sDisplayRosterTshirt = "";
        string sDisplayRosterPants = "";
        string sDisplayRosterCoachType = "";
        string sDisplayRosterVolunteerCoachName = "";
        string sDisplayRosterVolunteerCoachDayPhone = "";
        string sDisplayRosterVolunteerCoachCellPhone = "";
        string sDisplayRosterVolunteerCoachEmail = "";

        try
        {
            sIsSeriesChild = Convert.ToBoolean(iIsSeriesChild);
        }
        catch
        {
            sIsSeriesChild = false;
        }

        //NOTE: Since layouts used between the parent and the series children are pretty much the same thing,
        //we simply need to determine which one to show if/when it is a "serieschild" or not.
        if (sIsSeriesChild)
        {
            lcl_return += "  <tr class=\"receiptTransactionsRow\" valign=\"top\">";
            lcl_return += "      <td><span class=\"receiptLabel\">Activity: </span>" + iClassName + " (" + iActivityNo + ")</td>";
            lcl_return += "      <td>";
            lcl_return += "          <span class=\"receiptLabel\">Location: </span>" + iLocationName + "&nbsp;-&nbsp;" + iAddress1 + "<br />";
            lcl_return +=            iDaysOfWeek;
            lcl_return += "      </td>";
            lcl_return += "      <td>&nbsp;</td>";
            lcl_return += "  </tr>";
        }
        else
        {
            lcl_return  = "<table width=\"100%\" class=\"receiptTransactionsTable\">";
            lcl_return += "  <tr valign=\"top\">";
            lcl_return += "      <td><span class=\"receiptLabel\">Activity: </span>" + iClassName + " (" + iActivityNo + ")</td>";
            lcl_return += "      <td><span class=\"receiptLabel\">Location: </span>" + iLocationName + "&nbsp;-&nbsp;" + iAddress1 + "</td>";
            lcl_return += "      <td class=\"receiptSubTotal\">" + iDisplaySubTotal + "</td>";
            lcl_return += "  </tr>";
            lcl_return += "  <tr valign=\"top\">";
            lcl_return += "      <td>";
            lcl_return += "          <span class=\"receiptLabel\">Attendee: </span><span class=\"receiptAttendee\">" + iUserName + "</span>&nbsp;-&nbsp;" + iJournalItemStatus + "<br />";
            lcl_return += "          <span class=\"receiptLabel\">Qty: </span>" + iQuantity;
            lcl_return += "      </td>";
            lcl_return += "      <td>" + iDaysOfWeek + "</td>";
            lcl_return += "      <td>&nbsp;</td>";
            lcl_return += "  </tr>";
            lcl_return +=    iShowSeriesChildrenDetails;
            lcl_return += "</table>";

            //Determine if the org is using the "custom roster" info.
            if (sOrgHasFeatureCustomRegistrationCraigCO)
            {
                if (iRosterGrade != "")
                {
                    //sDisplayRosterGrade = "<span class=\"receiptLabel\">Grade: </span>" + iRosterGrade + "<br />";
                    sRosterInfoExists = true;

                    sDisplayRosterGrade = "<tr>";
                    sDisplayRosterGrade += "    <td class=\"receiptLabel\">Grade: </td>";
                    sDisplayRosterGrade += "    <td>" + iRosterGrade + "</td>";
                    sDisplayRosterGrade += "</tr>";
                }

                if (iRosterShirtSize != "" && iRosterShirtSize != ",")
                {
                    sOrgHasDisplayClassTeamRegistrationTshirtLabel = common.orgHasDisplay(Convert.ToString(iOrgID), "class_teamregistration_tshirt_label");

                    if (sOrgHasDisplayClassTeamRegistrationTshirtLabel)
                    {
                        sLabelTshirt = common.getOrgDisplay(Convert.ToString(iOrgID), "class_teamregistration_tshirt_label");
                    }

                    //sDisplayRosterTshirt = "<span class=\"receiptLabel\">" + sLabelTshirt + " Size: </span>" + iRosterShirtSize + "<br />";
                    sRosterInfoExists = true;

                    sDisplayRosterTshirt = "<tr>";
                    sDisplayRosterTshirt += "    <td class=\"receiptLabel\">" + sLabelTshirt + " Size: </td>";
                    sDisplayRosterTshirt += "    <td>" + iRosterShirtSize + "</td>";
                    sDisplayRosterTshirt += "</tr>";
                }

                if (iRosterPantsSize != "" && iRosterPantsSize != ",")
                {
                    //sDisplayRosterPants = "<span class=\"receiptLabel\">Pants Size: </span>" + iRosterPantsSize + "<br />";
                    sRosterInfoExists = true;

                    sDisplayRosterPants = "<tr>";
                    sDisplayRosterTshirt += "    <td class=\"receiptLabel\">Pants Size: </td>";
                    sDisplayRosterTshirt += "    <td>" + iRosterPantsSize + "</td>";
                    sDisplayRosterTshirt += "</tr>";
                }

                if (iRosterCoachType != "")
                {
                    //sDisplayRosterCoachType = "<span class=\"receiptLabel\">Would like to be: </span>" + iRosterCoachType + "<br />";
                    sRosterInfoExists = true;

                    sDisplayRosterCoachType = "<tr>";
                    sDisplayRosterCoachType += "    <td class=\"receiptLabel\">Would like to be: </td>";
                    sDisplayRosterCoachType += "    <td>" + iRosterCoachType + "</td>";
                    sDisplayRosterCoachType += "</tr>";
                }

                if (iRosterVolunteerCoachName != "")
                {
                    //sDisplayRosterVolunteerCoachName = "<span class=\"receiptLabel\">Coach Name: </span>" + iRosterVolunteerCoachName + "<br />";
                    sRosterInfoExists = true;

                    sDisplayRosterVolunteerCoachName = "<tr>";
                    sDisplayRosterVolunteerCoachName += "    <td class=\"receiptLabel\">Coach Name: </td>";
                    sDisplayRosterVolunteerCoachName += "    <td>" + iRosterVolunteerCoachName + "</td>";
                    sDisplayRosterVolunteerCoachName += "</tr>";
                }

                if (iRosterVolunteerCoachDayPhone != "")
                {
                    //sDisplayRosterVolunteerCoachDayPhone = "<span class=\"receiptLabel\">Day Phone: </span>" + iRosterVolunteerCoachDayPhone + "<br />";
                    sRosterInfoExists = true;

                    sDisplayRosterVolunteerCoachDayPhone = "<tr>";
                    sDisplayRosterVolunteerCoachDayPhone += "    <td class=\"receiptLabel\">Day Phone: </td>";
                    sDisplayRosterVolunteerCoachDayPhone += "    <td>" + iRosterVolunteerCoachDayPhone + "</td>";
                    sDisplayRosterVolunteerCoachDayPhone += "</tr>";
                }

                if (iRosterVolunteerCoachCellPhone != "")
                {
                    //sDisplayRosterVolunteerCoachCellPhone = "<span class=\"receiptLabel\">Cell Phone: </span>" + iRosterVolunteerCoachCellPhone + "<br />";
                    sRosterInfoExists = true;

                    sDisplayRosterVolunteerCoachCellPhone = "<tr>";
                    sDisplayRosterVolunteerCoachCellPhone += "    <td class=\"receiptLabel\">Cell Phone: </td>";
                    sDisplayRosterVolunteerCoachCellPhone += "    <td>" + iRosterVolunteerCoachCellPhone + "</td>";
                    sDisplayRosterVolunteerCoachCellPhone += "</tr>";
                }

                if (iRosterVolunteerCoachEmail != "")
                {
                    //sDisplayRosterVolunteerCoachEmail = "<span class=\"receiptLabel\">Email: </span>" + iRosterVolunteerCoachEmail + "<br />";
                    sRosterInfoExists = true;

                    sDisplayRosterVolunteerCoachEmail = "<tr>";
                    sDisplayRosterVolunteerCoachEmail += "    <td class=\"receiptLabel\">Email: </td>";
                    sDisplayRosterVolunteerCoachEmail += "    <td>" + iRosterVolunteerCoachEmail + "</td>";
                    sDisplayRosterVolunteerCoachEmail += "</tr>";
                }


                if (sRosterInfoExists)
                {
                    lcl_return += "<div class=\"receiptCustomRegistration\">";
                    lcl_return += "<table>";
                    lcl_return += sDisplayRosterGrade;
                    lcl_return += sDisplayRosterTshirt;
                    lcl_return += sDisplayRosterPants;
                    lcl_return += sDisplayRosterCoachType;
                    lcl_return += sDisplayRosterVolunteerCoachName;
                    lcl_return += sDisplayRosterVolunteerCoachDayPhone;
                    lcl_return += sDisplayRosterVolunteerCoachCellPhone;
                    lcl_return += sDisplayRosterVolunteerCoachEmail;
                    lcl_return += "</table>";
                    lcl_return += "</div>";
                }
            }

            if (iNotes != "")
            {
                lcl_return += "<div><span class=\"receiptLabel\">Activity Notes: </span>" + iNotes + "</div>";
            }

	    if (dStartDate != null)
	    {
		string twfStatus = getRegistrantStatus(iClassListID);
	    	if ((iOrgID == 37 && (twfStatus == "ACTIVE" || twfStatus == "DROPIN")) || (iOrgID == 60 && (twfStatus == "ACTIVE" || twfStatus == "DROPIN") && DateTime.Now.AddDays(2) <= dStartDate && !bNoRefunds))
	    	{
            		lcl_return += "<div><input type=\"button\" value=\"Process Refund\" onclick=\"location.href='drop_registrant_form.asp?" + sQueryString + "'\" />";
	    	}
	    }
        }

        return lcl_return;
    }

    public static string getJournalItemStatus(Int32 iPaymentID,
                                          Int32 iItemTypeID,
                                          Int32 iItemID)
    {
        string lcl_return = "";
        string sSQL = "";

        sSQL = "SELECT status ";
        sSQL += " FROM egov_journal_item_status ";
        sSQL += " WHERE paymentid = " + iPaymentID.ToString();
        sSQL += " AND itemtypeid = " + iItemTypeID.ToString();
        sSQL += " AND itemid = " + iItemID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToString(myReader["status"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string buildDaysOfWeek(Boolean iSunday,
                                         Boolean iMonday,
                                         Boolean iTuesday,
                                         Boolean iWednesday,
                                         Boolean iThursday,
                                         Boolean iFriday,
                                         Boolean iSaturday)
    {
        string lcl_return = "";

        if (iSunday)
        {
            lcl_return += "Su ";
        }

        if (iMonday)
        {
            lcl_return += "Mo ";
        }

        if (iTuesday)
        {
            lcl_return += "Tu ";
        }

        if (iWednesday)
        {
            lcl_return += "We ";
        }

        if (iThursday)
        {
            lcl_return += "Th ";
        }

        if (iFriday)
        {
            lcl_return += "Fr ";
        }

        if (iSaturday)
        {
            lcl_return += "Sa ";
        }

        return lcl_return;
    }

    public static double getRefundFee(Int32 iPaymentID,
                                      double iTotal,
                                      out string sRefundDetails)
    {
        double lcl_return = 0.00;
        double sRefundShortage = 0.00;
        double sRefundDebit = getRefundDebit(iPaymentID);

        string sSQL = "";

        sRefundDetails  = "";
        sRefundShortage = iTotal - sRefundDebit;

        sSQL  = "SELECT itemtype, ";
        sSQL += " itemid, ";
        sSQL += " amount ";
        sSQL += " FROM egov_accounts_ledger l, ";
        sSQL += " egov_item_types t, ";
        sSQL += " egov_paymenttypes p ";
        sSQL += " WHERE l.itemtypeid = t.itemtypeid ";
        sSQL += " AND l.ispaymentaccount = 1 ";
        sSQL += " AND entrytype = 'credit' ";
        sSQL += " AND p.isrefunddebit = 1 ";
        sSQL += " AND p.paymenttypeid = l.paymenttypeid ";
        sSQL += " AND l.paymentid = " + iPaymentID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToDouble(myReader["amount"]) + sRefundShortage;

            sRefundDetails  = "<fieldset class=\"fieldset\">";
            sRefundDetails += "<table border=\"0\" width=\"100%\">";
            sRefundDetails += "  <tr>";
            sRefundDetails += "      <td class=\"receiptLabel\">Refund Fees</td>";
            sRefundDetails += "      <td class=\"receiptSubTotal\">" + string.Format("{0:C}", lcl_return) + "</td>";
            sRefundDetails += "  </tr>";
            sRefundDetails += "</table>";
            sRefundDetails += "</fieldset>";
        }
        else
        {
            lcl_return = sRefundShortage;
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();        

        return lcl_return;
    }

    public static double getRefundDebit(Int32 iPaymentID)
    {
        double lcl_return = 0.00;

        string sSQL = "";

        sSQL  = "SELECT SUM(amount) as amount ";
        sSQL += " FROM egov_accounts_ledger ";
        sSQL += " WHERE ispaymentaccount = 0 ";
        sSQL += " AND entrytype = 'debit' ";
        sSQL += " AND paymentid = " + iPaymentID.ToString();
        sSQL += " GROUP BY paymentid ";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            lcl_return = Convert.ToDouble(myReader["amount"]);
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static Boolean genderRequirementsMet(Int32 iUserID, string ReqGender)
    {
        Boolean lcl_return                        = false;

        string sSQL            = "";


        sSQL  = "SELECT DISTINCT u.gender ";
        sSQL += " FROM egov_familymembers f ";
	sSQL += " INNER JOIN egov_users u ON u.userid = f.userid ";
        sSQL += " WHERE belongstouserid = " + iUserID.ToString();
        sSQL += " AND f.isdeleted = 0 ";
        
        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
		if (myReader["gender"].ToString() == ReqGender)
		{
                    lcl_return = true;
                    break;
		}
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }

    public static string getRegistrantStatus(Int32 iClassListID)
    {
        string lcl_return = "UNKNOWN";

        string sSQL            = "";

	sSQL = "SELECT status FROM egov_class_list WHERE classlistid = "  +  iClassListID;

        
        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
	       lcl_return = myReader["status"].ToString();
            }
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

        return lcl_return;
    }
}

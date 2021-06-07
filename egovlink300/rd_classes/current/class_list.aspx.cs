using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_classes_class_list : System.Web.UI.Page
{
    double startCounter = 0.00;

    static string sOrgID = common.getOrgId();
    static string sOrgName = common.getOrgName(sOrgID);
    //string sOrgVirtualSiteName = common.getOrgInfo(sOrgID, "orgVirtualSiteName");
    //string sPageTitle = "E-Gov Services " + sOrgName;
    //string lcl_isLoggedIn = "";
    //string lcl_checked_isLoggedInYes = "";
    //string lcl_checked_isLoggedInNo = "";

    //static Int32 iRootCategoryID = getFirstCategory(sOrgID);
    //Int32 sCategoryID = iRootCategoryID;

    //Boolean sViewPick = false;
    //Boolean sShowViewPicks = true;

    protected void Page_PreInit(object sender, EventArgs e)
    {
        // This is the earliest thing the page does, so set the start time here.
        startCounter = DateTime.Now.TimeOfDay.TotalSeconds;
    }

    protected void Page_PreRender(object sender, EventArgs e)
    {
        //THIS STILL NEEDS FIXED!!!!
        common.logThePageVisit(startCounter, "class_list.aspx", "public");
    }

    protected void Page_Load(object sender, EventArgs e)
    {
    }

    public void showMemberWarning(string iOrgID)
    {
        Int32 sOrgID = 0;
        string sOrgDisplay = "";

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

        if (common.orgHasDisplay(sOrgID.ToString(), "classdetailsnotice"))
        {
            sOrgDisplay = common.getOrgDisplay(sOrgID.ToString(), "classdetailsnotice");

            Response.Write("<div id=\"classdetailsnotice\">" + sOrgDisplay + "</div>");
        }
    }

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

    public void displaySubCategoryMenu(string iOrgID,
                                       Int32 iRootCategoryID,
                                       Boolean iShowViewPicks,
                                       Boolean iViewPick,
                                       Int32 iCategoryID)
    {
        string sSQL = "";
        Int32 sOrgID = 0;
        Int32 sLineCount = 0;

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

        sSQL = "SELECT categorytitle, ";
        sSQL += " subcategoryid, ";
        sSQL += " subcategorytitle ";
        sSQL += " FROM class_categories ";
        sSQL += " WHERE orgid = " + sOrgID;
        sSQL += " AND categoryid = " + iRootCategoryID;
        sSQL += " ORDER BY sequenceid, subcategorytitle";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                sLineCount = sLineCount + 1;

                if (sLineCount > 1)
                {
                    Response.Write("<span id=\"subcategorymenu_seperator\">|</span></li>");
                }
                else
                {
                    Response.Write("<div id=\"subcategorymenu_new\" style=\"display:none\">");
                    Response.Write("  <ul id=\"subcategorymenu_list\">");
                    Response.Write("    <li><a id=\"subcategorymenu_rootoption\" href=\"class_list.aspx?categoryid=" + iRootCategoryID.ToString() + "\">" + myReader["categorytitle"].ToString() + "</a></li>");
                }

                //Response.Write("    <li><a href=\"class_list.aspx?categoryid=" + myReader["subcategoryid"].ToString() + "\">" + myReader["subcategorytitle"].ToString() + "</a></li>");
                Response.Write("    <li><a href=\"class_list.aspx?categoryid=" + myReader["subcategoryid"].ToString() + "\">" + myReader["subcategorytitle"].ToString() + "</a>");
            }

            Response.Write("    </li>");
            Response.Write("  </ul>");
            Response.Write("</div>");
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
    }



    public void displayClassesSearchBox(Int32 iCategoryID)
    //public void displayClassesSearchBox(Boolean iShowViewPicks, Boolean iViewPick, Int32 iCategoryID)
    {

        //string lcl_selected_viewpick0 = "";
        //string lcl_selected_viewpick1 = "";

        Response.Write("<div id=\"classesSearchBox\">");
        Response.Write("<form name=\"frmSearch\" id=\"frmSearch\" method=\"post\" action=\"class_search_results.asp\">");
        Response.Write("  <input type=\"hidden\" name=\"categoryid\" id=\"categoryid\" value=\"" + iCategoryID.ToString() + "\" />");
        Response.Write("  <strong>Search: </strong>");
        Response.Write("  <input type=\"text\" name=\"txtsearchphrase\" id=\"txtsearchphrase\" value=\"STILL NEED TO CREATE SEARCH RESULTS PAGE\" />");
        Response.Write("  <input type=\"button\" name=\"searchButton\" id=\"searchButton\" value=\"Find\" class=\"button\" />");

        /*
        if (iShowViewPicks)
        {
            if (iViewPick)
            {
                lcl_selected_viewpick1 = " selected=\"selected\"";
            } else {
                lcl_selected_viewpick0 = " selected=\"selected\"";
            }

            Response.Write("<div id=\"viewpick\">");
            Response.Write("<strong>Order By: </strong>");
            Response.Write("<select name=\"viewpick\" id=\"dropdown_viewpick\">");
            Response.Write("  <option value=\"0\"" + lcl_selected_viewpick0 + ">View by Start Date then Class Name</option>");
            Response.Write("  <option value=\"1\"" + lcl_selected_viewpick1 + ">View by Class Name then Start Date</option>");
            Response.Write("</select>");
            Response.Write("</div>");
        }
        */

        Response.Write("</form>");
        Response.Write("</div>");
    }

    public void displayCategoryInformation(Int32 iOrgID,
                                           Int32 iCategoryID,
                                           Boolean iDisplayImage,
                                           Boolean iDisplayDesc)
    {
        string sSQL           = "";
        string sCategoryTitle = "";
        string sDisplayImage  = "";
        string sImgURL        = "images/class_category_default.gif";
        string sImgALT        = "Classes and Events Category";

        sSQL  = "SELECT categoryid, ";
        sSQL += " categorytitle, ";
        sSQL += " categorydescription, ";
        sSQL += " isroot, ";
        sSQL += " isnull(imgurl,'EMPTY') as imgurl, ";
        sSQL += " categorysubtitle, ";
        sSQL += " orgid, ";
        sSQL += " sequenceid, ";
        sSQL += " reportgrouping, ";
        sSQL += " imgalttag, ";
        sSQL += " parentcategoryid, ";
        sSQL += " isregatta ";
        sSQL += " FROM egov_class_categories ";
        sSQL += " WHERE categoryid = " + iCategoryID;
        sSQL += " AND orgid = " + iOrgID;

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            sCategoryTitle = myReader["categorytitle"].ToString();

            //check for a category pic (file URL) and ALT (tooltip) info.
            if(iDisplayImage)
            {
                if ((myReader["imgurl"].ToString() != null) && (myReader["imgurl"].ToString() != ""))
                {
                    sImgURL = myReader["imgurl"].ToString();
                    sImgALT = myReader["imgalttag"].ToString();
                }

                sDisplayImage  = "<a href=\"class_list.aspx?categoryid=" + iCategoryID.ToString() + "\">";
                sDisplayImage += "<img src=\"" + sImgURL + "\" align=\"left\" class=\"categoryimage\" title=\"" + sImgALT + "\" />";
                sDisplayImage += "</a>";
            }

            Response.Write("<div class=\"categorygroup\" onclick=\"location.href='class_list.aspx?categoryid=" + iCategoryID + "';\">");
            Response.Write("  <div class=\"categorytitle_mobile\">" + myReader["categorytitle"].ToString() + "</div>");
            //Response.Write("  <img src=\"bg_btn_bar.png\" class=\"categorygroup_rightarrow\" />");
            Response.Write("  <ul id=\"categoryOptionList\">");
            Response.Write("  <li>");
            Response.Write("  <fieldset class=\"fieldset_classes\">");
            Response.Write("    <legend><a href=\"class_list.aspx?categoryid=" + iCategoryID + "\" class=\"categorytitle_new\">" + myReader["categorytitle"].ToString() + "</a></legend>");
            Response.Write("    <div class=\"classCategoryOptions\">");
            Response.Write("      LEFT OFF HERE 2<br />");
            Response.Write(       sDisplayImage);

            if (myReader["categorysubtitle"].ToString() != "")
            {
                Response.Write("      <font class=\"categorysubtitle\">" + myReader["categorysubtitle"] + "</font><br /><br />");
            }

            if (iDisplayDesc)
            {
                if (myReader["categorydescription"].ToString() != "")
                {
                    Response.Write("      <font class=\"categorydescription\">" + myReader["categorydescription"].ToString() + "</font><br />");
                }
            }

            Response.Write("    </div>");
            Response.Write("  </fieldset>");
            Response.Write("  </li>");
            Response.Write("  </ul>");
            Response.Write("</div>");
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
    }

    public void setupMobileClassCategoryOptions()
    {
        Response.Write("<div id=\"mobile_subcategorymenu\">");
        Response.Write("  <a id=\"mobile_bookmark\" name=\"mobile_bookmark\"><div id=\"mobile_subcategorymenu_option\">Categories</div></a>");
        Response.Write("  <div id=\"mobile_subcategorylist\"></div>");
        Response.Write("</div>");
    }

    public void listCategories(Int32 iOrgID, 
                               Int32 iCategoryID, 
                               Boolean iDisplayImage,
                               Boolean iDisplayDesc)
    {
        string sSQL = "";

        sSQL  = "SELECT categoryid, ";
        sSQL += " categorytitle, ";
        sSQL += " categorydescription, ";
        sSQL += " isroot, ";
        sSQL += " imgurl, ";
        sSQL += " categorysubtitle, ";
        sSQL += " orgid, ";
        sSQL += " parentcategoryid, ";
        sSQL += " subcategoryid, ";
        sSQL += " childcategoryid, ";
        sSQL += " childcategorytitle, ";
        sSQL += " childcategorydescription, ";
        sSQL += " childimgurl, ";
        sSQL += " childcategorysubtitle, ";
        sSQL += " sequenceid ";
        sSQL += " FROM egov_subcategories ";
        sSQL += " WHERE categoryid = " + iCategoryID;
        sSQL += " ORDER BY sequenceid";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            Response.Write("<div class=\"categoryitem\">");

            while (myReader.Read())
            {
                displayCategoryInformation(iOrgID, 
                                           Convert.ToInt32(myReader["subcategoryid"]), 
                                           iDisplayImage,
                                           iDisplayDesc);

                listCategories(iOrgID, 
                               Convert.ToInt32(myReader["subcategoryid"]), 
                               iDisplayImage,
                               iDisplayDesc);
            }

            Response.Write("</div>");
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

    }

    public void listCategories2(Int32 iOrgID,
                                Int32 iCategoryID,
                                Boolean iDisplayImage,
                                Boolean iDisplayDesc)
    {
        string sSQL              = "";
        string sCategoryTitle    = "";
        string sSubCategoryTitle = "";

        Int32 sLineCount = 0;

        sSQL = "SELECT categoryid, ";
        sSQL += " categorytitle, ";
        sSQL += " categorydescription, ";
        sSQL += " isroot, ";
        sSQL += " imgurl, ";
        sSQL += " categorysubtitle, ";
        sSQL += " orgid, ";
        sSQL += " parentcategoryid, ";
        sSQL += " subcategoryid, ";
        sSQL += " childcategoryid, ";
        sSQL += " childcategorytitle, ";
        sSQL += " childcategorydescription, ";
        sSQL += " childimgurl, ";
        sSQL += " childcategorysubtitle, ";
        sSQL += " sequenceid ";
        sSQL += " FROM egov_subcategories ";
        sSQL += " WHERE categoryid = " + iCategoryID;
        sSQL += " ORDER BY sequenceid";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
                sLineCount        = sLineCount + 1;
                sCategoryTitle    = "";
                sSubCategoryTitle = myReader["childcategorytitle"].ToString();

                if (sLineCount == 1)
                {
                    sCategoryTitle = myReader["categorytitle"].ToString();

                    Response.Write("<ul id=\"categorylist\">");
                    Response.Write("<fieldset class=\"fieldset_classes\">");
                    Response.Write("  <legend>" + sCategoryTitle + "</legend>");
                    Response.Write("  <div class=\"mobile_subcategories\">" + sCategoryTitle + "</div>");
                    //Response.Write("<div style=\"border-bottom:1pt solid #404040;background-color:#eeeeee;font-size:1em;\">" + sCategoryTitle + "</div>");
                }


                Response.Write("<li>");
                Response.Write("  <div class=\"mobile_subcategories\">" + sSubCategoryTitle + "</div>");
                Response.Write("  <fieldset class=\"fieldset_subcategories\">");
                Response.Write("    <legend><a href=\"#\">" + sSubCategoryTitle + "</a></legend>");
                Response.Write("    <div class=\"subcategoryinfo\">");
                Response.Write("      subcategory info here");
                Response.Write("    </div>");
                Response.Write("  </fieldset>");
                Response.Write("</li>");
            }

            Response.Write("</ul>");
            Response.Write("</fieldset>");
        
        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
    }

    public void displayItem(Int32 iSubCategoryID,
                            Boolean iShowViewPicks,
                            Boolean iViewPick)
    {
        string sSQL        = "";
        string sTodaysDate = DateTime.Now.ToString();
        string sDateLine   = "";
        string lcl_selected_viewpick0 = "";
        string lcl_selected_viewpick1 = "";

        Int32 sLineCount = 0;

        Boolean sClassSeasonIsViewable = false;

        sSQL  = "SELECT classid, ";
        sSQL += " classtypeid, ";
        sSQL += " isparent, ";
        sSQL += " classname, ";
        sSQL += " categorytitle, ";
        sSQL += " classdescription, ";
        sSQL += " isnull(startdate,0) AS startdate, ";
        sSQL += " isnull(enddate,0) AS enddate ";
        sSQL += " FROM egov_class_to_categories ";
        sSQL += " WHERE categoryid = " + iSubCategoryID.ToString();
        sSQL += " AND (('" + sTodaysDate + "' BETWEEN publishstartdate AND publishenddate) OR noenddate = 1) ";
        sSQL += " AND displaytopublic = 1 ";
        sSQL += " AND statusname = 'ACTIVE' ";

        if (iViewPick)
        {
            sSQL += " ORDER BY classname, noenddate DESC, startdate, isparent DESC ";
        }
        else
        {
            sSQL += " ORDER BY noenddate DESC, startdate, isparent DESC, classname ";
        }

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            Response.Write("<div class=\"subcategoryitem\">");
            if (iShowViewPicks)
            {
                if (iViewPick)
                {
                    lcl_selected_viewpick1 = " selected=\"selected\"";
                }
                else
                {
                    lcl_selected_viewpick0 = " selected=\"selected\"";
                }

                Response.Write("<div id=\"viewpick\">");
                Response.Write("  <form id=\"reorderList\" action=\"class_list.aspx\" method=\"post\">");
                Response.Write("    <input type=\"hidden\" name=\"categoryid\" id=\"categoryid\" value=\"" + iSubCategoryID.ToString() + "\" />");
                Response.Write("  <strong>Order By: </strong>");
                Response.Write("  <select name=\"viewpick\" id=\"dropdown_viewpick\">");
                Response.Write("    <option value=\"0\"" + lcl_selected_viewpick0 + ">View by Start Date then Class Name</option>");
                Response.Write("    <option value=\"1\"" + lcl_selected_viewpick1 + ">View by Class Name then Start Date</option>");
                Response.Write("  </select>");
                Response.Write("  </form>");
                Response.Write("</div>");
            }
            //Response.Write("<ul id=\"classOptionList\">");

            while (myReader.Read())
            {
                sLineCount = sLineCount + 1;
                sClassSeasonIsViewable = classSeasonIsViewable(Convert.ToInt32(myReader["classid"]));

                if (sClassSeasonIsViewable)
                {
                    sDateLine = buildDateLine(Convert.ToDateTime(myReader["startdate"]),
                                              Convert.ToDateTime(myReader["enddate"]),
                                              Convert.ToInt32(myReader["classid"]),
                                              Convert.ToInt32(myReader["classtypeid"]),
                                              Convert.ToBoolean(myReader["isparent"]));

                    /*
                    Response.Write("<li onclick=\"clickToRegister()\">");
                    Response.Write("<fieldset>");
                    Response.Write("  <legend>" + myReader["classname"].ToString() + "</legend>");
                    Response.Write("  <div class=\"categoryDateLine\">" + sDateLine + "</div>");
                    Response.Write("  <div class=\"mobileClickHere\">REGISTER HERE</div>");
                    Response.Write("  <div class=\"categoryInfo\">category info goes here!</div>");
                    Response.Write("</fieldset>");
                    Response.Write("</li>");
                    */

                    Response.Write("<div class=\"classOptionItem_mobile\" onclick=\"clickToRegister()\">");
                    Response.Write("<table class=\"classOptionItemTable_mobile\">");
                    Response.Write("  <tr>");
                    Response.Write("      <td class=\"mobileinfo\">");
                    Response.Write("          <div class=\"className_mobile\">" + myReader["classname"].ToString() + "</div>");
                    Response.Write("          <div class=\"categoryDateLine_mobile\">" + sDateLine + "</div>");
                    Response.Write("      </td>");
                    Response.Write("      <td class=\"rightarrow\">></td>");
                    Response.Write("  </tr>");
                    Response.Write("</table>");
                    Response.Write("</div>");

                    Response.Write("<div class=\"classOptionItem\">");
                    Response.Write("<fieldset>");
                    Response.Write("  <legend>" + myReader["classname"].ToString() + "</legend>");
                    Response.Write("  <div class=\"categoryDateLine\">" + sDateLine + "</div>");
                    Response.Write("  <div class=\"categoryInfo\">category info goes here!</div>");
                    Response.Write("</fieldset>");
                    Response.Write("</div>");
                }
            }

            //Response.Write("</ul>");
            Response.Write("</div>");
            Response.Write("line count: [" + sLineCount.ToString() + "]");

        }

        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

    }

    public static Boolean classSeasonIsViewable(Int32 iClassID)
    {
        string sSQL        = "";
        Boolean lcl_return = false;

        sSQL  = "SELECT showpublic ";
        sSQL += " FROM egov_class_seasons s, ";
        sSQL +=      " egov_class c ";
        sSQL += " WHERE s.classseasonid = c.classseasonid ";
        sSQL += " AND c.classid = " + iClassID.ToString();

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            if(Convert.ToBoolean(myReader["showpublic"])) {
                lcl_return = true;
            }
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
        string sDates     = "";

        TimeSpan ts = iEndDate - iStartDate;
        Int32 sDateDiff = ts.Days;
        
        if (iStartDate == iEndDate)
        {
            sDates = string.Format("{0:MMMM d}", iStartDate);
        }
        else
        {
            sDates  = string.Format("{0:MMMM d}",iStartDate);
            sDates += " &ndash; ";
            sDates += string.Format("{0:MMMM d}",iEndDate);

            if (sDateDiff > 7)
            {
                sDaySuffix = "s";
            }
        }

        /// NEED TO FINISH REST OF DRAWDATELINE!!!


        //lcl_return = sDates + " [" + sDateDiff.ToString() + "]";
        lcl_return = sDates;

        return lcl_return;
    }

}

using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_classes_class_list_test : System.Web.UI.Page
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
        common.logThePageVisit(startCounter, "class_list.aspx", "public");
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

    public void displayClassList(Int32 iOrgID,
                                 Int32 iCategoryID,
                                 Boolean iShowViewPicks,
                                 Boolean iViewPick)
    {
        Boolean sClassSeasonIsViewable = false;

        Int32 sOrgID      = 0;
        Int32 sCategoryID = 0;
        Int32 sLineCount  = 0;

        string sSQL                   = "";
        string sTodaysDate            = DateTime.Now.ToString();
        string sDateLine              = "";
        string lcl_selected_viewpick0 = "";
        string lcl_selected_viewpick1 = "";
        string lcl_bgcolor            = "#eeeeee";
        string sClassOnClick          = "";
        string sRegistrationLine      = "";
        string sClassID               = "";
        string sClassName             = "";
        //string sCategoryTitle         = "";
        string sKeywordSearchMsg      = "";
        string sClassDescription      = "";
        string sKeywordSearch         = "";
        string sKeywordSearchOriginal = "";
        string sWhereClause           = "";

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
	sWhereClause = "";

        if (!String.IsNullOrEmpty(Request.QueryString["keywordSearch"]))
        {
            sKeywordSearchOriginal = Request.QueryString["keywordSearch"];
            sKeywordSearch         = sKeywordSearchOriginal;
            //sKeywordSearch         = common.dbSafe(sKeywordSearch);
            sKeywordSearch         = sKeywordSearch.Replace("'", "''");
            sKeywordSearch         = "'%" + sKeywordSearch + "%'";

	    if (!String.IsNullOrEmpty(sWhereClause))
	    {
		    sWhereClause += " AND ";
	    }
            sWhereClause += " (searchkeywords LIKE " + sKeywordSearch;
            sWhereClause += " OR classname LIKE " + sKeywordSearch;
            sWhereClause += " OR classdescription LIKE " + sKeywordSearch;
            sWhereClause += " OR activitynos LIKE " + sKeywordSearch;
            sWhereClause += " OR instructors LIKE " + sKeywordSearch + ")";

            //sKeywordSearchMsg = "Showing search results for <span class=\"classesKeywordSearchText\">" + sKeywordSearchOriginal + "</span>";

            //Response.Write("  <fieldset class=\"fieldset_subcategories_classlist\">");
            //Response.Write("  <fieldset>");
            //Response.Write("    <div class=\"subcategoryinfo\">");
            Response.Write("      <div id=\"classesSearchResultsText\">" + sKeywordSearchMsg + "</div>");
            //Response.Write("    </div>");
            //Response.Write("  </fieldset>");
        }


        //if (String.IsNullOrEmpty(sKeywordSearchOriginal) && Request.Querystring["categoryid"])
        //{
            //sCategoryID = iCategoryID;
        //}

        if (!String.IsNullOrEmpty(Request.QueryString["categoryid"]))
        {
	    if (!String.IsNullOrEmpty(sWhereClause))
	    {
		    sWhereClause += " AND ";
	    }
	    sKeywordSearchOriginal = "- ";
	    sWhereClause += " categories like '%##" + iCategoryID + "##%'";
	}
        if (!String.IsNullOrEmpty(Request.QueryString["season"]))
        {
	    if (!String.IsNullOrEmpty(sWhereClause))
	    {
		    sWhereClause += " AND ";
	    }
	    sKeywordSearchOriginal = "- ";
	    sWhereClause += " classseasonid = " + Request.QueryString["season"].ToString();
	}

	bool nameSort = true;
	try
	{
		if (!String.IsNullOrEmpty(Request.QueryString["sort"].ToString()))
		{
			nameSort = false;
		}
	}
	catch
	{
	}

        sSQL  = "SELECT r.classid, ";
        sSQL += " r.classtypeid, ";
        sSQL += " r.isparent, ";
        sSQL += " r.classname, ";
        sSQL += " r.classdescription, ";
        sSQL += " isnull(startdate,0) AS startdate, ";
        sSQL += " isnull(enddate,0) AS enddate ";

        if (sKeywordSearchOriginal != "")
        {
            sSQL += " , '' as categorytitle ";
            sSQL += " FROM egov_class_to_registration r, ";
            sSQL +=      " egov_class_status s ";
            sSQL += " WHERE r.statusid = s.statusid ";
            sSQL += " AND ((" + sWhereClause + ") ";
            sSQL += " AND s.iscancelled = 0 ";
            sSQL += " AND (('" + sTodaysDate + "' BETWEEN publishstartdate AND publishenddate)) OR noenddate = 1) ";
            sSQL += " AND orgid = " + sOrgID.ToString();
            sSQL += " AND showpublic = 1 ";
        }
        else
        {
            sSQL += " , categorytitle ";
            sSQL += " FROM egov_class_to_categories r ";
            sSQL += " WHERE categoryid = " + sCategoryID;
            sSQL += " AND (('" + sTodaysDate + "' BETWEEN publishstartdate AND publishenddate) OR noenddate = 1) ";
            sSQL += " AND statusname = 'ACTIVE' ";
        }

        sSQL += " AND displaytopublic = 1 ";

        if (iViewPick || nameSort)
        {
            sSQL += " ORDER BY ltrim(rtrim(classname)), noenddate DESC, startdate, isparent DESC ";
        }
        else
        {
            sSQL += " ORDER BY noenddate DESC, startdate, isparent DESC, ltrim(rtrim(classname)) ";
        }
	//Response.Write(sSQL);

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            Response.Write("<div class=\"classlist\">");

            if (iShowViewPicks && sCategoryID > 0)
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
                Response.Write("    <input type=\"hidden\" name=\"categoryid\" id=\"categoryid\" value=\"" + sCategoryID + "\" />");
                Response.Write("  <strong>Order By: </strong>");
                Response.Write("  <select name=\"viewpick\" id=\"dropdown_viewpick\">");
                Response.Write("    <option value=\"0\"" + lcl_selected_viewpick0 + ">View by Start Date then Class Name</option>");
                Response.Write("    <option value=\"1\"" + lcl_selected_viewpick1 + ">View by Class Name then Start Date</option>");
                Response.Write("  </select>");
                Response.Write("  </form>");
                Response.Write("</div>");
            }

            while (myReader.Read())
            {
                sLineCount             = sLineCount + 1;
                sClassSeasonIsViewable = classSeasonIsViewable(Convert.ToInt32(myReader["classid"]));
                sClassID               = myReader["classid"].ToString();
                sClassName             = common.decodeUTFString(myReader["classname"].ToString().Trim());
                sClassDescription      = common.decodeUTFString(myReader["classdescription"].ToString().Trim());
                //sCategoryTitle         = myReader["categorytitle"].ToString();

                if (sKeywordSearch != "")
                {
                    sCategoryID = classes.getFirstCategoryID_byClassID(sOrgID,
                                                                       Convert.ToInt32(sClassID));
                }

                if (sClassSeasonIsViewable)
                {
                    lcl_bgcolor = common.changeBGColor(lcl_bgcolor, 
                                                       "#eeeeee",
                                                       "#ffffff");

                    sDateLine = classes.buildDateLine(Convert.ToDateTime(myReader["startdate"]),
                                                      Convert.ToDateTime(myReader["enddate"]),
                                                      Convert.ToInt32(sClassID),
                                                      Convert.ToInt32(myReader["classtypeid"]),
                                                      Convert.ToBoolean(myReader["isparent"]));

                    sClassOnClick  = "viewClassDetails(";
                    sClassOnClick += "'" + sClassID    + "',";
                    sClassOnClick += "'" + sCategoryID + "'";
                    //sClassOnClick += "'" + sCategoryTitle + "'";
                    sClassOnClick += ");";

                    sRegistrationLine = classes.displayRegistrationLine(sClassID);

                    Response.Write("<div style=\"background-color: " + lcl_bgcolor + ";\" class=\"classOptionItem_mobile\" onclick=\"" + sClassOnClick + "\">");
                    Response.Write("  <div class=\"className_mobile\">" + sClassName + "</div>");
                    Response.Write("  <div class=\"categoryDateLine_mobile\">" + sDateLine + "</div>");
                    Response.Write("</div>");

                    Response.Write("<div class=\"classOptionItem\" onclick=\"" + sClassOnClick + "\">");
                    Response.Write("<fieldset class=\"classinfo_classlist\">");
                    Response.Write("  <legend>" + sClassName + "</legend>");
                    Response.Write("  <div class=\"categoryDateLine\">" + sDateLine + "</div>");
                    Response.Write("  <div class=\"categoryInfo\">");
                    //Response.Write( );
                    Response.Write(sRegistrationLine);
                    Response.Write("  </div>");
                    Response.Write("</fieldset>");
                    Response.Write("</div>");
                }
            }

            Response.Write("</div>");
        }
	else
	{
		Response.Write("<div style=\"margin-left:41px;\">No Results were found or no search criteria was entered.  Please enter or modify your search criteria above.</div>");
	}


        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
    }

    public void displayCategoryInfo(Int32 iOrgID,
                                    Int32 iCategoryID)
    {
        Int32 sClassID = 0;

        string sSQL                       = "";
        string sBreadCrumbsLine           = "";
        string sBreadCrumbsReturnLocation = "CLASSLIST";

        string sDisplayImage        = "";
        string sSubCategoryID       = "";
        string sSubCategoryTitle    = "";
        string sSubCategorySubTitle = "";
        string sSubCategoryImgURL   = "images/class_category_default.gif";
        string sSubCategoryImgALT   = "Classes and Events Category";
        string sSubCategoryDesc     = "";

        sSQL = "SELECT childcategoryid, ";
        sSQL += " childcategorytitle, ";
        sSQL += " childcategorysubtitle, ";
        sSQL += " childcategorydescription, ";
        sSQL += " childimgurl, ";
        sSQL += " childimgalttag ";
        sSQL += " FROM egov_subcategories ";
        sSQL += " WHERE subcategoryid = " + iCategoryID;

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            myReader.Read();

            sSubCategoryID       = common.decodeUTFString(myReader["childcategoryid"].ToString().Trim());
            sSubCategoryTitle    = common.decodeUTFString(myReader["childcategorytitle"].ToString().Trim());
            sSubCategorySubTitle = common.decodeUTFString(myReader["childcategorysubtitle"].ToString().Trim());
            sSubCategoryDesc     = common.decodeUTFString(myReader["childcategorydescription"].ToString().Trim());

            //Check for a category pic (file URL) and ALT (tooltip) info.
            if ((myReader["childimgurl"].ToString() != null) && (myReader["childimgurl"].ToString() != ""))
            {
                sSubCategoryImgURL = myReader["childimgurl"].ToString();
                sSubCategoryImgALT = sSubCategoryTitle;
		    sSubCategoryImgURL = sSubCategoryImgURL.Replace("http://www.egovlink.com","");
            }

            if ((myReader["childimgalttag"].ToString() != null) && (myReader["childimgalttag"].ToString() != ""))
            {
                sSubCategoryImgALT = myReader["childimgalttag"].ToString();
            }

            if ((sSubCategoryTitle != null) && (sSubCategoryTitle != ""))
            {
                sSubCategoryTitle = "<div class=\"subcategorysubtitle\">" + sSubCategoryTitle + "</div>";
            }

            sDisplayImage = "<img src=\"" + sSubCategoryImgURL + "\" align=\"left\" class=\"categoryimage\" title=\"" + sSubCategoryImgALT + "\" onclick=\"viewCategoryClassList('" + iCategoryID + "');\" />";

            sBreadCrumbsLine = classes.buildBreadCrumbsCategories(iOrgID,
                                                                  sBreadCrumbsReturnLocation,
                                                                  iCategoryID,
                                                                  sClassID);

            Response.Write(sBreadCrumbsLine);

            Response.Write("  <div class=\"mobile_subcategories_classlist\" onclick=\"viewCategoryClassList('" + iCategoryID + "');\">" + sSubCategoryTitle + "</div>");
            Response.Write("  <fieldset class=\"fieldset_subcategories_classlist\" title=\"Click to view all classes/events\" onclick=\"viewCategoryClassList('" + iCategoryID + "');\">");
            Response.Write("    <div class=\"subcategoryinfo\">");
            Response.Write(sDisplayImage);
            Response.Write(sSubCategoryTitle);
            Response.Write(sSubCategoryDesc);
            Response.Write("    </div>");
            Response.Write("  </fieldset>");
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

    public void displaySeasons(int intOrgID, string selSeason)
    {
	string sSql = "SELECT C.classseasonid, C.seasonname  ";
		sSql += "FROM egov_class_seasons C ";
		sSql += "INNER JOIN egov_seasons S  ON C.seasonid = S.seasonid  ";
		sSql += "INNER JOIN egov_class_to_registration r ON r.classseasonid = c.classseasonid ";
		sSql += "WHERE C.isclosed = 0 AND c.orgid = " + intOrgID;
		sSql += "AND (((GETDATE() BETWEEN publishstartdate AND publishenddate)) OR noenddate = 1) ";
 		sSql += "AND statusid = 1 AND displaytopublic = 1  and r.showpublic = 1  ";
		sSql += "GROUP BY C.classseasonid, C.seasonname, C.seasonyear, S.displayorder ";
		sSql += "ORDER BY C.seasonyear DESC, S.displayorder DESC, C.seasonname ";
        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSql, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
		    string selected = "";
		    if (myReader["classseasonid"].ToString() == selSeason)
		    {
			    selected = " selected";
		    }
		    Response.Write("<option value=\"" + myReader["classseasonid"].ToString() + "\" " + selected + ">" + myReader["seasonname"].ToString() + "</option>");
	    }
	}
        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
    }

    public void displayCategories(int intOrgID, int selCategory)
    {
	string sSql = "SELECT categoryid, categorytitle FROM egov_class_categories WHERE orgid = " + intOrgID;
	sSql += " AND isroot = 0 AND isregatta = 0 ORDER BY categorytitle";
        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSql, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
		    string selected = "";
		    if (myReader["categoryid"].ToString() == selCategory.ToString())
		    {
			    selected = " selected";
		    }
		    Response.Write("<option value=\"" + myReader["categoryid"].ToString() + "\" " + selected + ">" + myReader["categorytitle"].ToString() + "</option>");
	    }
	}
        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();
    }

}

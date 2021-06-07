using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_classes_class_categories : System.Web.UI.Page
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
        //THIS STILL NEEDS FIXED!!!!
        common.logThePageVisit(startCounter, "class_categories.aspx", "public");
    }

    //public bool HasFile { get; }

    protected void Page_Load(object sender, EventArgs e)
    {
        string sPublicURL = "";

        sPublicURL = common.getFeaturePublicURL(sOrgID, "activities");

        if (sPublicURL != "rd_classes/class_categories.aspx")
        {
            Response.Redirect("../" + sPublicURL);
        }

        /*
        if (IsPostBack)
        {
            Boolean fileOK = false;
            //String path = Server.MapPath("~/UploadedImages/");
            String path = Server.MapPath("~/rd_classes/");
            if (FileUpload1.HasFile)
            {
                String fileExtension =
                    System.IO.Path.GetExtension(FileUpload1.FileName).ToLower();
                String[] allowedExtensions = { ".gif", ".png", ".jpeg", ".jpg" };
                for (int i = 0; i < allowedExtensions.Length; i++)
                {
                    if (fileExtension == allowedExtensions[i])
                    {
                        fileOK = true;
                    }
                }
            }

            if (fileOK)
            {
                try
                {
                    FileUpload1.PostedFile.SaveAs(path
                        + FileUpload1.FileName);
                    Label1.Text = "File uploaded!";
                }
                catch (Exception ex)
                {
                    Label1.Text = "File could not be uploaded.";
                }
            }
            else
            {
                Label1.Text = "Cannot accept files of this type.";
            }
        }
        */
    }

    public void listCategories(Int32 iOrgID,
                               Int32 iCategoryID)
    {
        string sSQL                 = "";
        string sDisplayImage        = "";
        string sCategoryTitle       = "";
        string sSubCategoryID       = "";
        string sSubCategoryTitle    = "";
        string sSubCategorySubTitle = "";
        string sSubCategoryDesc     = "";
        string sSubCategoryImgURL   = "images/class_category_default.gif";
        string sSubCategoryImgALT   = "Classes and Events Category";

        string sSessionID = HttpContext.Current.Session.SessionID;

        Int32 sLineCount = 0;

        sSQL = "SELECT categoryid, ";
        sSQL += " categorytitle, ";
        sSQL += " categorydescription, ";
        sSQL += " isroot, ";
        sSQL += " imgurl, ";
        sSQL += " imgalttag, ";
        sSQL += " categorysubtitle, ";
        sSQL += " orgid, ";
        sSQL += " parentcategoryid, ";
        sSQL += " subcategoryid, ";
        sSQL += " childcategoryid, ";
        sSQL += " childcategorytitle, ";
        sSQL += " childcategorysubtitle, ";
        sSQL += " childcategorydescription, ";
        sSQL += " childimgurl, ";
        sSQL += " childimgalttag, ";
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
                sLineCount           = sLineCount + 1;
                sCategoryTitle       = "";
                sSubCategoryID       = myReader["childcategoryid"].ToString();
                sSubCategoryTitle    = myReader["childcategorytitle"].ToString();
                sSubCategorySubTitle = myReader["childcategorysubtitle"].ToString();
                sSubCategoryDesc     = myReader["childcategorydescription"].ToString();

                if (sLineCount == 1)
                {
                    sCategoryTitle = myReader["categorytitle"].ToString();

                    Response.Write("<ul id=\"categorylist\">");
                    Response.Write("<div class=\"mobile_subcategories\" onclick=\"viewCategoryList('" + iCategoryID.ToString() + "');\">" + sCategoryTitle + "</div>");
                    Response.Write("<fieldset class=\"fieldset_classes\">");
                    Response.Write("  <legend onclick=\"viewCategoryList('" + iCategoryID.ToString() + "');\">" + sCategoryTitle + "</legend>");
                }

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

                if ((sSubCategorySubTitle != null) && (sSubCategorySubTitle != ""))
                {
                    sSubCategorySubTitle = "<div class=\"subcategorysubtitle\">" + sSubCategorySubTitle + "</div>";
                }

                sDisplayImage = "<img src=\"" + sSubCategoryImgURL + "\" align=\"left\" class=\"categoryimage\" title=\"" + sSubCategoryImgALT + "\" onclick=\"viewCategoryClassList('" + sSubCategoryID.ToString() + "');\" />";

                Response.Write("<li>");
                Response.Write("  <div class=\"mobile_subcategories\" onclick=\"viewCategoryClassList('" + sSubCategoryID + "');\">" + sSubCategoryTitle + "</div>");
                Response.Write("  <fieldset class=\"fieldset_subcategories\" title=\"Click to view all classes/events\" onclick=\"viewCategoryClassList('" + sSubCategoryID + "');\">");
                Response.Write("    <legend>" + sSubCategoryTitle + "</legend>");
                Response.Write("    <div class=\"subcategoryinfo\">");
                Response.Write(sDisplayImage);
                Response.Write(sSubCategorySubTitle);
                Response.Write(sSubCategoryDesc);
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
}

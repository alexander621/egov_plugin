using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class rd_includes_egovcontent : System.Web.UI.UserControl
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    public void showMemberWarning()
    {
        string sOrgID      = common.getOrgId();
        string sOrgDisplay = "";

        if(common.orgHasDisplay(sOrgID,"classdetailsnotice"))
        {
            sOrgDisplay = common.getOrgDisplay(sOrgID, "classdetailsnotice");

            Response.Write("<div id=\"classdetailsnotice\">" + sOrgDisplay + "</div>");
        }
    }

    public static Int32 getFirstCategory()
    {
        Int32 lcl_return = 0;
        string sOrgID    = common.getOrgId();
        string sSQL      = "";

        sSQL  = "SELECT categoryid ";
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

    public void displaySubCategoryMenu(Int32 iRootCategoryID, Boolean iShowViewPicks, Boolean iViewPick, Int32 iCategoryID )
    {
        string sOrgID = common.getOrgId();
        string sSQL   = "";

        sSQL  = "SELECT categorytitle, ";
        sSQL += " subcategoryid, ";
        sSQL += " subcategorytitle ";
        sSQL += " FROM class_categories ";
        sSQL += " WHERE orgid = " + sOrgID;
        sSQL += " AND categoryid = " + iRootCategoryID;
        sSQL += " ORDER BY sequenceid, subcategorytitle";

//        if(iCategoryID == iRootCategoryID) {

//          displaySubCategoryMenu(Int32 iRootCategoryID, Boolean iShowViewPicks, Boolean iViewPick, Int32 iCategoryID )
          
//          Response.Write("here1");
//      } else {
//          Response.Write("here2");
//      }




    }
/*
    Sub DisplaySubCategoryMenu( ByVal orgid, ByVal iRootCategoryID, ByVal bShowViewPicks, ByVal iViewPick, ByVal iCategoryID )
	Dim sSql, oRs

        ' GET SUBCATEGORIES FOR THIS CATEGORY
        sSql = "SELECT categorytitle, subcategoryid, subcategorytitle FROM class_categories "
		sSql = sSql & "WHERE orgid = " & orgid & " AND categoryid = " & iRootCategoryID
		sSql = sSql & " ORDER BY sequenceid, subcategorytitle"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1
        blnFirst = True

        ' DISPLAY LIST OF LINK TO SUBCATEGORIES OF THIS CATEGORY
        If Not oRs.EOF Then
			Response.Write vbcrlf & "<p>" & vbcrlf & "<div class=""subcategorymenu"">" 
			' DISPLAY ROOT CATEGORY
'			response.write("<font class=""subcategorymenuheader"">Browse " & oRs("categorytitle") & ":<br></font>")
			response.write "<a class=""subcategorymenu"" href=""class_list.asp?categoryid=" & iRootCategoryID & """>" & oRs("categorytitle") &  "</a><br />"

			Do While NOT oRs.EOF
				' WRITE SPACER
				If Not blnFirst Then
					Response.Write(" | ")
				End If
				blnFirst = False

				' DISPLAY SUBCATEGORY LINKS
				Response.Write vbcrlf & "<a class=""subcategorymenu"" href=""class_list.asp?categoryid=" & oRs("subcategoryid") & """ >" & oRs("subcategorytitle") & "</a> "
			
				oRs.MoveNext
			Loop

			Response.Write vbcrlf & "</div>"
			
			' DISPLAY SEARCH BOX
			DisplayClassesSearchBox bShowViewPicks, iViewPick, iCategoryID

			response.write vbctlf & "</p>"
		End If

       ' CLEAN UP OBJECTS
		oRs.Close 
		Set oRs = Nothing		

End Sub
*/

}

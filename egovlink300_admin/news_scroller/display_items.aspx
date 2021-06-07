<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>

<script language="c#" runat="server">

//------------------------------------------------------------------------------
//
//------------------------------------------------------------------------------
// FILENAME: display_items.aspx
// AUTHOR:   David Boyer
// CREATED:  07/09/2009
// COPYRIGHT: Copyright 2006 eclink, inc.
//			 All Rights Reserved.

// Description:  This module toggles the display of the News Items.  Copied from display_items.asp

// MODIFICATION HISTORY
// 1.0 10/31/06	Steve Loar - Initial version (ASP).
// 1.1 07/09/09 David Boyer - Added "newstype" to split News and News Scroller items.
// 1.2 07/09/09 David Boyer - Converted to ASP.NET (aspx)

//------------------------------------------------------------------------------
//
//------------------------------------------------------------------------------
    
    
    protected void Page_Load(object sender, EventArgs e)
    {
        string sSQL;
        string iNewsItemID;
        string iNewsType;
        string iItemDisplay;

        int iNewDisplay;
        int lcl_newsitemid;
        int lcl_itemdisplay;

        iNewsItemID     = Request["newsitemid"];
        iNewsType       = Request["newstype"];
        iItemDisplay    = Request["itemdisplay"];

        lcl_newsitemid  = Int32.Parse(iNewsItemID);
        lcl_itemdisplay = Int32.Parse(iItemDisplay);

        if(lcl_itemdisplay == 1) {
           iNewDisplay = 0;
        }else{
           iNewDisplay = 1;
        }

        sSQL = "UPDATE egov_news_items SET itemdisplay = " + iNewDisplay.ToString() + " WHERE newsitemid = " + lcl_newsitemid.ToString();

        SqlConnection sqlConn   = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        SqlCommand sqlCommander = new SqlCommand();
        sqlCommander.Connection = sqlConn;
        sqlConn.Open();
        sqlCommander.CommandText = sSQL;
        //sqlCommander.CommandText = "DELETE FROM egov_merchandise WHERE merchandiseid = " + iMerchandiseid.ToString();
        sqlCommander.ExecuteNonQuery();
        
        sqlConn.Close();

        sqlCommander.Dispose();

        Response.Redirect("list_items.asp?newstype=" + iNewsType + "&success=ASPNET");
    }
    
   
</script>

<%
    //Response.Write(sMerchandiseid);
 %>

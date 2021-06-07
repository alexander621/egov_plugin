<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>

<script language="c#" runat="server">
    
    
    protected void Page_Load(object sender, EventArgs e)
    {
        string sMerchandiseid, sSql;
        int iMerchandiseid;

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);

        sMerchandiseid = Request["merchandiseid"];
        iMerchandiseid = Int32.Parse(sMerchandiseid);
        sSql = "DELETE FROM egov_merchandisecatalog WHERE merchandiseid = " + iMerchandiseid.ToString();
        SqlCommand sqlCommander = new SqlCommand();
        sqlCommander.Connection = sqlConn;

        sqlConn.Open();
        
        sqlCommander.CommandText = "DELETE FROM egov_merchandisecatalog WHERE merchandiseid = " + iMerchandiseid.ToString();
        sqlCommander.ExecuteNonQuery();
        sqlCommander.CommandText = "DELETE FROM egov_merchandise WHERE merchandiseid = " + iMerchandiseid.ToString();
        sqlCommander.ExecuteNonQuery();
        
        sqlConn.Close();

        sqlCommander.Dispose();

        Response.Redirect("merchandiselist.asp");
    }
    
   
</script>

<%
    //Response.Write(sMerchandiseid);
 %>

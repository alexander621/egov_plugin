using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;

public partial class getcoordinates : System.Web.UI.Page
{


    protected void Page_Load(object sender, EventArgs e)
    {
	if (String.IsNullOrEmpty(Request["orgid"]))
	{
		Response.Write("Add '?orgid=VALUE' and try again.");
		Response.End();
	}
	//Lookup addresses w/o coordinates
	string sSQL = "SELECT residentaddressid,residentstreetnumber + ' ' + residentstreetname + ' ' + residentcity + ', ' + residentstate + ' ' + residentzip as address ";
		sSQL += "FROM egov_residentaddresses WHERE orgid = " + Request["orgid"] + "  AND  latitude IS NULL and longitude IS NULL";

        SqlConnection sqlConn = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString);
        sqlConn.Open();

        SqlCommand myCommand = new SqlCommand(sSQL, sqlConn);
        SqlDataReader myReader;
        myReader = myCommand.ExecuteReader();

	int rows = 0;
        if (myReader.HasRows)
        {
            while (myReader.Read())
            {
		rows++;
	        string addy = myReader["address"].ToString();
		int recID = Int32.Parse(myReader["residentaddressid"].ToString());

		if (!String.IsNullOrEmpty(addy))
		{
	        	setCoordinates(addy, recID);
			//Response.Write(addy);
			//Response.End();
		}
	    }
	    Response.Write("Keep reloading until the page reads 'ALL POPULATED'. Count:" + rows.ToString());
	}
	else
	{
		Response.Write("ALL POPULATED");
	}
        myReader.Close();
        sqlConn.Close();
        myReader.Dispose();
        sqlConn.Dispose();

	    
    }

    public void setCoordinates(string recAddy, int recID)
    {
	string URL = "https://maps.googleapis.com/maps/api/geocode/xml?key=AIzaSyD8E3kqTlFodyJsmfrn_ZMARZu4_552Kw0&sensor=false&address=" + recAddy;
	//Response.Write(URL + "<br /><br />");
	//Response.Flush();
	
        WebRequest request = WebRequest.Create(URL);
        request.Method = "GET";

        WebResponse response = null;
        try
        {
            response = request.GetResponse();
        }
        catch
        {
        }

        Stream dataStream = response.GetResponseStream();

        // create a stream reader.
        StreamReader reader = new StreamReader(dataStream);

        // read the content into a string
        string serverResponse = reader.ReadToEnd();

        //Response.Write(serverResponse);

        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.LoadXml(serverResponse);

        string xmlRoot = "/GeocodeResponse/result";

	string lat = "";
	string lng = "";
        try
        {
            lat = xmlDoc.DocumentElement.SelectSingleNode(xmlRoot + "/geometry/location/lat").InnerText;
        }
        catch
        {
            //Nothing
        }
        try
        {
            lng = xmlDoc.DocumentElement.SelectSingleNode(xmlRoot + "/geometry/location/lng").InnerText;
        }
        catch
        {
            //Nothing
        }

	string sSQL = "UPDATE egov_residentaddresses SET latitude = '" + lat + "',longitude = '" + lng + "' WHERE residentaddressid = '" + recID + "'";
	//Response.Write(sSQL + "<br />");

	if (!String.IsNullOrEmpty(lat) && !String.IsNullOrEmpty(lat))
	{
        	string lcl_return = common.RunInsertStatement(sSQL);
		//Response.Write("WOULD HAVE EXECUTED<br />");
	}
	else
	{
		Response.Write(serverResponse + "<br /><br />");
		Response.Flush();
	}
    }

}

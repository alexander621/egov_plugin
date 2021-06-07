using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using System.Text.RegularExpressions;

/// <summary>
/// These are the common methods for charts on the admin side. You do not need to instantiate this class.
/// Try to keep these in alphabetical order, please.
/// </summary>

public class chartcommon
{

	public static void GetBarChartValues( string _ChartId, out string _ChartQuery, out string _ChartTitle, out string _ChartType, out Int32 _ChartHeight, out Int32 _ChartWidth, out Boolean _ShowLegend, out string _LegendTitle, out string _Columnwidth )
	{
		string iOrgId = common.getOrgId( );
		_ChartQuery = "SELECT * FROM egov_charts WHERE chartid = 0";
		_ChartTitle = "";
		_ChartType = "bar";
		_ChartHeight = 600;
		_ChartWidth = 800;
		_Columnwidth = "0.6";
		_LegendTitle = "Legend";
		_ShowLegend = true;

		string sSql = "SELECT chartquery, ISNULL(charttitle,'') AS charttitle, ISNULL(charttype,'bar') AS charttype, ";
		sSql += "showlegend, ISNULL(columnwidth,'0.6') AS columnwidth, ISNULL(chartheight,600) AS chartheight, ";
		sSql += "ISNULL(chartwidth, 800) AS chartwidth, ISNULL(legendtitle,'Legend') AS legendtitle ";
		sSql += "FROM egov_charts WHERE chartid = " + _ChartId + " AND orgid = " + iOrgId;

		SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );

		sqlConn.Open( );

		SqlCommand myCommand = new SqlCommand( sSql, sqlConn );
		SqlDataReader myReader;
		myReader = myCommand.ExecuteReader( CommandBehavior.CloseConnection );

		while ( myReader.Read( ) )
		{
			// There will be just one row for this
			_ChartQuery = myReader["chartquery"].ToString( );
			_ChartTitle = myReader["charttitle"].ToString( );
			_ChartType = myReader["charttype"].ToString( );
			_LegendTitle = myReader["legendtitle"].ToString( );
			_ChartHeight = Int32.Parse( myReader["chartheight"].ToString( ) );
			_ChartWidth = Int32.Parse( myReader["chartwidth"].ToString( ) );
			_Columnwidth = myReader["columnwidth"].ToString( );
			if ( Boolean.Parse(myReader["showlegend"].ToString()) )
				_ShowLegend = true;
			else
				_ShowLegend = false;
		}

		myReader.Close( );
		sqlConn.Close( );
	}


	public static void GetPieChartValues( string _ChartId, out string _ChartQuery, out string _ChartTitle, out Int32 _ChartHeight, out Int32 _ChartWidth, out string _CollectedThreshold, out string _CollectedLabel, out string _CollectedLegendText )
	{
		string iOrgId = common.getOrgId( );
		_ChartQuery = "SELECT * FROM egov_charts WHERE chartid = 0";
		_ChartTitle = "";
		_CollectedThreshold = "0";
		_CollectedLabel = "";
		_CollectedLegendText = "";
		_ChartHeight = 600;
		_ChartWidth = 800;

		string sSql = "SELECT chartquery, ISNULL(charttitle,'') AS charttitle, ISNULL(chartheight,600) AS chartheight, ISNULL(chartwidth, 800) AS chartwidth, ";
		sSql += "ISNULL(collectedthreshold,'0') AS collectedthreshold, ISNULL(collectedlabel,'') AS collectedlabel, ISNULL(collectedlegendtext,'') AS collectedlegendtext ";
		sSql += "FROM egov_charts WHERE chartid = " + _ChartId + " AND orgid = " + iOrgId;

		SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );

		sqlConn.Open( );

		SqlCommand myCommand = new SqlCommand( sSql, sqlConn );
		SqlDataReader myReader;
		myReader = myCommand.ExecuteReader( CommandBehavior.CloseConnection );

		while ( myReader.Read( ) )
		{
			// There will be just one row for this
			_ChartQuery = myReader["chartquery"].ToString();
			_ChartTitle = myReader["charttitle"].ToString();
			_ChartHeight = Int32.Parse( myReader["chartheight"].ToString( ) );
			_ChartWidth = Int32.Parse( myReader["chartwidth"].ToString( ) );
			_CollectedThreshold = myReader["collectedthreshold"].ToString();
			_CollectedLabel = myReader["collectedlabel"].ToString();
			_CollectedLegendText = myReader["collectedlegendtext"].ToString();
		}

		myReader.Close( );
		sqlConn.Close( );
	}

}

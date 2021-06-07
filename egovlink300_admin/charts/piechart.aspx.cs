using System;
using System.Collections;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.Web.UI.DataVisualization.Charting;

namespace System.Web.UI.DataVisualization.Charting.Samples
{
	/// <summary>
	/// Summary description for Pie3D.
	/// </summary>
	public partial class Pie3D : System.Web.UI.Page
	{
		# region Fields

		protected System.Web.UI.WebControls.Label Label1;
		protected System.Web.UI.WebControls.Label Label2;
		protected System.Web.UI.WebControls.Label Label3;
		protected System.Web.UI.WebControls.Label Label4;

		#endregion
	
		protected void Page_Load(object sender, System.EventArgs e)
		{
			string ChartId = Int32.Parse( Request["cid"] ).ToString();
			string ChartQuery;
			string ChartTitle;
			string LegendTitle;
			string CollectedThreshold;
			string CollectedLabel;
			string CollectedLegendText;
			Int32 CollectedTotal = 0;
			Int32 CollectedThresholdAmount = 0;
			string xValue;
			Int32 ChartHeight;
			Int32 ChartWidth;

			chartcommon.GetPieChartValues( ChartId, out ChartQuery, out ChartTitle, out ChartHeight, out ChartWidth, out CollectedThreshold, out CollectedLabel, out CollectedLegendText );
			//Response.Write(ChartQuery);
			//Response.End

			Chart1.Series["Default"].ChartType = SeriesChartType.Pie;
			Chart1.Series["Default"]["PieDrawingStyle"] = "Default";
			Chart1.Series["Default"]["PieLabelStyle"] = "Outside";
			Chart1.Titles[0].Text = ChartTitle;
			//Chart1.Legends[0].Title = LegendTitle;
			Chart1.ImageType = ChartImageType.Jpeg;
			Chart1.Height = ChartHeight;
			Chart1.Width = ChartWidth;
			
			if ( Int32.Parse( CollectedThreshold) > Int32.Parse("0") )
				CollectedThresholdAmount = Int32.Parse(CollectedThreshold);

			ArrayList aXvalues = new ArrayList();
			ArrayList aYvalues = new ArrayList();

			//string sSql = "SELECT ISNULL([Form Name],'empty') AS xvalue, COUNT([Form Name]) AS yvalue FROM egov_rpt_actionline WHERE [Date Submitted] >= '8/1/2007' AND [Date Submitted] <= '8/10/2010' AND orgid = 5 AND UPPER(status) <> 'DISMISSED' AND UPPER(status) <> 'RESOLVED' GROUP BY [Form Name] ORDER BY COUNT([Form Name]) DESC, [FORM NAME]";
			SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
			sqlConn.Open( );

			SqlCommand myCommand = new SqlCommand( ChartQuery, sqlConn );
			SqlDataReader myReader;
			myReader = myCommand.ExecuteReader( CommandBehavior.CloseConnection );

			while ( myReader.Read( ) )
			{
				xValue = myReader["xvalue"].ToString();
				xValue = xValue.Replace("\n","");
				xValue = xValue.Replace( "\r", "" );
				
				aXvalues.Add( xValue );
				aYvalues.Add( myReader["yvalue"] );
				if (Int32.Parse(myReader["yvalue"].ToString()) <= CollectedThresholdAmount )
					CollectedTotal += Int32.Parse(myReader["yvalue"].ToString());
			}

			CollectedLabel += "(" + CollectedTotal.ToString( ) + ")";

			myReader.Close( );
			sqlConn.Close( );

			// Populate series data
			for ( int x = 0; x < aXvalues.Count; x++ )
			{
				Chart1.Series["Default"].Points.AddXY( aXvalues[x] + "(" + aYvalues[x] + ")", aYvalues[x] );
			}

			Chart1.Series["Default"]["CollectedThreshold"] = CollectedThreshold;
			Chart1.Series["Default"]["CollectedLabel"] = CollectedLabel;
			Chart1.Series["Default"]["CollectedLegendText"] = CollectedLegendText;
			
		}

		protected void PostPaint_EmptyChart1( object sender, ChartPaintEventArgs e )
		{
			if ( Chart1.Series["Default"].Points.Count == 0 )
			{
				// create text to draw
				String TextToDraw;
				TextToDraw = "\n\n\nNo data was found. Please change the options and try again.";
				// get graphics tools
				System.Drawing.Graphics g = e.ChartGraphics.Graphics;
				System.Drawing.Font DrawFont = System.Drawing.SystemFonts.CaptionFont;
				System.Drawing.Brush DrawBrush = System.Drawing.Brushes.Red;
				// see how big the text will be
				int TxtWidth = ( int )g.MeasureString( TextToDraw, DrawFont ).Width;
				int TxtHeight = ( int )g.MeasureString( TextToDraw, DrawFont ).Height;
				// where to draw
				int x = 20;  // a few pixels from the left border
				int y = 20; // ( int )e.Chart.Height.Value;
				//y = y - TxtHeight - 155; // a few pixels off the bottom
				// draw the string        
				g.DrawString( TextToDraw, DrawFont, DrawBrush, x, y );
			}
		}

		#region Web Form Designer generated code
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: This call is required by the ASP.NET Web Form Designer.
			//
			InitializeComponent();
			base.OnInit(e);
		}
		
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{    

		}
		#endregion

	}
}

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
	/// Bar chart
	/// </summary>
	public partial class BarColumnChart : System.Web.UI.Page
	{
		protected System.Web.UI.WebControls.Label Label1;
		protected System.Web.UI.WebControls.Label Label2;
	
		protected void Page_Load(object sender, System.EventArgs e)
		{
			string ChartId = Int32.Parse( Request["cid"] ).ToString( );
			string ChartQuery;
			string ChartTitle;
			string ChartType;
			string SeriesName = "none";
			Int32 SeriesCount = -1;
			string xValue;
			string LegendTitle;
			Int32 ChartHeight;
			Int32 ChartWidth;
			string ColumnWidth;
			Boolean ShowLegend;

			chartcommon.GetBarChartValues( ChartId, out ChartQuery, out ChartTitle, out ChartType, out ChartHeight, out ChartWidth, out ShowLegend, out LegendTitle, out ColumnWidth );

			Chart1.Titles[0].Text = ChartTitle;
			Chart1.ImageType = ChartImageType.Jpeg;
			Chart1.Height = ChartHeight;
			Chart1.Width = ChartWidth;

			if ( ShowLegend )
			{
				Chart1.Legends[0].Enabled = true;
				Chart1.Legends[0].Title = LegendTitle;
			}
			else
			{
				Chart1.Legends[0].Enabled = false;
				Chart1.Legends[0].Title = "";
			}

			SqlConnection sqlConn = new SqlConnection( ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString );
			sqlConn.Open( );

			SqlCommand myCommand = new SqlCommand( ChartQuery, sqlConn );
			SqlDataReader myReader;
			myReader = myCommand.ExecuteReader( CommandBehavior.CloseConnection );

			while ( myReader.Read( ) )
			{
				if ( myReader["seriesname"].ToString() != SeriesName )
				{
					SeriesName = myReader["seriesname"].ToString();
					SeriesCount++;
					Chart1.Series.Add(SeriesName);

					if ( ChartType == "bar" )
					{
						Chart1.Series[SeriesCount].ChartType = SeriesChartType.Bar;
					}
					else
					{
						Chart1.Series[SeriesCount].ChartType = SeriesChartType.Column;
					}

					Chart1.Series[SeriesCount]["PointWidth"] = ColumnWidth;
					Chart1.Series[SeriesCount]["DrawingStyle"] = "Default";
				}
				xValue = myReader["xvalue"].ToString( );
				xValue = xValue.Replace( "\n", "" );
				xValue = xValue.Replace( "\r", "" );

				Chart1.Series[SeriesCount].Points.AddXY( xValue, myReader["yvalue"] );
			}

            if (SeriesCount < 0)
		    {
                Chart1.Series.Add(SeriesName);
            }

			myReader.Close( );
			sqlConn.Close( );

            Chart1.AlignDataPointsByAxisLabel( );
            Chart1.ChartAreas[0].AxisX.Interval = 1; // THis should force all labels to display

			if ( ChartType == "column" )
				Chart1.ChartAreas[0].AxisX.LabelStyle.Angle = 40;
			else
				Chart1.ChartAreas[0].AxisX.LabelStyle.Angle = 0;
		}

		protected void PostPaint_EmptyChart1( object sender, ChartPaintEventArgs e )
		{
			if ( Chart1.Series[0].Points.Count == 0 )
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

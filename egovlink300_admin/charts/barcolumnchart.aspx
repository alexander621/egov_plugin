<%@ Page language="c#" Inherits="System.Web.UI.DataVisualization.Charting.Samples.BarColumnChart" CodeFile="BarColumnChart.aspx.cs" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<html>
	<head>
		<title>E-Gov Administration Console</title>
		<meta content="Microsoft Visual Studio 7.0" name="GENERATOR" />
		<meta content="C#" name="CODE_LANGUAGE" />
		<meta content="JavaScript" name="vs_defaultClientScript" />
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
		<link media="all" type="text/css" rel="stylesheet" href="charts.css" />
	</head>
	<body>
		<form id="Form1" method="post" runat="server">
			<p align="center">
			<table class="sampleTable" cellpadding="0" cellspacing="0" border="0">
				<tr>
					<td class="tdchart">

<!--		removed: OnPostPaint="PostPaint_EmptyChart1" because if there were zero records in the series it would crash the chart -->
						<asp:CHART id="Chart1" runat="server" OnPostPaint="PostPaint_EmptyChart1" Palette="BrightPastel" Width="1000px" 
							Height="600px" BorderDashStyle="Solid" 
							BorderWidth="2" BorderColor="181, 64, 1" ImageType="Jpeg" BackSecondaryColor="White">
							<titles>
								<asp:Title ShadowColor="32, 0, 0, 0" Font="Trebuchet MS, 14.25pt, style=Bold" ShadowOffset="3" Text="Column Chart" Name="Title1" ForeColor="26, 59, 105"></asp:Title>
							</titles>
							<legends>
								<asp:Legend TitleFont="Arial Unicode MS, 8.25pt, style=Bold" 
									BackColor="Transparent" Font="Arial Unicode MS, 8pt" 
									IsTextAutoFit="False" Name="Default" LegendStyle="Column" Title="Legend"></asp:Legend>
							</legends>
							<chartareas>
								<asp:ChartArea Name="ChartArea1" BorderColor="64, 64, 64, 64" 
									BackSecondaryColor="White" BackColor="White" ShadowColor="Transparent" 
									BackGradientStyle="TopBottom">
									<area3dstyle Rotation="10" Perspective="10" Inclination="15" IsRightAngleAxes="False" WallWidth="0" IsClustered="False" />
									<axisy LineColor="64, 64, 64, 64"  LabelAutoFitMaxFontSize="8" 
										islabelautofit="False">
										<LabelStyle Font="Arial Unicode MS, 9pt" />
										<MajorGrid LineColor="64, 64, 64, 64" />
									</axisy>
									<axisx LineColor="64, 64, 64, 64"  LabelAutoFitMaxFontSize="8" 
										IsLabelAutoFit="False" interval="1">
										<LabelStyle Font="Arial, 9pt" IsEndLabelVisible="False" 
											Angle="45" />
										<MajorGrid LineColor="64, 64, 64, 64" Interval="Auto" IntervalOffset="Auto" 
											IntervalOffsetType="Auto" IntervalType="Auto" />
										<MajorTickMark Interval="Auto" />
									</axisx>
								</asp:ChartArea>
							</chartareas>
						</asp:CHART>
					</td>
				</tr>
			</table>
			</p>
		</form>
	</body>
</html>

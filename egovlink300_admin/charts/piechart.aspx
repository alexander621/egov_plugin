
<%@ Page language="c#" Inherits="System.Web.UI.DataVisualization.Charting.Samples.Pie3D" CodeFile="piechart.aspx.cs" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<html>
	<head>
		<title>E-Gov Administration Console</title>
		<meta content="Microsoft Visual Studio 7.0" name="GENERATOR"/>
		<meta content="C#" name="CODE_LANGUAGE"/>
		<meta content="JavaScript" name="vs_defaultClientScript"/>
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
		<link media="all" type="text/css" rel="stylesheet" href="charts.css" />
	</head>
	<body>
		<form id="PieChart" method="post" runat="server">
			<p align="center" class="chartcontainer">
				<table cellpadding="0" cellspacing="0" border="0" class="sampleTable">
					<tr>
						<td>
							<asp:CHART id="Chart1" runat="server" OnPostPaint="PostPaint_EmptyChart1" Palette="BrightPastel" Height="600px" 
								Width="1000px" BorderDashStyle="Solid" 
								BackSecondaryColor="White" BackGradientStyle="TopBottom" BorderWidth="1" 
								BorderColor="26, 59, 105" ImageLocation="~/TempImages/ChartPicPie_#SEQ(300,3)" 
								ImageType="Jpeg" BorderlineWidth="0">
								<titles>
									<asp:Title ShadowColor="32, 0, 0, 0" Font="Trebuchet MS, 14.25pt, style=Bold" ShadowOffset="3" Text="Pie Chart" Name="Title1" ForeColor="26, 59, 105"></asp:Title>
								</titles>
								<legends>
									<asp:Legend BackColor="Transparent" Font="Trebuchet MS, 8pt, style=Bold" 
										Name="Default" LegendStyle="Column" DockedToChartArea="ChartArea1" 
										Title="Forms" TitleAlignment="Near" Docking="Left" Enabled="False" MaximumAutoSize="40" 
										TextWrapThreshold="45"></asp:Legend>
								</legends>
								<series>
									<asp:Series Name="Default" ChartType="Pie" BorderColor="180, 26, 59, 105" 
										Color="220, 65, 140, 240" 
										
										CustomProperties="PieStartAngle=90, PieLabelStyle=Outside, CollectedThresholdUsePercent=False, MinimumRelativePieSize=25"></asp:Series>
								</series>
								<chartareas>
									<asp:ChartArea Name="ChartArea1" BorderColor="64, 64, 64, 64" BackSecondaryColor="Transparent" BackColor="Transparent" ShadowColor="Transparent" BorderWidth="0">
										<area3dstyle Rotation="0" />
										<axisy LineColor="64, 64, 64, 64">
											<LabelStyle Font="Trebuchet MS, 8.25pt, style=Bold" />
											<MajorGrid LineColor="64, 64, 64, 64" />
										</axisy>
										<axisx LineColor="64, 64, 64, 64">
											<LabelStyle Font="Trebuchet MS, 8.25pt, style=Bold" />
											<MajorGrid LineColor="64, 64, 64, 64" />
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

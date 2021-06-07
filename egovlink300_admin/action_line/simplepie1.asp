<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="../includes/SimpleChart.inc" -->
<!-- #include file="../includes/common.asp" //-->

<%

Dim pie
Dim conn
Dim oRs


Set pie = Server.CreateObject("SimpleChart.PieChartGenerator.4")
pie.Caption = "Pie Chart Test"
pie.Width = 1000
pie.Height = 800

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLOLEDB; Data Source=(local);Initial Catalog=SimpleFood;Trusted_Connection=yes", "", ""
'Alternative connection string
'conn.Open "Provider=SQLOLEDB; Data Source=(local);Initial Catalog=SimpleFood;User Id=sa;Password=let_me_in", "", ""

Set rsSales = Server.CreateObject("ADODB.Recordset")
rsSales.Open "SELECT Categories.CategoryName, YearlySales.YearlySales FROM Categories INNER JOIN YearlySales ON Categories.CategoryID = YearlySales.CategoryID ORDER BY Categories.CategoryName", conn, adOpenForwardOnly, adLockReadOnly, adCmdText
While Not rsSales.EOF
	pie.ChartData.Add rsSales("CategoryName"), rsSales("YearlySales"), scFillModeAutoSolid
	rsSales.MoveNext
Wend

rsSales.Close
conn.Close

	
'Display adjustments
pie.ChartData(1).Exploded = True            'Explode the beverages segment
pie.ShowAllValues(True)                     'Show the value of each sale
pie.ValueFormat = "$#.00"                   'Display sales values as currencies
pie.BackColor = RGB(255, 239, 213)			'PapayaWhip

Response.ContentType = "image/jpeg"
Response.BinaryWrite pie.SaveToStream(scImageFormatJpeg)
%>
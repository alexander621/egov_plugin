<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalcategoryselection.asp
' AUTHOR: Steve Loar
' CREATED: 10/11/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Categories to be selected for making a simple reservation.
'
' MODIFICATION HISTORY
' 1.0   10/11/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, iSearchItem

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "create simple reservations", sLevel	' In common.asp


%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalsstyles.css" />

</head>

<body>

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Make Simple Reservations</strong></font><br />
			</p>
			<p id="rentalpagedescription">
				Select a Category from the list below to start the reservation process. 
				You will only be able to make a reservation for one date and time, although you can add 
				other dates and times after you complete the initial reservation.
			</p>
			<!--END: PAGE TITLE-->


<%				'Pull the category list
				ShowRentalCategories
%>			

		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void ShowRentalCategories
'--------------------------------------------------------------------------------------------------
Sub ShowRentalCategories( )
	Dim sSql, oRs, iCount, sClass

	iCount = clng(0)
	sSql = "SELECT recreationcategoryid, ISNULL(categorytitle,'') AS categorytitle, "
	sSql = sSql & "ISNULL(categorydescription,'') AS categorydescription "
	sSql = sSql & "FROM egov_recreation_categories "
	sSql = sSql & "WHERE isforrentals = 1 AND isroot = 0 AND orgid = " & session("orgid")
	sSql = sSql & " ORDER BY categorytitle"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		iCount = iCount + 1
'		If iCount Mod 2 <> 0 Then
'			sClass = " altrow"
'		Else
			sClass = ""
'		End If 

		response.Write vbcrlf & "<p><div class=""rentalcategorygroup"" onClick=""location.href='rentalofferings.asp?categoryid=" & oRs("recreationcategoryid") &"';"">"
		response.write "<table class=""rentalcategory"" cellpadding=""0"" cellspacing=""2"" border=""0"">"
		response.write "<tr class=""categoryrow" & sClass & """><td>"
		response.Write vbcrlf & "<div class=""categorydesc"">"
		response.write oRs("categorytitle") & "<br />"
		response.write "<p class=""categorydesc"">" & oRs("categorydescription") & "</p>"
		response.Write vbcrlf & "</div>"
		response.write "</td></tr></table>"
		response.write vbcrlf & "</div></p>"
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub

%>

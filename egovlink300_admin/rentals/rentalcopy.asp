<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentaldaycopy.asp
' AUTHOR: Steve Loar
' CREATED: 08/13/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Copy the daily schedule from one day to another
'
' MODIFICATION HISTORY
' 1.0   08/13/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, iSearchItem

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "create edit rentals", sLevel	' In common.asp

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalsstyles.css" />

	<script language="Javascript">
	<!--

		function CopyThis()
		{
			document.frmCopyRental.submit();
		}

	//-->
	</script>

</head>

<body>

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Copy A Rental To A New Rental</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->
			<p>
				<input type="button" class="button" value="<< Back" onclick="location.href='rentalslist.asp';" /><br /><br />
			</p>

			<form name="frmCopyRental" method="post" action="rentalcopydo.asp">
				<p>
					Select a Rental to Copy: <% ShowRentalPicks %>
				</p>
				<p>
					<input type="button" class="button" value="Copy To A New Rental" onclick="CopyThis();" />
				</p>
			</form>

		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub ShowRentalPicks()
'--------------------------------------------------------------------------------------------------
Sub ShowRentalPicks()
	Dim sSql, oRs

	sSql = "SELECT R.rentalid, R.rentalname, L.name AS locationname FROM egov_rentals R, egov_class_location L "
	sSql = sSql & "WHERE R.orgid = " & session("orgid") & " AND R.locationid = L.locationid "
	sSql = sSql & "ORDER BY R.rentalname, L.name"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""rentalid"">"
	Do While Not oRs.EOF
		response.write vbcrlf & vbtab & "<option value=""" & oRs("rentalid") & """ >" & oRs("rentalname") & " (" & oRs("locationname") & ")</option>"
		oRs.MoveNext 
	Loop
	response.write vbcrlf & "</select>"
	
	oRs.Close
	Set oRs = Nothing 
End Sub 



%>
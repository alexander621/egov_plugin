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
Dim iRentalId

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "create edit rentals", sLevel	' In common.asp

iRentalId = CLng(request("rentalid"))

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalsstyles.css" />

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>


	<script language="Javascript">
	<!--
		
		function CopyThis()
		{
			if ($("sourcedayid").selectedIndex == $("targetdayid").selectedIndex)
			{
				alert("You have selected the same day as both the source and target.\nPlease correct this and try again.");
				return false;
			}
			//alert("good to go");
			document.frmCopyDay.submit();
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
				<font size="+1"><strong>Daily Schedule/Rates Copy</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<input type="button" class="button" value="<< Back" onclick="location.href='rentaledit.asp?rentalid=<%=iRentalId%>';" />

			<p>
				Select a source day and a target day. The information for the target day will be replaced by the 
				information from the source day.<br /><br />
			</p>

			<form name="frmCopyDay" method="post" action="rentaldaycopydo.asp">
				<input type="hidden" name="rentalid" value="<%=iRentalId%>" />
				<table id="copydaytable" cellpadding="2" cellspacing="0" border="0">
					<tr><th>Source Day</th><th>Target Day</th></tr>
					<tr>
						<td align="center"><% ShowDayPicks "sourcedayid", iRentalId %></td>
						<td align="center"><% ShowDayPicks "targetdayid", iRentalId %></td>
					</tr>
				</table>

				<p>
					<input type="button" class="button" value="Copy Day" onclick="CopyThis();" />
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
' Sub ShowDayPicks( sSelectName, iRentalId )
'--------------------------------------------------------------------------------------------------
Sub ShowDayPicks( ByVal sSelectName, ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT dayid, isoffseason, weekdayname FROM egov_rentaldays WHERE orgid = " & session("orgid") & " AND rentalid = " & iRentalId & " ORDER BY isoffseason, dayofweek"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select id=""" & sSelectName & """ name=""" & sSelectName & """>"
	Do While Not oRs.EOF
		response.write vbcrlf & vbtab & "<option value=""" & oRs("dayid") & """ >" & oRs("weekdayname") & " ("
		If oRs("isoffseason") Then
			response.write "Off Season"
		Else
			response.write "In Season"
		End If 
		response.write ")</option>"
		oRs.MoveNext 
	Loop
	response.write vbcrlf & "</select>"
	
	oRs.Close
	Set oRs = Nothing 
End Sub 



%>
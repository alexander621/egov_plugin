<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitvaluationtypelist.asp
' AUTHOR: Steve Loar
' CREATED: 04/14/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of permit fixture types
'
' MODIFICATION HISTORY
' 1.0   04/14/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch

sLevel = "../" ' Override of value from common.asp

PageDisplayCheck "permit valuation types", sLevel	' In common.asp

If request("searchtext") = "" Then
	sSearch = ""
Else
	sSearch = request("searchtext")
End If 

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script language="javascript" src="../scripts/modules.js"></script>

	<script language="Javascript">
	<!--

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
				<font size="+1"><strong>Permit Valuation Types</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->
			<form name="frmFixtureSearch" method="post" action="permitfixturetypelist.asp">
				<div>
					<input type="text" name="searchtext" value="<%=Replace(sSearch,"""","&quot;")%>" size="50" maxlength="150" /> &nbsp; &nbsp;
					<input type="submit" class="button ui-button ui-widget ui-corner-all" value="Search" />
					&nbsp; &nbsp; <input type="button" name="new" class="button ui-button ui-widget ui-corner-all" value="New Valuation Type" onclick="location.href='permitvaluationtypeedit.asp?permitvaluationtypeid=0';" />
					<br /><br />
				</div>
			</form>

			<div class="shadow">
			<table id="categorytypes" cellpadding="0" cellspacing="0" border="0">
				<tr><th>Valuation Types</th></tr>
				<%	
					ShowPermitValuationTypes session("orgid"), sSearch 
				%>
			</table>
			</div>
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
' Sub ShowPermitValuationTypes( iOrgid )
'--------------------------------------------------------------------------------------------------
Sub ShowPermitValuationTypes( iOrgid, sSearch )
	Dim sSql, oRates, iRowCount, iPermitValuationTypeid

	iRowCount = 0
	iPermitValuationTypeid = CLng(0)

	sSql = "SELECT permitvaluationtypeid, permitvaluation "
	sSql = sSql & " FROM egov_permitvaluationtypes "
	sSql = sSql & " WHERE orgid = "& iOrgid 
	If sSearch <> "" Then
		sSql = sSql & " AND permitvaluation LIKE '%" & dbsafe(sSearch) & "%' "
	End If 
	sSql = sSql & " ORDER BY permitvaluation, permitvaluationtypeid"

	Set oRates = Server.CreateObject("ADODB.Recordset")
	oRates.Open sSQL, Application("DSN"), 3, 1

	If Not oRates.EOF Then
		Do While Not oRates.EOF
			'If iPermitValuationTypeid <> CLng(oRates("permitvaluationtypeid")) Then 
				'If iPermitValuationTypeid > CLng(0) Then
				'	response.write "</tr>"
				'End If 
				'iPermitValuationTypeid = CLng(oRates("permitvaluationtypeid"))
				iRowCount = iRowCount + 1
				response.write vbcrlf & "<tr id=""" & iRowCount & """"
				If iRowCount Mod 2 = 0 Then
					response.write " class=""altrow"" "
				End If 
				response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
				response.write "<td class=""leftcol"" onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='permitvaluationtypeedit.asp?permitvaluationtypeid=" & oRates("permitvaluationtypeid") & "';"">&nbsp;" & oRates("permitvaluation") & "</td>"
				response.write "</tr>"
			'End If 
			oRates.MoveNext 
		Loop 
		
	Else
		If sSearch <> "" Then
			response.write vbcrlf & "<tr><td>&nbsp;No Valuation Types could be found that match your search criteria.</td></tr>"
		Else 
			response.write vbcrlf & "<tr><td>&nbsp;No Valuation Types could be found. Click on the New Valuation Type button to start entering data.</td></tr>"
		End If 
	End If  
	
	oRates.Close
	Set oRates = Nothing 

End Sub 



%>

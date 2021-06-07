<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: constructiontypelist.asp
' AUTHOR: Steve Loar
' CREATED: 12/11/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   12/11/07	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch

sLevel = "../" ' Override of value from common.asp

PageDisplayCheck "construction type rates", sLevel	' In common.asp

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
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
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
		<div class="gutters">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Construction Type Rates</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<form name="frmFixtureSearch" method="post" action="constructiontypelist.asp">
			<div>
				<input type="text" name="searchtext" value="<%=Replace(sSearch,"""","&quot;")%>" size="50" maxlength="150" /> &nbsp; &nbsp;
				<input type="submit" class="button ui-button ui-widget ui-corner-all" value="Search" />
				&nbsp; &nbsp; <input type="button" name="new" class="button ui-button ui-widget ui-corner-all" value="New Rates" onclick="location.href='constructiontypeedit.asp?occupancytypeid=0';" />
				<br /><br />
			</div>
			</form>

			<div id="constructiontypesshadow" class="shadow">
			<table id="constructiontypes" cellpadding="0" cellspacing="0" border="0">
				<tr style="background:#93bee1;"><th style="border-right:1px solid black;">&nbsp;</th><th align="center" colspan="9">Type of Construction</th></tr>
				<%	
					If ShowConstructionTypeHeaderRow( session("orgid") ) Then 
						ShowConstructionTypeRates session("orgid"), sSearch
					Else
						response.write "<tr><td colspan=""10"">&nbsp;No data exists for this feature. Contact EC Link for assistance.</td></tr>"
					End If 
				%>
			</table>
			</div>
			<p>*NP = Not Permitted</p>
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
' Function ShowConstructionTypeHeaderRow( iOrgid ) 
'--------------------------------------------------------------------------------------------------
Function ShowConstructionTypeHeaderRow( iOrgid ) 
	Dim sSql, oHeader, bReturn

	bReturn = False 
	sSql = "SELECT constructiontype FROM egov_constructiontypes WHERE orgid = " & iOrgid & " ORDER BY displayorder"

	Set oHeader = Server.CreateObject("ADODB.Recordset")
	oHeader.Open sSQL, Application("DSN"), 3, 1

	If Not oHeader.EOF Then 
		bReturn = True 
		response.write vbcrlf & "<tr style=""background:#93bee1""><th class=""leftcol"">Use Group</th>"
		Do While Not oHeader.EOF 
			response.write "<th>" & oHeader("constructiontype") & "</th>"
			oHeader.MoveNext
		Loop 
		response.write "</tr>"
	End If 

	oHeader.Close
	Set oHeader = Nothing 

	ShowConstructionTypeHeaderRow = bReturn 
End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowConstructionTypeRates( iOrgid )
'--------------------------------------------------------------------------------------------------
Sub ShowConstructionTypeRates( iOrgid, sSearch )
	Dim sSql, oRates, iRowCount, iOccupancytypeid

	iRowCount = 0
	iOccupancytypeid = CLng(0)
	sSql = "SELECT F.occupancytypeid, F.constructiontypeid, O.usegroupcode, O.occupancytype, F.constructiontyperate, F.isnotpermitted "
	sSql = sSql & " FROM egov_constructionfactors F, egov_occupancytypes O, egov_constructiontypes T "
	sSql = sSql & " WHERE F.occupancytypeid = O.occupancytypeid AND T.orgid = O.orgid and F.constructiontypeid = T.constructiontypeid AND O.orgid = "& iOrgid 
	If sSearch <> "" Then
		sSql = sSql & " AND (O.occupancytype LIKE '%" & dbsafe(sSearch) & "%' OR O.usegroupcode LIKE '%" & dbsafe(sSearch) & "%') "
	End If 
	sSql = sSql & " ORDER BY O.usegroupcode, O.occupancytype, T.displayorder"

	Set oRates = Server.CreateObject("ADODB.Recordset")
	oRates.Open sSQL, Application("DSN"), 3, 1

	If Not oRates.EOF Then
		Do While Not oRates.EOF
			If iOccupancytypeid <> CLng(oRates("occupancytypeid")) Then 
				If iOccupancytypeid > CLng(0) Then
					response.write "</tr>"
				End If 
				iOccupancytypeid = CLng(oRates("occupancytypeid"))
				iRowCount = iRowCount + 1
				response.write vbcrlf & "<tr id=""" & iRowCount & """"
				If iRowCount Mod 2 = 0 Then
					response.write " class=""altrow"" "
				End If 
				response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
				response.write "<td class=""leftcol"" onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='constructiontypeedit.asp?occupancytypeid=" & oRates("occupancytypeid") & "';"">&nbsp;" & oRates("usegroupcode") & "&nbsp;" & oRates("occupancytype") & "</td>"
			End If 
			' output the rates
			response.write "<td align=""center"" onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='constructiontypeedit.asp?occupancytypeid=" & oRates("occupancytypeid") & "';"">" 
			If Not oRates("isnotpermitted") Then 
				response.write FormatNumber(oRates("constructiontyperate"),4,,,0)
			Else
				response.write "NP"
			End If 
			response.write "</td>"
			oRates.MoveNext 
		Loop 
		response.write "</tr>"
	Else
		response.write vbcrlf & "<tr><td colspan=""10"">&nbsp;Click on the New Rates button to start entering data.</td></tr>"
	End If  
	
	oRates.Close
	Set oRates = Nothing 

End Sub 



%>

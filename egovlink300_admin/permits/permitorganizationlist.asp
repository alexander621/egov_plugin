<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitorganizationlist.asp
' AUTHOR: Steve Loar
' CREATED: 02/13/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of permit organizations
'
' MODIFICATION HISTORY
' 1.0   02/13/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, sWhere, iViewChoice

sLevel = "../" ' Override of value from common.asp

PageDisplayCheck "permit organizations", sLevel	' In common.asp

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
				<font size="+1"><strong>Permit Organizations</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->
			<form name="frmPermitSearch" method="post" action="permitcontactlist.asp">
				<div>
					<input type="text" name="searchtext" value="<%=Replace(sSearch,"""","&quot;")%>" size="50" maxlength="150" /> &nbsp; &nbsp;
					<input type="submit" class="button ui-button ui-widget ui-corner-all" value="Search" /> &nbsp; &nbsp; 
					<input type="button" name="new" class="button ui-button ui-widget ui-corner-all" value="New Organization" onclick="location.href='permitcontacttypeedit.asp?permitcontacttypeid=0&isorganization=1';" />
					<br /><br />
				</div>
			</form>

			<div class="shadow">
			<table id="categorytypes" cellpadding="0" cellspacing="0" border="0">
				<tr><th>Organization</th></tr>
				<%	
					ShowPermitContactTypes sSearch, sWhere
				%>
			</table>
			</div>
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
' Sub ShowPermitContactTypes( sSearch, sWhere )
'--------------------------------------------------------------------------------------------------
Sub ShowPermitContactTypes( sSearch, sWhere )
	Dim sSql, oRs, iRowCount, iPermitContactTypeid, sContractorType

	iRowCount = 0
	iPermitContactTypeid = CLng(0)

	sSql = "SELECT permitcontacttypeid, ISNULL(company,'') AS company, ISNULL(firstname,'') AS firstname, "
	sSql = sSql & " ISNULL(lastname,'') AS lastname, ISNULL(lastname,'') + ISNULL(firstname,'') + ISNULL(company,'') AS sortname, "
	sSql = sSql & " ISNULL(contractortypeid,0) AS contractortypeid "
	sSql = sSql & " FROM egov_permitcontacttypes WHERE isorganization = 1 AND orgid = " & session("orgid") & sWhere
	If sSearch <> "" Then
		sSql = sSql & " AND (lastname LIKE '%" & dbsafe(sSearch) & "%' OR firstname LIKE '%" & dbsafe(sSearch) & "%' OR company LIKE '%" & dbsafe(sSearch) & "%' ) "
	End If 
	sSql = sSql & " ORDER BY 5"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		Do While Not oRs.EOF
			If iPermitContactTypeid <> CLng(oRs("permitcontacttypeid")) Then 
				If iPermitContactTypeid > CLng(0) Then
					response.write "</tr>"
				End If 
				iPermitContactTypeid = CLng(oRs("permitcontacttypeid"))
				iRowCount = iRowCount + 1
				response.write vbcrlf & "<tr id=""" & iRowCount & """"
				If iRowCount Mod 2 = 0 Then
					response.write " class=""altrow"" "
				End If 
				response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
				response.write "<td class=""leftcol"" title=""click to edit"" onClick=""location.href='permitcontacttypeedit.asp?permitcontacttypeid=" & oRs("permitcontacttypeid") & "';"">&nbsp;" 
				
				If oRs("firstname") = "" Then
					response.write Replace(oRs("company"),"""","&quot;")
				Else 
					response.write oRs("firstname") & " " & oRs("lastname")
					If oRs("company") <> "" Then 
						response.write "&nbsp;( " & Replace(oRs("company"),"""","&quot;") & " )"
					End If 
				End If 

				response.write "</td>"
			End If 
			oRs.MoveNext 
		Loop 
		response.write "</tr>"
	Else
		If sSearch <> "" Then
			response.write vbcrlf & "<tr><td colspan=""2"">&nbsp;No Permit Organizations could be found that match your search criteria.</td></tr>"
		Else 
			response.write vbcrlf & "<tr><td colspan=""2"">&nbsp;No Permit Organizations could be found. Click on the New Organization button to start entering data.</td></tr>"
		End If 
	End If  
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Function ContactHasPermits( iPermitContactTypeId )
'--------------------------------------------------------------------------------------------------
Function ContactHasPermits( iPermitContactTypeId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(permitcontacttypeid) AS hits FROM egov_permitcontacts "
	sSql = sSql & " WHERE permitcontacttypeid = " & iPermitContactTypeId
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then 
			ContactHasPermits = True 
		Else
			ContactHasPermits = False 
		End If 
	Else
		ContactHasPermits = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


%>

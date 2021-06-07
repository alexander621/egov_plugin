<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitinspectiontypelist.asp
' AUTHOR: Steve Loar
' CREATED: 01/14/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of permit inspection types
'
' MODIFICATION HISTORY
' 1.0   01/14/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, iPermitCategoryId

sLevel = "../" ' Override of value from common.asp

PageDisplayCheck "permitv2 types", sLevel	' In common.asp

If request("searchtext") = "" Then
	sSearch = ""
Else
	sSearch = request("searchtext")
End If 

If request("permitcategoryid") <> "" Then
	iPermitCategoryId = request("permitcategoryid")
Else 
	iPermitCategoryId = "0"
End If 

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />

	<script language="javascript" src="../scripts/modules.js"></script>

	<script language="Javascript">
	<!--
		
		function refreshPage()
		{
			document.frmPermitSearch.submit();
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
		<div class="gutters">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Permit Types</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->
			<form name="frmPermitSearch" method="post" action="permittypelist.asp">
				<div id="functionlinks">
					<input type="text" name="searchtext" value="<%=Replace(sSearch,"""","&quot;")%>" size="50" maxlength="150" /> &nbsp; &nbsp;
					<input type="submit" class="button ui-button ui-widget ui-corner-all" value="Search" />
					&nbsp; &nbsp; <input type="button" name="new" class="button ui-button ui-widget ui-corner-all" value="New Permit Type" onclick="location.href='permittypeedit.asp?permittypeid=0';" />
					<br /><br />Filter by Category:
<%					ShowPermitCategoryPicks iPermitCategoryId	%>
				</div>
			</form>

			<table id="categorytypes" cellpadding="0" cellspacing="0" border="0">
				<tr><th>Permit Type</th><th>Permit Category</th></tr>
				<%	
					ShowPermitTypes sSearch, iPermitCategoryId
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
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void ShowPermitTypes sSearch, iPermitCategoryId
'--------------------------------------------------------------------------------------------------
Sub ShowPermitTypes( ByVal sSearch, ByVal iPermitCategoryId )
	Dim sSql, oRs, iRowCount, iPermitTypeid

	iRowCount = 0
	iPermitTypeid = CLng(0)

	sSql = "SELECT T.permittypeid, ISNULL(T.permittype,'') AS permittype, ISNULL(T.permittypedesc,'') AS permittypedesc, "
	sSql = sSql & " C.permitcategory FROM egov_permittypes T, egov_permitcategories C "
	sSql = sSql & " WHERE T.orgid = "& session("orgid") & " AND T.permitcategoryid = C.permitcategoryid "
	If sSearch <> "" Then
		sSql = sSql & " AND (T.permittype LIKE '%" & dbsafe(sSearch) & "%' OR T.permittypedesc LIKE '%" & dbsafe(sSearch) & "%' ) "
	End If 
	If CLng(iPermitCategoryId) > CLng(0) Then
		sSql = sSql & " AND T.permitcategoryid = " & iPermitCategoryId
	End If 
	sSql = sSql & " ORDER BY T.permittype, T.permittypedesc, T.permittypeid"

	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		Do While Not oRs.EOF
			If iPermitTypeid <> CLng(oRs("permittypeid")) Then 
				If iPermitTypeid > CLng(0) Then
					response.write "</tr>"
				End If 
				iPermitTypeid = CLng(oRs("permittypeid"))
				iRowCount = iRowCount + 1
				response.write vbcrlf & "<tr id=""" & iRowCount & """"
				If iRowCount Mod 2 = 0 Then
					response.write " class=""altrow"" "
				End If 
				response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
				response.write "<td class=""leftcol"" title=""click to edit"" onClick=""location.href='permittypeedit.asp?permittypeid=" & oRs("permittypeid") & "';"">&nbsp;" & oRs("permittype") 
				If oRs("permittype") <> "" And oRs("permittypedesc") <> "" Then 
					response.write " &ndash; " 
				End If 
				response.write oRs("permittypedesc") & "</td>"

				response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='permittypeedit.asp?permittypeid=" & oRs("permittypeid") & "';"">&nbsp;"
				response.write oRs("permitcategory")
				response.write "</td>"
			End If 
			oRs.MoveNext 
		Loop 
		response.write "</tr>"
	Else
		If sSearch <> "" Then
			response.write vbcrlf & "<tr><td colspan=""2"">&nbsp;No Permit Types could be found that match your search criteria.</td></tr>"
		Else 
			response.write vbcrlf & "<tr><td colspan=""2"">&nbsp;No Permit Types could be found. Click on the New Permit Type button to start entering data.</td></tr>"
		End If 
	End If  
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' ShowPermitCategoryPicks iPermitCategoryId	
'--------------------------------------------------------------------------------------------------
Sub ShowPermitCategoryPicks( ByVal iPermitCategoryId )
	Dim sSql, oRs

	sSQL = "SELECT permitcategoryid, permitcategory FROM egov_permitcategories "
	sSql = sSql & " WHERE orgid = " & SESSION("orgid")
	sSql = sSql & " ORDER BY permitcategory"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	response.write vbcrlf & "<select id=""permitcategoryid"" name=""permitcategoryid"" onchange=""refreshPage();"">"
	response.write vbcrlf & "<option value=""0"">View All Categories</option>"

	Do While NOT oRs.EOF 
		response.write vbcrlf & "<option value=""" & oRs("permitcategoryid") & """ "  
		If CLng(iPermitCategoryId) = CLng(oRs("permitcategoryid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("permitcategory")
		response.write "</option>"
		oRs.MoveNext
	Loop

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 


%>

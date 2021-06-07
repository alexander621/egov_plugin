<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: featureordering.asp
' AUTHOR: Steve Loar
' CREATED: 09/12/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of permit reviews
'
' MODIFICATION HISTORY
' 1.0   09/12/2008	Steve Loar - INITIAL VERSION
' 1.1	04/11/2011	Steve Loar - Added mobile feature ordering
' 1.2	10/20/2011	Steve Loar - Converted to HTML5
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iParentFeatureId, iAdminDisplayOrder

sLevel = "../" ' Override of value from common.asp

If Not UserIsRootAdmin( session("UserID") ) Then
	response.redirect "../default.asp"
End If 

If request("parentfeatureid") <> "" Then 
	iParentFeatureId = CLng(request("parentfeatureid"))
Else
	'iParentFeatureId = GetFirstParentFeatureId()
	iParentFeatureId = CLng(0)
End If 

If CLng(iParentFeatureId) = CLng(0) Then
	If request("admindisplayorder") <> "" Then 
		iAdminDisplayOrder = clng(request("admindisplayorder"))
	Else 
		iAdminDisplayOrder = 1
	End If 
Else 
	iAdminDisplayOrder = 1
End If 
%>

<html lang="en">
<head>
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

	<title>E-Gov Organization Feature Ordering</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="admin.css" />
	
	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

	<script language="javascript" src="../scripts/tablednd.js"></script>

	<script language="Javascript">
	<!--

		function ShowFeatures()
		{
			document.pickForm.submit();
		}

		function Init()
		{
			var table = $('featurelist');
			var tableDnD = new TableDnD();
			tableDnD.init(table);

			// Redefine the onDrop so that we can update things
			tableDnD.onDrop = function(table, row) 
			{
				var iRowNo = -1;
				var rows = this.table.tBodies[0].rows;
				var debugStr = 'rows now: ';
				for (var i=0; i<rows.length; i++) 
				{
					iRowNo += 1;
					// skip the header row
					if (iRowNo > 0)
					{
						debugStr += iRowNo + ' = ' + rows[i].id + '\n';
						rows[i].cells[1].innerHTML = iRowNo; // change the display order
						rows[i].className = iRowNo & 1? '':'altrow';  // set the row background class
						//alert($("oldrow" + rows[i].id).value + ": " + $("featureid" + rows[i].id).value + " now " + iRowNo);
						// Fire off ajax routine here to reorder the rows to this order
						doAjax('featureorderingupdate.asp', 'featureid=' + $("featureid" + rows[i].id).value + '&displayorder=' + iRowNo + '&admindisplayorder=<%=iAdminDisplayOrder%>', '', 'get', '0');
						$("oldrow" + rows[i].id).value = iRowNo;
					}
				}
			}

		}

	//-->
	</script>

</head>

<body onload="Init();">

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Feature Ordering</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<p>
				<input type="button" class="button" name="back" value="<< Back to Feature Management" onClick="javascript:window.location='managefeatures.asp'" /> &nbsp;&nbsp;
			</p>
			<form name="pickForm" method="post" action="featureordering.asp">
				<p>
					<%	ShowParentFeatures iParentFeatureId  %>&nbsp;&nbsp;
<%
					If CLng(iParentFeatureId) = CLng(0) Then		
						response.write vbcrlf & "<select name=""admindisplayorder"" onchange=""ShowFeatures();"">"
						
						response.write vbcrlf & "<option value=""1"""
						If clng(iAdminDisplayOrder) = clng(1) Then
							response.write " selected=""selected"" "
						End If 
						response.write ">Admin Display Order</option>"

						response.write vbcrlf & "<option value=""0"""
						If clng(iAdminDisplayOrder) = clng(0) Then
							response.write " selected=""selected"" "
						End If 
						response.write ">Public Display Order</option>"

						response.write vbcrlf & "<option value=""-1"""
						If clng(iAdminDisplayOrder) = clng(-1) Then
							response.write " selected=""selected"" "
						End If 
						response.write ">Mobile Display Order</option>"

						response.write vbcrlf & "</select>"
					Else
						response.write vbcrlf & "<input type=""hidden"" name=""admindisplayorder"" value=""1"" />"
					End If										
%>

				</p>
			</form>

<%	
			ShowFeatures iParentFeatureId, iAdminDisplayOrder		
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
' void ShowParentFeatures iFeatureId 
'--------------------------------------------------------------------------------------------------
Sub ShowParentFeatures( ByVal iFeatureId )
	Dim sSql, oRs

	sSql = "SELECT featureid, featurename FROM egov_organization_features WHERE parentfeatureid = 0 ORDER BY admindisplayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""parentfeatureid"" onchange='ShowFeatures();'>"
		response.write vbcrlf & "<option value=""0"""
		If CLng(iFeatureId) = CLng(0) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">Top Level Features</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("featureid") & """"
			If CLng(iFeatureId) = CLng(oRs("featureid")) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("featurename") & "</option>"
			oRs.MoveNext 
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' integer GetFirstParentFeatureId()
'--------------------------------------------------------------------------------------------------
Function GetFirstParentFeatureId()
	Dim sSql, oRs

	sSql = "SELECT featureid FROM egov_organization_features WHERE parentfeatureid = 0 ORDER BY admindisplayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetFirstParentFeatureId = CLng(oRs("featureid"))
	Else
		GetFirstParentFeatureId = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowFeatures iParentFeatureId 
'--------------------------------------------------------------------------------------------------
Sub ShowFeatures( ByVal iParentFeatureId, ByVal iAdminDisplayOrder )
	Dim sSql, oRs, iRowCount, sOrderBy, sWhere

	iRowCount = 0
	sWhere = ""

	If CLng(iParentFeatureId) = CLng(0) Then
		If clng(iAdminDisplayOrder) = clng(1) Then 
			sOrderBy = "admindisplayorder"
		Else
			If clng(iAdminDisplayOrder) = clng(0) Then 
				sOrderBy = "publicdisplayorder"
				sWhere = " AND haspublicview = 1 "
			Else
				sOrderBy = "mobiledefaultdisplayorder"
				sWhere = " AND hasmobileview = 1 "
			End If 
		End If 
	Else
		sOrderBy = "securitydisplayorder"
	End If 

	sSql = "SELECT featureid, featurename, featuretype FROM egov_organization_features "
	sSql = sSql & " WHERE parentfeatureid = " & iParentFeatureId & sWhere & " ORDER BY " & sOrderBy

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		'response.write vbcrlf & "<div class=""shadow"" id=""featurelistshadow"">"
		response.write vbcrlf & "<table id=""featurelist"" cellpadding=""2"" cellspacing=""0"" border=""0"">"
		response.write vbcrlf & "<tr noDrop=""true"" noDrag=""true""><th>&nbsp;</th><th>Display<br />Order</th><th>FeatureId</th><th>Feature</th><th>Feature<br />Type</th></tr>"
		
		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
			' Drag and Drop Icon
			response.write vbcrlf & "<tr id=""" & iRowCount & """"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write ">"
			response.write "<td align=""center""><img src=""..\images\up_down_arrow.gif"" class=""DRAGIMG"" width=""13"" height=""19"" border=""0"" alt=""drag and drop"" />"
			response.write "<input type=""hidden"" id=""featureid" & iRowCount & """ name=""featureid" & iRowCount & """ value=""" & oRs("featureid") & """ />"
			response.write "<input type=""hidden"" id=""oldrow" & iRowCount & """ name=""oldrow" & iRowCount & """ value=""" & iRowCount & """ />"
			response.write "</td>"

			' Order
			response.write "<td align=""center"">" & iRowCount & "</td>"

			' FeatureId
			response.write "<td align=""center"">" & oRs("featureid") & "</td>"

			' Feature
			response.write "<td>" & oRs("featurename") & "</td>"

			' Feature Type
			If UCASE(oRs("featuretype")) = "S" Then 
				lcl_featuretype = "Security"
			Else 
				lcl_featuretype = "Navigation"
			End If 

			response.write "<td align=""center"">" & lcl_featuretype & "</td>"

 			response.write "</tr>"

			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</table>"
	Else
		response.write "<p>There are no sub-features for the selected feature.</p>"
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' boolean FeatureIsTopLevel( iFeatureId )
'-------------------------------------------------------------------------------------------------
Function FeatureIsTopLevel( ByVal iFeatureId )
	Dim sSql, oRs

	sSql = "SELECT parentfeatureid FROM egov_organization_features WHERE featureid = " & iFeatureId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If CLng(oRs("parentfeatureid")) = CLng(0) Then
		FeatureIsTopLevel = True 
	Else
		FeatureIsTopLevel = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 



%>

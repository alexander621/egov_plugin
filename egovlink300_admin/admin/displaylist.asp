<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: displaylist.asp
' AUTHOR: Steve Loar
' CREATED: 04/30/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of displays. From here you can create or edit displays. 
'				This is not the org setup list
'
' MODIFICATION HISTORY
' 1.0   04/30/2010	Steve Loar - INITIAL VERSION
' 1.1	10/20/2011	Steve Loar - Converted to HTML5
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iFeatureId

sLevel = "../"  'Override of value from common.asp

If Not UserIsRootAdmin( session("UserID") ) Then 
	response.redirect "../default.asp"
End If 

If request("featureid") <> "" Then
	iFeatureId = CLng(request("featureid"))
Else
	iFeatureId = 0
End If 

If request("msg") <> "" Then
	If request("msg") = "d" Then
		sLoadMsg = "displayScreenMsg('The Display Was Successfully Deleted');"
	End If
End If 

%>
<html lang="en">
<head>
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="admin.css" />

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

	<script language="javascript" src="../scripts/modules.js"></script>

	<script language="Javascript">
	<!--

		function RefreshResults()
		{
			document.frmDisplaySearch.submit();
		}

		function SetUpPage()
		{
			<%=sLoadMsg%>
		}

		function displayScreenMsg(iMsg) 
		{
			if(iMsg!="") 
			{
				$("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("screenMsg").innerHTML = "";
		}

	//-->
	</script>

</head>

<body onload="SetUpPage();">

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Create/Edit Displays</strong></font><br />
			</p>
			
			<p>This is where displays are created and edited. This is not where the setup for a client is done.</p>

			<!--END: PAGE TITLE-->
			<table id="screenMsgtable"><tr><td>
				<span id="screenMsg"></span>
			</td></tr></table>

			<!--BEGIN: FILTER SELECTION-->
			<div class="filterselection">
				<fieldset class="filterselection">
				   <legend class="filterselection">Search Options</legend>
					<p>
						<form name="frmDisplaySearch" method="post" action="displaylist.asp">
							<table cellpadding="2" cellspacing="0" border="0">
								<tr>
									<td>Category:</td><td><% ShowCategoryPicks iFeatureId %></td>
								</tr>
								<tr>
			    					<td colspan="2"><input class="button" type="button" value="Refresh Results" onclick="RefreshResults();" /></td>
  								</tr>
							</table>
						</form>
					</p>
				</fieldset>
			</div>
			<!--END: FILTER SELECTION-->

			<input type="button" class="button" id="newdisplaybutton" name="newdisplaybutton" value="Create A New Display" onclick="location.href='displayedit.asp?displayid=0';" /> &nbsp;&nbsp; 

<%			ShowDisplays iFeatureId			%>

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
' void ShowDisplays iFeatureId
'--------------------------------------------------------------------------------------------------
Sub ShowDisplays( ByVal iFeatureId )
	Dim sSql, oRs, sWhere, iCount

	If CLng(iFeatureId) <> CLng(0) Then
		sWhere = " WHERE featureid = " & iFeatureId & " "
	Else
		sWhere = ""
	End If 

	iCount = 0

	sSql = "SELECT displayid, displayname, ISNULL(featureid,0) AS featureid FROM egov_organization_displays " & sWhere
	sSql = sSql & " ORDER BY displayname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<table id=""displayslist"" cellpadding=""1"" cellspacing=""0"" border=""0"">"
		response.write vbcrlf & "<tr><th>Display Name</th><th>Category</th></tr>"
		Do While Not oRs.EOF
			iCount = iCount + 1
			response.write vbcrlf & "<tr id=""" & iRowCount & """"
			If iCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"

			response.write "<td class=""firstcol"" align=""left"" title=""click to edit"" onClick=""location.href='displayedit.asp?displayid=" & oRs("displayid") & "';"" nowrap=""nowrap"">"
			response.write oRs("displayname") & "</td>"

			response.write "<td align=""center"">"
			If CLng(oRs("featureid")) > CLng(0) Then 
				response.write GetFeatureName( oRs("featureid") )
			Else
				response.write "&nbsp;"
			End If 
			
			response.write "</td>"

			response.write "</tr>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</table>"
	Else
		response.write vbcrlf & "<p>No Displays were found to match your search criteria.</p>"
	End If 

	oRs.Close
	Set oRs = Nothing
	
End Sub


'--------------------------------------------------------------------------------------------------
' void ShowCategoryPicks iFeatureId
'--------------------------------------------------------------------------------------------------
Sub ShowCategoryPicks( ByVal iFeatureId )
	Dim sSql, oRs

	sSql = "SELECT featureid, featurename FROM egov_organization_features "
	sSql = sSql & "WHERE parentfeatureid = 0 AND featuretype = 'N' "
	sSql = sSql & "AND LOWER(featurename) NOT IN ('log off','e-gov administration') "
	sSql = sSql & "ORDER BY admindisplayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select name=""featureid"">"
	response.write vbcrlf & "<option value=""0"">All Categories</option>"

	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("featureid") & """ "
		If clng(oRs("featureid")) = clng(iFeatureId) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("featurename") & "</option>"
		oRs.MoveNext 
	Loop

	response.write vbcrlf & "</select>"
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' string GetFeatureName( iFeatureId )
'--------------------------------------------------------------------------------------------------
Function GetFeatureName( ByVal iFeatureId )
	Dim sSql, oRs

	sSql = "SELECT featurename FROM egov_organization_features "
	sSql = sSql & "WHERE featureid = " & iFeatureId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetFeatureName = oRs("featurename") 
	Else
		GetFeatureName = ""
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Function 

%>

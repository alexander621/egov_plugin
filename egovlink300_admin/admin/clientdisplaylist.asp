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
				<font size="+1"><strong>Setup Client Displays</strong></font><br />
			</p>
			
			<p>This is where the displays are setup for a client. You should be logged into the client's admin site to use this.</p>

			<!--END: PAGE TITLE-->
			<table id="screenMsgtable"><tr><td>
				<span id="screenMsg"></span>
			</td></tr></table>

			<!--BEGIN: FILTER SELECTION-->
			<div class="filterselection">
				<fieldset class="filterselection">
				   <legend class="filterselection">Search Options</legend>
					<p>
						<form name="frmDisplaySearch" method="post" action="clientdisplaylist.asp">
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
		sWhere = " AND D.featureid = " & iFeatureId & " "
	Else
		sWhere = ""
	End If 

	iCount = 0

	sSql = "SELECT D.displayid, D.displayname, ISNULL(D.featureid,0) AS featureid, ISNULL(F.featurename,'') AS featurename "
	sSql = sSql & "FROM egov_organization_displays D LEFT OUTER JOIN egov_organization_features F ON D.featureid = F.featureid "
	sSql = sSql & "WHERE D.admincanedit = 1 " & sWhere
	sSql = sSql & " ORDER BY F.featurename, D.displayname"
'	response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<table id=""displayslist"" cellpadding=""1"" cellspacing=""0"" border=""0"">"
		response.write vbcrlf & "<tr><th>Category</th><th>Display Name</th></tr>"
		Do While Not oRs.EOF
			iCount = iCount + 1
			response.write vbcrlf & "<tr id=""" & iRowCount & """"
			If iCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"

			response.write "<td align=""center"">"
			If CLng(oRs("featureid")) > CLng(0) Then 
				response.write oRs("featurename")
			Else
				response.write "&nbsp;"
			End If 
			response.write "</td>"

			response.write "<td class=""firstcol"" align=""left"" title=""click to edit"" onClick=""location.href='clientdisplayedit.asp?displayid=" & oRs("displayid") & "';"" nowrap=""nowrap"">"
			response.write oRs("displayname") & "</td>"

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


%>

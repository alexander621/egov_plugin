<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: season_list.asp
' AUTHOR: Steve Loar
' CREATED: 02/27/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0	02/27/2007	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iSeasonClosed, iOtherSeasonPick

'Check to see if the feature is offline
if isFeatureOffline("activities") = "Y" then
   response.redirect "../admin/outage_feature_offline.asp"
end if

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "seasons" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

If request("seasonclosed") = "" Or clng(request("seasonclosed")) = clng(1) Then
	iSeasonClosed = 0
	sShowText = "Show All Seasons"
	sWhere = " and isclosed = 0 "
Else
	iSeasonClosed = 1
	sShowText = "Show Open Seasons"
	sWhere = "" 
End If 

%>

<html>
<head>
 	<title>E-Gov Administration Console</title>

	 <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

 	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
 	<link rel="stylesheet" type="text/css" href="../global.css" />
 	<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
 	<link rel="stylesheet" type="text/css" href="classes.css" />

	<script language="javascript" src="../scripts/ajaxLib.js"></script>
	<script language="javascript" src="../scripts/modules.js"></script>
	
	<script language="javascript">
	<!--
		function deleteconfirm(ID, sName) 
		{
			if(confirm('Do you wish to delete \'' + sName + '\'?')) 
			{
				window.location="dl_delete.asp?iDLid=" + ID;
			}
		}

		function openWin2(url, name) 
		{
			popupWin = window.open(url, name,"resizable,width=380,height=380");
		}

		function ChangeRosterDefault( iClassSeasonId )
		{
			// Fire off the roster default change code without any return handler
			doAjax('setrosterdefaultseason.asp', 'classseasonid=' + iClassSeasonId, '', 'get', '0');
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
				<h3><%=Session("sOrgName")%> Seasons</h3>
			<!--	<a href="../recreation/default.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>-->
			</p>
			<!--END: PAGE TITLE-->

			<div id="functionlinks">
				<a href="season_edit.asp?sid=0"><img src="../images/go.gif" align="absmiddle" border="0" />&nbsp;New Season</a>&nbsp;&nbsp;
				<a href="season_list.asp?seasonclosed=<%=iSeasonClosed%>" ><%=sShowText%></a> 
			</div>

			<!--BEGIN: CLASS LIST-->
				<% ListSeasons sWhere %> 
			<!--END: CLASS LIST-->

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
' SUB LISTLOCATIONS()
'--------------------------------------------------------------------------------------------------
Sub ListSeasons( ByVal sWhere )
	Dim sSql, oSeason, iRowCount

	iRowCount = 0
	' GET ALL LOCATIONS FOR ORG
	sSql = "SELECT C.classseasonid, C.seasonyear, C.seasonname, C.isclosed, C.showpublic, S.season, C.isrosterdefault "
	sSql = sSql & " FROM egov_class_seasons C, egov_seasons S "
	sSql = sSql & " WHERE C.seasonid = S.seasonid and orgid = " & SESSION("ORGID") & sWhere & " Order by C.seasonyear desc, S.displayorder desc, C.seasonname"

	Set oSeason = Server.CreateObject("ADODB.Recordset")
	oSeason.Open sSQL, Application("DSN"), 0, 1

	' DRAW LINK TO NEW Season

	If NOT oSeason.EOF Then

		' DRAW TABLE 
		response.write vbcrlf & "<div class=""seasonshadow""><table cellpadding=""0"" cellspacing=""0"" border=""0"" id=""seasonlist"">"
		
		' HEADER ROW
		response.write vbcrlf & "<tr><th class=""namecell"">Name</th><th>Year</th><th>Season</th><th>Show Public</th><th>Roster<br />Default</th></tr>"
		
		' LOOP THRU AND DISPLAY ROWS
		Do While Not oSeason.EOF
			iRowCount = iRowCount + 1
			If iRowCount Mod 2 = 0 Then
				sClass = " class=""altrow"" "
			Else
				sClass = ""
			End If 

			response.write vbcrlf & "<tr " & sClass & " id=""" & iRowCount & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
			'response.write "<td nowrap=""nowrap"" class=""namecell""><a href=""season_edit.asp?sid=" & oSeason("classseasonid") & """>" & Trim(oSeason("seasonname")) & "</a></td>"
			response.write "<td nowrap=""nowrap"" class=""namecell"" title=""click to edit"" onClick=""location.href='season_edit.asp?sid=" & oSeason("classseasonid") & "';"">" & Trim(oSeason("seasonname")) & "</td>"
			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='season_edit.asp?sid=" & oSeason("classseasonid") & "';"">" & oSeason("seasonyear") & "</td>"
			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='season_edit.asp?sid=" & oSeason("classseasonid") & "';"">" & oSeason("season") & "</td>"

			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='season_edit.asp?sid=" & oSeason("classseasonid") & "';"">"
			If oSeason("showpublic") Then 
				response.write "Yes"
			Else
				'response.write "No"
				response.write "&nbsp;"
			End If 
			response.write "</td>"

			response.write "<td align=""center"">"
			response.write "<input type=""radio"" name=""rosterdefault"" value=""" & oSeason("classseasonid") & """"
			If oSeason("isrosterdefault") = True Then
				response.write " checked=""checked"" "
			End If 
			response.write " onClick='ChangeRosterDefault(" & oSeason("classseasonid") & ");' />"
			response.write "</td>"

			response.write "</tr>"
			oSeason.MoveNext
		Loop 

		' ClOSE TABLE AND FREE OBJECTS
		response.write vbcrlf & "</table></div>"
	
	Else
		' NO distribution lists WERE FOUND
		response.write "<font color=""red""><strong>No Seasons currenty exist.</strong></font>"
	
	End If
	oSeason.close
	Set oSeason = Nothing 

End Sub


%>


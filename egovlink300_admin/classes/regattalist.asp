<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: regattalist.asp
' AUTHOR: Steve Loar
' CREATED: 02/24/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0	2/24/2009	Steve Loar	-	Initial version
' 1.1	4/7/2010	Steve Loar - Modified to remove things related to adding regatta team members
' 1.2	5/14/2010	Steve Loar - Split captain name into first and last
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iStatusid, iCategoryid, iClasstypeid, iDatefilter, sStartdate, sEnddate, sDefaultRange, iClassSeasonId
Dim sSearchName, sSearchActivity, iTeamGroupId, iOrderById, iClassId, bFilter, sSearchMember

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "regatta registration", sLevel	' In common.asp

If request("classseasonid") = "" Then 
   iClassSeasonId = GetRosterSeasonId()
   bFilter = False 
Else
   iClassSeasonId = CLng(request("classseasonid"))
   bFilter = True
End If 

sSearchName = request("searchname")
sSearchMember = request("searchmember")

If request("regattateamgroupid") <> "" Then 
	iTeamGroupId = CLng(request("regattateamgroupid"))
Else
	iTeamGroupId = 0
End If 

If request("orderbyid") <> "" Then 
	iOrderById = CLng(request("orderbyid"))
Else
	iOrderById = 0
End If 

iClassId = 0

%>
<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />


	<script language="Javascript" src="tablesort.js"></script>
	<script language="Javascript" src="../scripts/modules.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>

	<script language="Javascript">
	<!--
		function deleteconfirm(ID, sName) 
		{
			if(confirm('Do you wish to delete ' + sName + '?')) 
			{
				window.location="class_delete.asp?classid=" + ID;
			}
		}

		function gotosinglesignup( id )
		{
			window.location="regattasinglesignup.asp?classid=" + id;
		}

		function doCalendar(sField) 
		{
			var w = (screen.width - 350)/2;
			var h = (screen.height - 350)/2;
			eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=regattalist", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function Validate()
		{
			document.RegattaList.submit();
		}

		function ViewCart()
		{
			location.href='class_cart.asp';
		}

		function ExportTeams()
		{
			// build a string of selected teams
			var sTeamPicks = '';
			for (var t = 0; t <= parseInt($("teamcount").value); t++)
			{
				// See if a row exists for this one
				if ($("teamid" + t))
				{
					// If it is marked for export, then add it
					if ($("teamid" + t).checked == true)
					{
						if (sTeamPicks != '')
						{
							sTeamPicks += ',' ;
						}
						sTeamPicks += $("teamid" + t).value;
					}
				}
			}
			if (sTeamPicks != '')
			{
				sTeamPicks = '(' + sTeamPicks + ')';
				location.href='regattateamexport.asp?teampicks=' + sTeamPicks;
			}
			else
			{
				alert("Please select some teams for the export, first.");
			}
		}

		function SelectAllTeams()
		{
			for (var t = 0; t <= parseInt($("teamcount").value); t++)
			{
				// See if a row exists for this one
				if ($("teamid" + t))
				{
					// check it
					$("teamid" + t).checked = true;
				}
			}
		}

		function DelselectAllTeams()
		{
			for (var t = 0; t <= parseInt($("teamcount").value); t++)
			{
				// See if a row exists for this one
				if ($("teamid" + t))
				{
					// uncheck it
					$("teamid" + t).checked = false;
				}
			}
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

<% If CartHasItems() Then %>
	<div id="topbuttons">
		<input type="button" name="viewcart" class="button" value="View Cart" onclick="ViewCart();" />
	</div>
<%	End If %>

	<!--BEGIN: PAGE TITLE-->
	<p>
		<font size="+1"><strong>River Regatta Team Rosters and Registration</strong></font><br />
	</p>
	<!--END: PAGE TITLE-->

	<!--BEGIN: FILTER SELECTION-->
	<div class="filterselection">
	 	<fieldset class="filterselection">
			<legend class="filterselection">Search Options</legend>
			<p>
				<form name="RegattaList" method="post" action="regattalist.asp">
					<table border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td>Season: </td>
							<td colspan="2">
								<% ShowSeasonPicks iClassSeasonId  %>
							</td>
						</tr>
						<tr>
							<td>Team Groups:</td>
							<td colspan="2">
								<%	ShowTeamGroups iTeamGroupId		%>
							</td>
						</tr>
						<tr>
							<td>Team/Captain Like:</td>
							<td colspan="2"><input type="text" name="searchname" value="<%=sSearchName%>" size="100" maxlength="100" /></td>
						</tr>
						<tr>
							<td>Order By:</td>
							<td colspan="2">
								<%	ShowOrderByPicks iOrderById		%>
							</td>
						</tr>
						<tr>
							<td>&nbsp;</td>
							<td colspan="2"><input class="button" type="button" onclick="Validate()" value="Refresh Results" /></td>
						</tr>
					</table>
				</form>
			</p>
 		</fieldset>
	</div>
	<!--END: FILTER SELECTION-->

		<p>
<%		
'	<tr>
'		<td>Team Member Like:</td>
'		<td colspan="2"><input type="text" name="searchmember" value="=sSearchMember" size="100" maxlength="100" /></td>
'	</tr>

		If CLng(iClassSeasonId) > CLng(0) Then		
			iClassId = GetRegattaClassId( iClassSeasonId, "isteamsignup" )
%>
			
			<input type="button" class="button" value="Register a New Team" onclick="location.href='regattateamsignup.asp?classid=<%=iClassId%>';" /> &nbsp;
<%		End If		%>

	<!--BEGIN: CLASS LIST-->

	<% 
		DisplayRegattaTeams iClassSeasonId, sSearchName, iTeamGroupId, iOrderById, sSearchMember
	%>

	<!--END: CLASS LIST-->
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
' void DisplayRegattaTeams( iClassSeasonId, sSearchName, iTeamGroupId, iOrderById, sSearchMember )
'--------------------------------------------------------------------------------------------------
Sub DisplayRegattaTeams( ByVal iClassSeasonId, ByVal sSearchName, ByVal iTeamGroupId, ByVal iOrderById, ByVal sSearchMember )
	Dim sSql, sWhere, oClasslist, sFrom, iClassId, iTotalMembers, sSelect

	sWhere = ""
	sFrom = ""
	iTotalMembers = CLng(0)

	' Get the classid for the add to team event
	iClassId = GetRegattaClassId( iClassSeasonId, "isaddtoteam" )

	If CLng(iClassSeasonId) <> CLng(0) Then 
		sWhere = " AND C.classseasonid = " & iClassSeasonId
	End If 

	If sSearchName <> "" Then 
		sSearchName = dbsafe(LCase(sSearchName))
'		If sWhere = "" Then
'			sWhere = " WHERE "
'		Else
'			sWhere = sWhere & " AND "
'		End If 
		sWhere = sWhere & " AND (LOWER(T.regattateam) LIKE LOWER('%" & sSearchName & "%') OR LOWER(T.captainname) LIKE LOWER('%" & sSearchName & "%')) "
	End If 

	If CLng(iTeamGroupId) > CLng(0) Then
'		If sWhere = "" Then
'			sWhere = " WHERE "
'		Else
'			sWhere = sWhere & " AND "
'		End If 
		sWhere = sWhere & " AND T.regattateamgroupid = " & iTeamGroupId
	End If 

	If CLng(iOrderById) > CLng(0) Then
		sOrderBy = "G.displayorder, "
		sSelect = ", G.displayorder "
	Else
		sOrderBy = ""
		sSelect = ""
	End If 

'	sSql = "SELECT DISTINCT T.regattateamid, T.regattateam, T.captainname, ISNULL(G.regattateamgroup,'') AS regattateamgroup " & sSelect
'	sSql = sSql & " FROM egov_class C INNER JOIN egov_regattateams T ON C.classid = T.classid "
'	sSql = sSql & " INNER JOIN egov_regattateamgroups G ON T.orgid = G.orgid AND "
'	sSql = sSql & " T.regattateamgroupid = G.regattateamgroupid " & sFrom & sWhere
'	sSql = sSql & " ORDER BY " & sOrderBy & " T.regattateam"

	sSql = "SELECT DISTINCT T.regattateamid, T.regattateam, T.captainfirstname, T.captainlastname, ISNULL(G.regattateamgroup,'') AS regattateamgroup, A.paymentid " & sSelect
	sSql = sSql & " FROM egov_class C, egov_regattateams T, egov_regattateamgroups G, egov_class_list L, egov_accounts_ledger A "
	sSql = sSql & " WHERE C.classid = T.classid AND T.regattateamgroupid = G.regattateamgroupid " & sWhere
	sSql = sSql & " AND L.regattateamid = T.regattateamid AND C.classid = L.classid AND L.classlistid = A.itemid "
	sSql = sSql & " ORDER BY " & sOrderBy & " T.regattateam"

	'response.write sSql & "<br /><br />"
	'response.End 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then

		If CLng(iClassSeasonId) <> CLng(0) Then	
			' if the add to teams event has been created give them a button
'			If CLng(iClassId) > CLng(0) Then 
'				response.write vbcrlf & "<input type=""button"" class=""button"" value=""Add Members to a Team"" onclick=""location.href='regattamembersignup.asp?classid=" & iClassId & "';"" /> &nbsp;"
'			End If 
			response.write vbcrlf & "<input type=""button"" class=""button"" value=""Export Selected Teams to Excel"" onclick=""ExportTeams()"" />"
			response.write vbcrlf & "<br /><br /><input type=""button"" class=""button"" value=""Select All"" onclick=""SelectAllTeams()"" /> &nbsp;&nbsp; " 
			response.write vbcrlf & "<input type=""button"" class=""button"" value=""Deselect All"" onclick=""DelselectAllTeams()"" />"
			response.write vbcrlf & "</p>"
		End If		

		'DRAW TABLE WITH CLASSES LISTED
		response.write vbcrlf & "<div class=""shadow"">" 
		response.write vbcrlf & "<table id=""regattateamlist"" cellpadding=""5"" cellspacing=""0"" border=""0"">" 
		
		'HEADER ROW
		response.write vbcrlf & "<tr><th>Select</th><th>Team Name</th><th>Group</th><th>Team Captain</th><th>Receipt</th></tr>"

		iRowCount = 0
		
		' LOOP THRU AND DISPLAY The EVENTS
		Do While Not oRs.EOF
  			iRowCount = iRowCount + 1
		  	response.write vbcrlf & "<tr id=""" & iRowCount & """"
   			If iRowCount Mod 2 = 0 Then 
			    	response.write " class=""altrow"" "
   			End If 

			response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">" 

			response.write "<td align=""center""><input type=""checkbox"" name=""teamid" & iRowCount & """ id=""teamid" & iRowCount & """ value=""" & oRs("regattateamid") & """ />"
			response.write "</td>"
			
			' Team name
			response.write "<td align=""left"" onClick=""location.href='regattateamlist.asp?regattateamid=" & oRs("regattateamid") & "';"">" 
			response.write oRs("regattateam")
			response.write "</td>"

			response.write "<td align=""center"" onClick=""location.href='regattateamlist.asp?regattateamid=" & oRs("regattateamid") & "';"">" 
			response.write oRs("regattateamgroup")
			response.write "</td>"

			response.write "<td align=""center"" onClick=""location.href='regattateamlist.asp?regattateamid=" & oRs("regattateamid") & "';"">" 
			response.write oRs("captainfirstname") & " " & oRs("captainlastname")
			response.write "</td>"

'			response.write "<td align=""right"" onClick=""location.href='regattateamlist.asp?regattateamid=" & oRs("regattateamid") & "';"">" 
'			response.write oRs("membercount")
'			iTotalMembers = iTotalMembers + CLng(oRs("membercount"))
'			response.write "</td>"

			response.write "<td align=""center"">"
			response.write "<a href=""view_receipt.asp?iPaymentId=" & oRs("paymentid") & """>" & oRs("paymentid") & "</a>"
			response.write "</td>"

			response.write " </tr>"

  			oRs.MoveNext
		Loop 

		' Put in the total member count row at the end of the table
'		response.write vbcrlf & "<tr class=""teammembertotalrow"" >"
'		response.write "<td align=""right"" colspan=""4""><strong>Total Team Members</strong>"
		
'		response.write "<input type=""hidden"" id=""teampicks"" name=""teampicks"" value="""" />"
'		response.write "</td>"
'		response.write "<td align=""right""><strong>" & iTotalMembers & "</strong></td>"
'		response.write " </tr>"

		response.write vbcrlf & "</table>"
		response.write vbcrlf & "</div>" 

		response.write "<input type=""hidden"" id=""teamcount"" name=""teamcount"" value=""" & iRowCount & """ />"
		response.write "<input type=""hidden"" id=""teampicks"" name=""teampicks"" value="""" />"

	Else
		response.write vbcrlf & "</p>"
		' NO teams WERE FOUND
		response.write "<font color=""red""><b>No Regatta Teams could be found.</b></font>"
	End If

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' void  ShowSeasonFilterPicks( iClassSeasonId )
'--------------------------------------------------------------------------------------------------
Sub ShowSeasonFilterPicks( ByVal iClassSeasonId )
	Dim sSql, oRs

	sSql = "SELECT C.classseasonid, C.seasonname FROM egov_class_seasons C, egov_seasons S  "
	sSql = sSql & " WHERE C.isclosed = 0 AND C.seasonid = S.seasonid AND orgid = " & SESSION("orgid")
	sSql = sSql & " ORDER BY C.seasonyear DESC, S.displayorder DESC, C.seasonname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""classseasonid"">" 

		Do While NOT oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("classseasonid") & """ "  
			If CLng(iClassSeasonId) = CLng(oRs("classseasonid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oSeasons("seasonname") & "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If
	oRs.Close
	Set oRs = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' void  ShowTeamGroups( iTeamGroupId )
'--------------------------------------------------------------------------------------------------
Sub ShowTeamGroups( ByVal iTeamGroupId )
	Dim sSql, oRs

	sSql = "SELECT regattateamgroupid, regattateamgroup FROM egov_regattateamgroups "
	sSql = sSql & " WHERE orgid = " & SESSION("orgid")
	sSql = sSql & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If not oRs.EOF Then
		response.write vbcrlf & "<select name=""regattateamgroupid"">" 
		response.write vbcrlf & "<option value=""0"">All Groups</option>"
		Do While NOT oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("regattateamgroupid") & """ "  
			If CLng(iTeamGroupId) = CLng(oRs("regattateamgroupid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("regattateamgroup") & "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If
	oRs.Close
	Set oRs = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowOrderByPicks( iOrderById )
'--------------------------------------------------------------------------------------------------
Sub ShowOrderByPicks( ByVal iOrderById )

	response.write vbcrlf & "<select name=""orderbyid"">"
	response.write vbcrlf & "<option value=""0"""
	If CLng(iOrderById) = CLng(0) Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Team Name</option>"
	response.write vbcrlf & "<option value=""1"""
	If CLng(iOrderById) = CLng(1) Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Team Group</option>"
	response.write vbcrlf & "</select>"

End Sub 



%>

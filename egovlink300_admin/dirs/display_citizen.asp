<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<% 
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: display_citizen.asp
' AUTHOR: ??? John Stullenberger ???
' CREATED: 01/24/2002
' COPYRIGHT: Copyright 2002 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is a list of registered citizens in E-Gov
'
' MODIFICATION HISTORY
' 1.0   01/24/2002  ????? - INITIAL VERSION
' 3.1	03/28/2011	Steve Loar - Adding links to view rental reservations for a citizen
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sShowCitizens, sShowCitizens2, sShowLinkText, sView, bEditFamily

sLevel = "../" ' Override of value from common.asp
sView = ""

session("RedirectPage") = GetCurrentURL()
Session("RedirectLang") = "Return to Citizen List"
Session("RedirectSubPage") = ""
Session("RedirectSubLang") = ""


If Not UserHasPermission( Session("UserId"), "edit citizens" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

If request("v") = "" Then
	If OrgHasFeature("hasfamily") Then
		' Everyone including family members
		sShowCitizens = 1
		sView = ""
		bEditFamily = True 
	Else
		' Head of Households only
		sShowCitizens = 1
		'sView = " and useremail is not NULL "
		sView = " and headofhousehold = 1 "
		bEditFamily = False 
	End If 
ElseIf clng(request("v")) = clng(2) Then 
	' Head of Households 
	sShowCitizens = 2
'	sShowCitizens2 = 2
'	sShowLinkText = "Show Citizens Needing Residency Verification"
	'sView = " and useremail is not NULL "
	sView = " and headofhousehold = 1 and userlname IS NOT NULL "
	bEditFamily = False 
ElseIf clng(request("v")) = clng(3) Then 
	' Those needing residency verification
	sShowCitizens = 3
'	sShowCitizens2 = 1
'	sShowLinkText = "Show Registered Citizens"
	'sView = " and u.residencyverified = 0 and residenttype = 'R' and useremail is not NULL "
	sView = " and u.residencyverified = 0 and residenttype = 'R' and headofhousehold = 1 and userlname IS NOT NULL "
	bEditFamily = False 
ElseIf clng(request("v")) = clng(1) Then 
	' Everyone including family members
	sShowCitizens = 1
	sView = ""
	bEditFamily = True 
ElseIf clng(request("v")) = clng(4) Then 
	' Blocked citizens
	sShowCitizens = 4
	'sView = " and u.registrationblocked = 1 and useremail is not NULL "
	sView = " and u.registrationblocked = 1 and headofhousehold = 1 and userlname IS NOT NULL "
	bEditFamily = False 
ElseIf clng(request("v")) = clng(5) Then 
	' Subscription emails
	sShowCitizens = 5
	'sView = " and u.registrationblocked = 1 and useremail is not NULL "
	sView = " and userfname IS NULL and userlname IS NULL and headofhousehold = 1 "
	bEditFamily = True 
End If 

' Handle searches
If request("searchvalue") <> "" Then
	' Clean up to see if it is a phone number (all numeric)
	sNumber = Trim(request("searchvalue"))
	sNumber = Replace(sNumber,"-","")
	sNumber = Replace(sNumber,"(","")
	sNumber = Replace(sNumber,")","")
	sNumber = Replace(sNumber," ","")
	If IsNumeric(sNumber) Then  ' Phone numbers
		sView = sView & " and (userhomephone like '%" & sNumber & "%'"
		sView = sView & " or usercell like '%" & sNumber & "%'"
		sView = sView & " or userworkphone like '%" & sNumber & "%'"
		sView = sView & " or userfax like '%" & sNumber & "%')"
	Else   ' not all numeric
		sView = sView & " and (userlname like '%" & dbsafe(request("searchvalue")) & "%'"
		sView = sView & " or userfname like '%" & dbsafe(request("searchvalue")) & "%'"
		If request("includeemails") = "on" Then 
			sView = sView & " or useremail like '%" & dbsafe(request("searchvalue")) & "%'"
			sCheckEmails = " checked=""checked"" "
		Else
			sCheckEmails = ""
		End If 
		sView = sView & ")"
	End If 
End If 
If request("searchvalue2") <> "" Then
	' Clean up to see if it is a phone number (all numeric)
	sNumber = Trim(request("searchvalue2"))
	sNumber = Replace(sNumber,"-","")
	sNumber = Replace(sNumber,"(","")
	sNumber = Replace(sNumber,")","")
	sNumber = Replace(sNumber," ","")
	If IsNumeric(sNumber) Then  ' Phone numbers
		sView = sView & " and (userhomephone like '%" & sNumber & "%'"
		sView = sView & " or usercell like '%" & sNumber & "%'"
		sView = sView & " or userworkphone like '%" & sNumber & "%'"
		sView = sView & " or userfax like '%" & sNumber & "%')"
	Else   ' not all numeric
		sView = sView & " and (userlname like '%" & dbsafe(request("searchvalue2")) & "%'"
		sView = sView & " or userfname like '%" & dbsafe(request("searchvalue2")) & "%'"
		If request("includeemails") = "on" Then 
			sView = sView & " or useremail like '%" & dbsafe(request("searchvalue2")) & "%'"
			sCheckEmails = " checked=""checked"" "
		Else
			sCheckEmails = ""
		End If 
		sView = sView & ")"
	End If 
End If 

%>

<html lang="en">
<head>
	<meta charset="UTF-8">
	<title><%=langBSCommittees%></title>

	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="./reservationliststyles.css" />

	<script async src="../scripts/selectAll.js"></script>
	<script async src="../scripts/modules.js"></script>

	<script>
	<!--

		reloadpage = function( ) {
			//var iSelected = document.displayForm1.v.selectedIndex + 1;
			var iSelected = document.displayForm1.v[document.displayForm1.v.selectedIndex].value;
			var sSearch = document.displayForm1.searchvalue.value;
			var sSecondSearch = document.displayForm1.searchvalue2.value;
			//alert(iSelected);
			location.href='display_citizen.asp?v=' + iSelected + '&searchvalue=' + sSearch + '&searchvalue2=' + sSecondSearch ;
		};

		FamilyList = function( sUserId ) {
			location.href='family_list.asp?userid=' + sUserId;
		};
		
		openWin2 = function(url, name) {
			popupWin = window.open(url, name, "resizable,width=380,height=300");
		};
	
	//-->
	</script>


</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="menu">
    <tr>
      <td background="../images/back_main.jpg">

			<% ShowHeader sLevel %>
			<!--#Include file="../menu/menu.asp"--> 
      </td>
    </tr>
  </table>

<!-- #include file="dir_constants.asp"-->

<div id="content">
	<div id="centercontent">

<%  %>
  <table border="0" cellpadding="10" cellspacing="0" width="100%" class="start">
    <tr>
		<td valign="top">
<%
			Dim pagesize, totalpages,RA,totalrecords,groupname, thisname,currentpage,conn,rs,groupmode,strSQL,CName,AdditonURL
			Dim numstartid,numendid,i,deleteurl,EventOrNot,Str_Bgcolor,username,password,str_image,editurl,FullName

			'groupmode=1, display individual group
			'groupmode=2, display all member

			'pagesize = Session("PageSize")
			pagesize = GetUserPageSize( Session("UserId") ) ' Steve Loar 2/6/2007

			'totalpages=1

			thisname = request.servervariables("script_name")

			If Not IsEmpty(request("currentpage")) Then 
				CurrentPage = CLng(request("currentpage"))
			Else 
				CurrentPage = 1
			End If 

			' the above value will be provided by the display_committee.asp
			'page size, RA, pagerecord, currentpage values must be declared to global variables.
			 DisplayRecords CurrentPage 
%>
		</td>
	</tr>
 </table>

 </div>
 </div>

<!--#Include file="../admin_footer.asp"-->  

<!--#include file='footer.asp'-->


<%

'--------------------------------------------------------------------------------------------------
' void DISPLAYRECORDS
'--------------------------------------------------------------------------------------------------
Sub  DisplayRecords( ByVal iCurrentPage )
	Dim pagesize, sSql, sCSVSql, oRs

	thisname = request.servervariables("script_name")

	If Trim(request("groupid")) <> "" Then 
		'response.write "<br>display group<br>"
		GroupMode = 1
		sSql = "SELECT u.userid, userfname, userlname, useremail, userpassword, userlname + ' ' + userfname as username,"
		sSql = sSql & " ug.groupname, u.registrationblocked, u.residenttype, birthdate, residencyverified, isnull(accountbalance, 00000.0000) as accountbalance, familyid, u.headofhousehold FROM egov_users u INNER JOIN vwcitizengroups ug"
		sSql = sSql & " ON u.userid=ug.citizenid WHERE u.orgid = " & Session("OrgID") & sView 
		sSql = sSql & " AND isdeleted = 0 AND userregistered = '1' AND ug.groupid = " & CLng(request("groupid")) & " ORDER BY userlname, userfname, useremail"
		' Build QUERY FOR EXPORT TO CSV
		sCSVSql = "SELECT userlname + ' ' + userfname as username, familyid, useremail, "
		sCSVSql = sCSVSql & " ug.groupname, u.registrationblocked, u.residenttype, birthdate, residencyverified, isnull(accountbalance, 00000.0000) as accountbalance, u.headofhousehold FROM egov_users u INNER JOIN  vwcitizengroups ug "
		sCSVSql = sCSVSql & " ON u.userid = ug.citizenid WHERE orgid = " & session("orgid") & sView
		sCSVSql = sCSVSql & " AND isdeleted = 0 AND userregistered = '1' AND ug.groupid = " & CLng(request("groupid")) & " ORDER BY userlname, userfname, useremail"
	Else 
		GroupMode = 2
		'response.write "<br>display All<br>"
		sSql = "SELECT userid, userfname,userlname, useremail, userpassword, userlname + ' ' + userfname as username, "
		sSql = sSql & " u.registrationblocked, u.residenttype, birthdate, residencyverified, isnull(accountbalance, 00000.0000) as accountbalance, familyid, u.headofhousehold FROM egov_users u WHERE u.orgid = " & Session("OrgID") & sView 
		sSql = sSql & " AND isdeleted = 0 AND userregistered = '1' ORDER BY userlname, userfname, useremail"
		' Build QUERY FOR EXPORT TO CSV
		sCSVSql = "SELECT userfname as [first name], userlname as [last name], familyid, CONVERT(varchar,birthdate,101) AS birthdate, (DATEDIFF(dd, birthdate, getdate())/365) AS age, useremail AS email, u.registrationblocked, "
		sCSVSql = sCSVSql & " u.residenttype, residencyverified, userhomephone AS [home phone], usercell AS [cell phone], "
		sCSVSql = sCSVSql & " useraddress AS address, useraddress2 AS address2, userunit AS unit, "
		sCSVSql = sCSVSql & " usercity AS city, userstate AS state, userzip AS zip, userbusinessname AS [business name], "
		sCSVSql = sCSVSql & " userbusinessaddress as [business address], userworkphone AS [work phone], "
		sCSVSql = sCSVSql & " ISNULL(accountbalance, 00000.0000) AS accountbalance, u.headofhousehold "
		sCSVSql = sCSVSql & " FROM egov_users u WHERE orgid = " & session("orgid") & sView
		sCSVSql = sCSVSql & " AND isdeleted = 0 AND userregistered = '1' ORDER BY userlname, userfname, useremail"
	End If

'	response.write sCSVSql & "<br /><br />"
	' STORE QUERY FOR EXPORT TO CSV
	session("DISPLAYQUERY") = sCSVSql
	
	pagesize = GetUserPageSize( Session("UserId") ) 
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.PageSize = pagesize
	oRs.CacheSize = pagesize
	'set oRs.ActiveConnection = conn
	oRs.CursorLocation = 3 
	oRs.Open sSql, Application("DSN"), 3, 1

	If request("pagenum") <> "" Then 
		pagenum = request("pagenum")
		sPageNum = "&pagenum=" & pagenum
	Else 
		pagenum = 0
		sPageNum = ""
	End If 

	If (Len(iCurrentPage) = 0 Or clng(iCurrentPage) < 1) And Not oRs.EOF Then 
		oRs.AbsolutePage = 1
	ElseIf Not oRs.EOF Then 
		If clng(iCurrentPage) <= oRs.PageCount Then 
			oRs.AbsolutePage = iCurrentPage
		Else 
			oRs.AbsolutePage = 1
		End If 
	End If 

	'response.write "oRs.AbsolutePage = " & oRs.AbsolutePage & "<br />"

	Dim abspage, pagecnt
	abspage = oRs.AbsolutePage
	pagecnt = oRs.PageCount

	'------- the following code dealing with the recordcount=0-------
	If oRs.recordcount = 0 Then 
		oRs.Close
		If groupmode = 1 Then 
			sSql = "SELECT groupname FROM citizengroups WHERE groupid = " & CLng(request("groupid"))
			oRs.Open sSql
			CName = langCommittee & ":&nbsp;" & oRs("groupname")	
			groupname = oRs("groupname")	
			oRs.Close
			'Set oRs = Nothing 
		Else 
			CName = langTabCommittees & ":&nbsp;" & langDiaplyMember
		End If 
		CName = "Citizen Management"
		statistics CName

		navagatorbar 0, True, iCurrentPage

		response.write vbcrlf & "<table border=""0"" cellpadding=""5"" cellspacing=""0"" id=""citizenlist"">"
		response.write vbcrlf & "<tr><th align=left width='60%'>&nbsp;&nbsp;&nbsp;&nbsp;" & langUser & "</th><th  align='center'>&nbsp;" & langTypeEmail & "</th></tr>"
		response.write vbcrlf & "<tr><td align=left colspan=2>&nbsp;&nbsp;<FONT SIZE=1 COLOR=red><strong>" & langNoRecords & "</strong></font></td></tr>"
		response.write vbcrlf & "</table>"

		Exit Sub 
	End If 
	'--------------------------------------------------------------------

	oRs.movefirst
	totalrecords = oRs.RecordCount
	TotalPages = (totalrecords \ pagesize) + 1  '\means integer/integer
	If totalrecords Mod pagesize = 0 And TotalPages > 0 Then 
		TotalPages = TotalPages-1
	End If 
	If totalrecords <= pagesize Then 
		TotalPages = 1 
	End If 
	If TotalPages < 1 Then
		TotalPages = 1
	End If 
	If isNumeric(iCurrentPage) Then 
		If iCurrentPage < 1 Then
			iCurrentPage = 1
		End If 
		If iCurrentPage > TotalPages Then
			iCurrentPage = TotalPages
		End If 
	Else 
		iCurrentPage = 1
	End If 
	oRs.AbsolutePage = iCurrentPage

	numstartid	= (iCurrentPage-1) * PageSize
	numendid = IIf(numstartid + PageSize < totalrecords, numstartid+pagesize- 1, totalrecords - 1)
	'response.write "<br>totalrecords=" & totalrecords & "  numberstartid=" & numstartid & " numendid=" & numendid
	'response.write "<br>currentpage=" & iCurrentPage & "  totalpages=" & TotalPages & "<br /><br /><br />"
	'==========================================================================================
	'RA=oRs.GetRows
	If groupmode = 1 Then 
		'CName=langCommittee&":&nbsp;"&RA(6,i)
		CName = langCommittee & ":&nbsp;" & oRs("groupname")
		'groupname=RA(6,i)
		groupname = oRs("groupname")
	Else 
		CName = langTabCommittees & ":&nbsp;" & langDiaplyMember
	End If 

	CName = "Citizen Management"

	statistics CName 

	navagatorbar 1, True, iCurrentPage

	deleteurl = "delete_multiplecitizen.asp?previousURL=" & thisname & "&Extra=" & request.querystring

	response.write vbcrlf & "<form name=""DeleteMember"" method=""post"" action=""" & deleteurl & """>"

	'==========  the following will display the whole table=========================================
	response.write vbcrlf & "<table border=""0"" cellpadding=""5"" cellspacing=""0"" id=""citizenlist"">"
	response.write vbcrlf & "<tr><th align=""left"">"
	' This is the checkbox that selects all displayed users for deletion'
	response.write "<input class=""listCheck"" type=checkbox name=""chkSelectAll"" onClick=""selectAll('DeleteMember', this.checked, 'delete')"">"
	response.write "</th><th align=left width=""60%"">"&langUser&"</th>"

	If OrgHasFeature("hasfamily") Then
		response.write "<th>Family Id</th>"
		response.write "<th>Age</th>"
		response.write "<th>Resident</th>"
		If OrgHasFeature("registration blocking") Then 
			response.write "<th>Blocked</th>"
		End If 
	End If 
	response.write "<th align='center' colspan=2>&nbsp;" & langTypeEmail & "</th>"
	If OrgHasFeature("activities") Then
		response.write "<th>Recreation Activities</th>"
	End If
	'If OrgHasFeature("rentals") Then
	If UserHasPermission( session("userid"), "rentals" ) Then 
		response.write "<th>Rental<br />Reservations</th>"
	End If
	If OrgHasFeature("citizen accounts") Then
		response.write "<th>Account<br />Balance</th>"
	End If 
	If OrgHasFeature("groups") Then
		response.write "<th>Edit Groups</th>"
	End If 
	response.write "</tr>"

	iRowCount = 0
	'--------------------------------------------------------------------------
	For i = numstartid To numendid
		'-- alternateviely show different color-----
		EventOrNot=(i+2) Mod 2
		If EventOrNot = 0 Then 
			sRowClass = ""
		Else 
			sRowClass = " class=""altrow"" "
		End If 
		'-------------------------------------------
		'username = RA(4,i)
		username = oRs("username")
		'password = RA(5,i)
		password = oRs("userpassword")
		'if (isnull(username)) or (username="") then
		'If Not RA(12,i) Then 
		If Not oRs("headofhousehold") Then 
			' Family Member image
			str_image="<img src='../images/newcontact.gif'>"
		Else
			' Head of family image
			str_image="<img src='../images/newuser.gif'>"
		end if
		'-----------------------------------------
		'GroupNumber=entry(RA(0,i))
		If bEditFamily Then
			'If RA(3,i) = "" Or IsNull(RA(3,i)) Then 
			If Not oRs("headofhousehold") Then
				' Family Members
				'editurl = "manage_family_member.asp?u=" & RA(0,i) & "&iReturn=0"
				editurl = "manage_family_member.asp?u=" & oRs("userid") & "&iReturn=0"
			Else
				' Head of Household
				'editurl = "display_individual_citizen.asp?userid=" & RA(0,i)
				'editurl = "update_citizen.asp?userid=" & RA(0,i)
				editurl = "update_citizen.asp?userid=" & oRs("userid")
			End If 
		Else 
			'editurl = "display_individual_citizen.asp?userid=" & RA(0,i)
			'editurl = "update_citizen.asp?userid=" & RA(0,i)
			editurl = "update_citizen.asp?userid=" & oRs("userid")
		End If 

		If IsNull(oRs("userlname")) Then
			' Handle subscriptions
			FullName = oRs("useremail")
		Else 
			FullName = Trim(oRs("userlname")) & ",&nbsp;" & Trim(oRs("userfname"))
		End If 
		iRowCount = iRowCount + 1
		response.write 	"<tr id=""" & iRowCount & """" & sRowClass & " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
	'	if HasPermission("CanEditUser")  and request.querystring("groupid")="" Then
		If request("groupid") = "" Then 
			response.write  "<td><input type=""checkbox"" name=""delete"" value=""" & oRs("userid") & """ />"
		Else
			response.write "<td onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='" & editurl & "';"">&nbsp;"
		end if
		response.write "</td>"
		'response.write " <td><a href=""" & editurl & """>" & FullName & "</a></td>"
		response.write "<td onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='" & editurl & "';"">" & str_image & "&nbsp;" & FullName & "</td>"
		If OrgHasFeature("hasfamily") Then
			' Show FamilyId
			response.write "<td>"
			response.write "<a href=""javascript:FamilyList('" & oRs("familyid") & "');"">" & oRs("familyid") & "</a>"
			response.write "</td>"
			' ********************************************************************************************
			' Show Age
			response.write "<td onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='" & editurl & "';"">"
			If Trim(oRs("birthdate")) = "" Or IsNull(oRs("birthdate")) Then
				response.write "Adult"
			Else
				response.write GetCitizenAge( oRs("birthdate") )
			End If 
			response.write "</td>"
			' show residency
			response.write " <td onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='" & editurl & "';"" align=""center"">"
			If oRs("residenttype") = "R" Then 
				response.write "Yes" 
				If OrgHasFeature( "residency verification" ) And Not oRs("residencyverified") Then
					response.write "?"
				End If 
			Else
				response.write "&nbsp;"
			End If 
			response.write "</td>"

			If OrgHasFeature("registration blocking") Then
				' Show blocked status
				response.write " <td onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='" & editurl & "';"" align=""center"">"
				If oRs("registrationblocked") Then 
					response.write "Yes" 
				Else
					response.write "&nbsp;"
				End If 
				response.write "</td>"
			End If 

		End If 
		If oRs("useremail") <> "" Then 
			response.write "<td align='center'><a href=""mailto:" & oRs("useremail") & """><img src=""../images/newmail_small.gif"" border=""0"" /></a></td>"
			response.write "<td onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='" & editurl & "';"">&nbsp;" & oRs("useremail") & "</td>"
		Else
			sMailTo = GetFamilyEmail( oRs("userid") )
			If sMailTo <> "" Then
				response.write "<td align='center'><a href=""mailto:" & sMailTo & """><img src=""../images/newmail_small.gif"" border=""0""></a></td>"
				response.write "<td onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='" & editurl & "';"">&nbsp;" & sMailTo & "</td>"
			Else
				response.write "<td onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='" & editurl & "';"" align='center'>&nbsp;</td>"
				response.write "<td onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='" & editurl & "';"">&nbsp;</td>"
			End If 
			
		End If 

		If OrgHasFeature("activities") Then
			If CitizenHasRecreationActivities( oRs("userid") ) Then 
				If oRs("useremail") = "" Or IsNull(oRs("useremail")) Then
					response.write "<td align=""center""><a href=""activities_list.asp?u=" & oRs("userid") & "&v=1"">View</a>"
				Else
					response.write "<td align=""center""><a href=""activities_list.asp?u=" & oRs("userid") & "&v=2"">View</a>"
				End If 
			Else
				response.write "<td onMouseOver=""this.title='click to edit';"" onMouseOut=""this.title='';"" onClick=""location.href='" & editurl & "';"" align=""center"">&nbsp;"
			End If 
			response.write "</td>"
		End If 

		' Link to view rental reservations 
		If UserHasPermission( session("userid"), "rentals" ) Then 
			response.write "<td align=""center"">"
			If UserHasRentalReservations( oRs("userid") ) Then 
				response.write "<a href=""reservationslist.asp?u=" & oRs("userid") & """>View</a>"
			Else
				response.write "&nbsp;"
			End If 
			response.write "</td>"
		End If 

		' Citizen accounts
		If OrgHasFeature("citizen accounts") Then
			response.write "<td align=""right""><a href=""citizen_account_history.asp?u=" & oRs("userid") & """>" & FormatCurrency(oRs("accountbalance"), 2) & "</a></td>"
		End If

		' Citizen groups for Documents security
		'response.write    "<td align='center'>&nbsp;"&RA(3,i)&"</td>"
		If OrgHasFeature("groups") Then
			response.write "<td><i><a href=javascript:openWin2('ManageCitizenGroup.asp?userid=" & oRs("userid") & "','_blank')>" & langEdit & "</a></i></td>"
		End If 
		response.write "</tr>"

		oRs.MoveNext 
	Next   
	'----------------------------------------------------------------------------
	response.write vbcrlf & "</table>"

	response.write vbcrlf & "</form>"
	
	navagatorbar 1, False, iCurrentPage

	'=======  end of displaying the whole table================================================
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void statistics CName
'--------------------------------------------------------------------------------------------------
Sub statistics( ByVal CName )

	response.write vbcrlf & "<table cellpadding=""0"" cellspacing=""0"" border=""0"">"
	response.write vbcrlf & "<tr><td><font size='+1'><b>" & CName & "</b></font><br />"
	response.write "</td></tr></table><br /><br />"

End Sub 


'--------------------------------------------------------------------------------------------------
' void  navagatorbar record, bHeader
'--------------------------------------------------------------------------------------------------
Sub navagatorbar( ByVal record, ByVal bHeader, ByVal iCurrentPage )

	If groupmode = 1 Then 
		AdditonURL = "groupid=" & request.querystring("groupid") & "&"
	Else 
		AdditonURL = ""
	End If

	If bHeader Then 
		response.write vbcrlf & "<form name=""displayForm1"" method=""post"" action=""display_citizen.asp"">"  
	End If 

	response.write "<div style='font-size:10px; padding-bottom:5px;'>"
	If iCurrentPage > 1 Then 
		response.write "<a href='" & thisname & "?" & AdditonURL & "currentpage=" & (iCurrentPage-1) & "&v=" & sShowCitizens & "&searchvalue=" & request("searchvalue") & "&searchvalue2=" & request("searchvalue2") & "&includeemails=" & request("includeemails") & "'>"
		response.write "<img src='../images/arrow_back.gif' align='absmiddle' border=0>"&langPrev&" "&pagesize&"</a>"
	Else 
		response.write "<a href='" & thisname & "?" & AdditonURL & "currentpage=" & (iCurrentPage-1) & "&v=" & sShowCitizens & "&searchvalue=" & request("searchvalue") & "&searchvalue2=" & request("searchvalue2") & "&includeemails=" & request("includeemails") & "'>"
		response.write "<img src='../images/arrow_back.gif' align='absmiddle' border=0><font color=#999999>" & langPrev & " " & pagesize & "</font></a>"
	End If 

	'response.write "<br>currentpage=" & iCurrentPage & "  totalpages=" & TotalPages
	If iCurrentPage < totalpages Then 
		response.write "&nbsp;&nbsp;<a href='" & thisname & "?" & AdditonURL & "currentpage=" & (iCurrentPage+1) & "&v=" & sShowCitizens & "&searchvalue=" & request("searchvalue") & "&searchvalue2=" & request("searchvalue2") & "&includeemails=" & request("includeemails") & "'>" & langNext & " " & pagesize
	Else 
		response.write "&nbsp;&nbsp;<!a href='" & thisname & "?" & AdditonURL & "currentpage=" & (iCurrentPage+1) & "&v=" & sShowCitizens & "&searchvalue=" & request("searchvalue") & "&searchvalue2=" & request("searchvalue2") & "&includeemails=" & request("includeemails") & "'><font color=#999999>" & langNext & " " & pagesize & "</font>" 
	End If 
	response.write 	"<img src='../images/arrow_forward.gif' align='absmiddle' border=""0"" /></a>"

	'------- the following is the additional convenient links showing on the top of table---
	If request.querystring("groupid")<>"" Then 
		response.write "&nbsp;&nbsp;<img src='../images/newgroup.gif' width='16' height='16' align='absmiddle'>&nbsp;&nbsp;"
		response.write "<a href=javascript:openWin2('ManageCitizenGroupMember.asp?groupid=" & request("groupid") & "','_blank')>" & langEdit & " " & langmemberShip & "</a>"
	Else 
		If record > 0  Then 
			' This is the Delete button at the top of the page'
			deleteConfirm = langWanttoDeleteMember
			If OrgHasFeature("citizen accounts") Then
				deleteConfirm = deleteConfirm & "\n\nIf a user has an account balance, they will not be deleted."
			End If 
			response.write "&nbsp;&nbsp;<img src=""../images/small_delete.gif"" align=""absmiddle"">&nbsp;<a href=""javascript:document.DeleteMember.submit();"" onClick=""javascript: return confirm('" & deleteConfirm & "');"">" & langDelete & "</a>"
		End If 

		If bHeader Then 
			ShowDisplayPicks sShowCitizens 
		
			' Search stuff here
			if Session("OrgID") <> "60" then
				response.write vbcrlf & "<br /><br /><input type=""text"" name=""searchvalue"" value=""" & request("searchvalue") & """ id=""searchvalue"" maxlength=""30"" size=""30"" /> &nbsp; "
			else
				response.write vbcrlf & "<br /><br /><input type=""text"" name=""searchvalue"" value=""" & request("searchvalue") & """ id=""searchvalue"" maxlength=""30"" size=""30"" /> &nbsp; <input type=""text"" name=""searchvalue2"" value=""" & request("searchvalue2") & """ id=""searchvalue2"" maxlength=""30"" size=""30"" /> &nbsp; "
			end if
			response.write "<input type=""submit"" class=""button"" value=""Search List"" name=""searchbutton"" /> &nbsp; "
			response.write "<input type=""button"" class=""button"" name=""export"" value=""Export to CSV"" onClick=""location.href='../export/csv_export.asp';"" /> &nbsp; "
			
			If OrgHasFeature("aged accounts report") Then
				response.write "<input type=""button"" class=""button"" value=""Aged Account Report"" onClick=""location.href='agedaccountexport.asp';"" /> &nbsp; "
			End If 

			response.write "<br /><input type=""checkbox"" name=""includeemails""" & sCheckEmails & " />&nbsp;Include Emails in Search<br /><br />"
		End If 

	End If

	If bHeader Then 
		response.write "</form>"
	End If 

End Sub  


'--------------------------------------------------------------------------------------------------
' void ShowDisplayPicks iShowCitizens 
'--------------------------------------------------------------------------------------------------
Sub ShowDisplayPicks( ByVal iShowCitizens )

	response.write "&nbsp; &nbsp; Display: "
	response.write vbcrlf & "<select name=""v"" onChange='reloadpage( );'>"

	response.write vbcrlf & "<option value=""1"""
	If clng(iShowCitizens) = clng(1) Then 
		response.write " selected=""selected"" " 
	End If
	response.write ">Everyone</option>"

	If OrgHasFeature( "hasfamily" ) Then 
		response.write vbcrlf & "<option value=""2"""
		If clng(iShowCitizens) = clng(2) Then 
			response.write " selected=""selected"" " 
		End If
		response.write ">Head of Household</option>"
	End If 

	If OrgHasFeature("residency verification") Then 
		response.write vbcrlf & "<option value=""3"""
		If clng(iShowCitizens) = clng(3) Then 
			response.write " selected=""selected"" " 
		End If
		response.write ">Need Residency Verification</option>"
	End If 

	If OrgHasFeature("registration blocking") Then 
		response.write vbcrlf & "<option value=""4"""
		If clng(iShowCitizens) = clng(4) Then 
			response.write " selected=""selected"" " 
		End If
		response.write ">Blocked Households</option>"
	End If 

	If OrgHasFeature("subscriptions") Then 
		response.write vbcrlf & "<option value=""5"""
		If clng(iShowCitizens) = clng(5) Then 
			response.write " selected=""selected"" " 
		End If
		response.write ">Email Only Subscribers</option>"
	End If

	response.write vbcrlf & "</select>"

End Sub 


'--------------------------------------------------------------------------------------------------
' string IIF( BCHECK, STRUE, SFALSE )
'--------------------------------------------------------------------------------------------------
Function IIf( ByVal bCheck, ByVal sTrue, ByVal sFalse )
	
	If bCheck Then
		IIf = sTrue 
	Else 
		IIf = sFalse
	End If 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean UserHasRentalReservations( iUserid )
'--------------------------------------------------------------------------------------------------
Function UserHasRentalReservations( ByVal iUserid )
	Dim sSql, oRs

	sSql = "SELECT COUNT(R.reservationid) AS hits "
	sSql = sSql & "FROM egov_rentalreservations R, egov_rentalreservationtypes T "
	sSql = sSql & "WHERE R.reservationtypeid = T.reservationtypeid AND T.isreservation = 1 AND R.orgid = " & session("OrgId")
	sSql = sSql & " AND T.reservationtypeselector = 'public' AND R.rentaluserid = " & iUserid
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then
			UserHasRentalReservations = True 
		Else
			UserHasRentalReservations = False 
		End If 
	Else
		UserHasRentalReservations = False 
	End If
	
	oRs.Close 
	Set oRs = Nothing 

End Function 



%>




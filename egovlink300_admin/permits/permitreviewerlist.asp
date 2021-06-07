<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitreviewerlist.asp
' AUTHOR: Steve Loar
' CREATED: 08/04/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of permit reviews
'
' MODIFICATION HISTORY
' 1.0   08/04/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, sFrom, iPageSize, bInitialLoad, sPermitNo, iReviewerId, iStatusItemCount, sStatusItem
Dim sParcelIdNumber, sStreetNumber, sStreetName, iPermitId, sPermitLocation, iPermitCategoryId

ReDim aReviewStatuses(0)

sLevel = "../" ' Override of value from common.asp

PageDisplayCheck "edit permit reviews", sLevel	' In common.asp

sSearch = ""
bInitialLoad = False 
sFrom = ""

If request("pagesize") <> "" Then 
	iPageSize = CLng(request("pagesize"))
Else
	iPageSize = GetUserPageSize( Session("UserId") ) ' In common.asp
	bInitialLoad = True 
End If 

If request("permitid") <> "" Or request("permitreviewid") <> "" Then 
	' They are coming in from a link in an email, so fileter to that and skip the rest
	If request("permitid") <> "" Then
		iPermitId = CLng(request("permitid"))
		sSearch = sSearch & " AND P.permitid = " & iPermitId
	End If
	If request("permitreviewid") <> "" Then
		sSearch = sSearch & " AND R.permitreviewid = " & CLng(request("permitreviewid"))
		If request("permitid") = "" Then
			iPermitId = GetPermitIdByPermitReviewId( CLng(request("permitreviewid")) )
		End If 
	End If 
	sPermitNo = GetPermitNumber( iPermitId )
	bInitialLoad = False 
Else 
	' This is the normal page search stuff
	If request("permitno") <> "" Then 
		sPermitNo = Trim(request("permitno"))
		sSearch = sSearch & BuildPermitNoSearch( sPermitNo )
	End If 

	If request("revieweruserid") <> "" Then 
		iReviewerId = CLng(request("revieweruserid"))
		If iReviewerId > CLng(0) Then 
			sSearch = sSearch & " AND R.revieweruserid = " & iReviewerId
		End If 
	Else
		If bInitialLoad Then
			' If they are a reviewer, default to their pick
			iReviewerId = GetDefaultReviewerUserId( Session("UserId") )
			If iReviewerId > CLng(0) Then 
				sSearch = sSearch & " AND R.revieweruserid = " & iReviewerId
			End If
		Else 
			iReviewerId = CLng(0)
		End If 
	End If 

	If CLng(request("reviewstatusid").count) > CLng(0) Then 
		sSearch = sSearch & " AND ( "
		iStatusItemCount = CLng(0) 
		For Each sStatusItem In request("reviewstatusid") 
			If iStatusItemCount > UBound(aReviewStatuses) Then 
				Redim Preserve aReviewStatuses(iStatusItemCount)
			End If 
			aReviewStatuses(iStatusItemCount) = sStatusItem
			iStatusItemCount = iStatusItemCount + CLng(1)
			If iStatusItemCount > CLng(1) Then 
				sSearch = sSearch & " OR "
			End If 
			If CLng(sStatusItem) > CLng(0) Then
				sSearch = sSearch & " R.reviewstatusid = " & CLng(sStatusItem)
			End If 
		Next 
		sSearch = sSearch & " ) "
	Else
		' None selected, or first time page displays
		If bInitialLoad Then 
			sSearch = sSearch & " AND RS.isinitialstatus = 1 "
		Else 
			sSearch = sSearch & " AND R.reviewstatusid = 0 "
		End If 
		aReviewStatuses(0) = 0
	End If 

	If request("residentstreetnumber") <> "" Then 
		sStreetNumber = request("residentstreetnumber")
		sSearch = sSearch & "AND A.residentstreetnumber = '" & dbsafe(request("residentstreetnumber")) & "' "
	End If 
	If request("streetname") <> "" And request("streetname") <> "0000" Then 
		sStreetName = request("streetname")
		sSearch = sSearch & " AND (A.residentstreetname = '" & dbsafe(sStreetName) & "' "
		sSearch = sSearch & " OR A.residentstreetname + ' ' + A.streetsuffix = '" & dbsafe(sStreetName) & "' "
		sSearch = sSearch & " OR A.residentstreetprefix + ' ' + A.residentstreetname + ' ' + A.streetsuffix = '" & dbsafe(sStreetName) & "' "
		sSearch = sSearch & " OR A.residentstreetprefix + ' ' + A.residentstreetname + ' ' + A.streetsuffix + ' ' + A.streetdirection = '" & dbsafe(sStreetName) & "' )"
	End If 

	If request("parcelidnumber") <> "" Then 
		sParcelIdNumber = request("parcelidnumber")
		sSearch = sSearch & " AND A.parcelidnumber = '" & dbsafe(request("parcelidnumber")) & "' "
	End If 

	If request("permitlocation") <> "" Then
		sPermitLocation = request("permitlocation")
		sSearch = sSearch & " AND P.permitlocation LIKE '%" & dbsafe(request("permitlocation")) & "%' "
	End If 

	If request("permitcategoryid") <> "" Then
		iPermitCategoryId = CLng(request("permitcategoryid"))
		If CLng(iPermitCategoryId) > CLng(0) Then 
			sSearch = sSearch & " AND T.permitid = P.permitid AND T.permitcategoryid = " & iPermitCategoryId & " "
			sFrom = ", egov_permitpermittypes T"
		End If 
	End If 

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
	<script language="Javascript" src="../reporting/scripts/dates.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

	<script type="text/javascript" src="../scripts/fastinit.js"></script>
	<script language="Javascript" src="../scripts/tablesort2.js"></script>

	<script language="Javascript">
	<!--

		function doCalendar(ToFrom) {
		  w = (screen.width - 350)/2;
		  h = (screen.height - 350)/2;
		  eval('window.open("../recreation/gr_calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function GoToPage( iPageNum )
		{
			$("pagenum").value = iPageNum;
			document.frmReviewSearch.submit();
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
				<font size="+1"><strong>Permit Reviews</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<!--BEGIN: FILTER SELECTION-->
			<div class="filterselection">
				<fieldset class="filterselection">
				   <legend class="filterselection">Search Options</legend>
					<p>
						<form name="frmReviewSearch" method="post" action="permitreviewerlist.asp">
							<input type="hidden" id="pagenum" name="pagenum" value="1" />
							<input type="hidden" name="rq" value="1" />
							<table cellpadding="2" cellspacing="0" border="0">
								<tr>
									<td class="reviewerlistlabel">Category:</td><td><% ShowPermitCategoryPicks iPermitCategoryId %></td>
								</tr>
								<tr>
									<td class="reviewerlistlabel">Reviewer:</td><td><% ShowPermitReviewers iReviewerId %></td>
								</tr>
								<tr>
									<td class="reviewerlistlabel">Permit #:</td><td><input type="text" name="permitno" size="20" maxlength="20" value="<%=sPermitNo%>" /></td>
								</tr>
								<tr>
									<td class="reviewerlistlabel">Review Status:</td><td><% ShowReviewStatuses aReviewStatuses, bInitialLoad %></td>
								</tr>
								<tr>
									<td class="reviewerlistlabel">Address:</td><td><%  DisplayLargeAddressList sStreetNumber, sStreetName %></td>
								</tr>
								<tr>
									<td class="reviewerlistlabel">Location Like:</td><td><input type="text" name="permitlocation" size="100" maxlength="100" value="<%=sPermitLocation%>" /></td>
								</tr>
								<tr>
									<td class="reviewerlistlabel">Parcel Id #:</td><td><input type="text" name="parcelidnumber" size="20" maxlength="20" value="<%=sParcelIdNumber%>" /></td>
								</tr>
								<tr>
									<td class="reviewerlistlabel">Records per Page:</td><td><input type="text" name="pagesize" size="10" maxlength="10" value="<%=iPageSize%>" /></td>
								</tr>
								<tr>
			    					<td colspan="2"><input class="button ui-button ui-widget ui-corner-all" type="submit" value="Refresh Results" /></td>
  								</tr>
							</table>
						</form>
					</p>
				</fieldset>
			</div>
			<!--END: FILTER SELECTION-->

<%				ShowReviews sSearch, sFrom, iPageSize %>			
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
' void ShowReviews sSearch, iSearchItem, sYearPick 
'--------------------------------------------------------------------------------------------------
Sub ShowReviews( ByVal sSearch, ByVal sFrom, ByVal iPageSize )
	Dim sSql, oRs, iRowCount

	iRowCount = 0

	sSql = "SELECT P.permitid, R.permitreviewid, R.permitreviewtype, RS.reviewstatus, ISNULL(R.revieweruserid,0) AS revieweruserid, "
	sSql = sSql & "ISNULL(A.residentstreetprefix,'') AS residentstreetprefix, releaseddate, "
	sSql = sSql & "A.residentstreetnumber, ISNULL(A.residentunit,'') AS residentunit, A.residentstreetname, "
	sSql = sSql & "ISNULL(A.streetsuffix,'') AS streetsuffix, ISNULL(A.streetdirection,'') AS streetdirection, "
	sSql = sSql & "ISNULL(permitlocation,'') AS permitlocation, L.locationtype "
	sSql = sSql & "FROM egov_permits P, egov_permitreviews R, egov_permitstatuses S, egov_reviewstatuses RS, "
	sSql = sSql & "egov_permitaddress A, egov_permitlocationrequirements L" & sFrom
	sSql = sSql & " WHERE releaseddate IS NOT NULL AND P.permitid = R.permitid AND P.permitstatusid = S.permitstatusid "
	sSql = sSql & "AND P.isonhold = 0 and P.isvoided = 0 AND iscompletedstatus = 0 AND R.reviewstatusid = RS.reviewstatusid "
	sSql = sSql & "AND P.permitlocationrequirementid = L.permitlocationrequirementid AND A.permitid = P.permitid "
	sSql = sSql & "AND P.orgid = " & session("orgid") & sSearch
	sSql = sSql & " ORDER BY RS.reviewstatusorder, P.permitid"

	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.PageSize = iPageSize
	oRs.CacheSize = iPageSize
	oRs.CursorLocation = 3
	oRs.Open sSQL, Application("DSN"), 3, 1

	If request("pagenum") <> "" Then 
		pagenum = CLng(request("pagenum"))
	Else 
		pagenum = CLng(0)
	End If 

	If (Len(pagenum) = 0 or CLng(pagenum) < CLng(1)) And Not oRs.EOF Then 
		oRs.AbsolutePage = 1
	ElseIf Not oRs.EOF Then 
		iPageCount = CLng(oRs.PageCount)
		If CLng(Request("pagenum")) <= CLng(oRs.PageCount) Then 
			oRs.AbsolutePage = Request("pageNum")
		Else 
			oRs.AbsolutePage = 1
		End If 
	End If 

	Dim abspage, pagecnt
	abspage = oRs.AbsolutePage
	pagecnt = oRs.PageCount
%>
     	<input type="button"<% if abspage <= 1 then response.write " disabled "%> name="prevRecordsButton" id="prevRecordsButton" value="<< Back" class="button ui-button ui-widget ui-corner-all" onclick="GoToPage(<%=abspage-1%>);"  />
       	<input type="button"<% if abspage >= pagecnt then response.write " disabled "%> name="nextRecordsButton" id="nextRecordsButton" value="Next >>" class="button ui-button ui-widget ui-corner-all" onclick="GoToPage(<%=abspage+1%>);"  />&nbsp;&nbsp;

		<table id="categorytypes" cellpadding="2" cellspacing="0" border="0" class="sortable">
			<tr><th>Permit #</th><th>Released</th><th>Address/Location</th><th>Review</th><th>Review<br />Status</th><th>Reviewer</th></tr>


<%
	For intRec = 1 To oRs.PageSize
		If Not oRs.EOF Then
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr id=""" & iRowCount & """"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
			response.write "<td title=""click to edit"" onClick=""location.href='permitrevieweredit.asp?permitreviewid=" & oRs("permitreviewid") & "';"">"
			response.write "&nbsp;" & GetPermitNumber( oRs("permitid") )
			response.write "</td>"
			
			' Released Date
			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='permitrevieweredit.asp?permitreviewid=" & oRs("permitreviewid") & "';"">" & oRs("releaseddate") & "</td>"

			response.write "<td title=""click to edit"" onClick=""location.href='permitrevieweredit.asp?permitreviewid=" & oRs("permitreviewid") & "';"">"

			Select Case oRs("locationtype")
				Case "address"
					response.write "&nbsp;" & oRs("residentstreetnumber")
					If oRs("residentstreetprefix") <> "" Then
						response.write " " & oRs("residentstreetprefix")
					End If 
					response.write " " & oRs("residentstreetname")
					If oRs("streetsuffix") <> "" Then
						response.write " " & oRs("streetsuffix")
					End If 
					If oRs("streetdirection") <> "" Then
						response.write " " & oRs("streetdirection")
					End If 
					response.write " " & oRs("residentunit")

				Case "location"
					response.write "&nbsp;" & Left(oRs("permitlocation"),25)
					If clng(Len(oRs("permitlocation"))) > clng(25) Then
						response.write "..."
					End If 

				Case Else
					response.write "&nbsp;"

			End Select  

			response.write "</td>"
			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='permitrevieweredit.asp?permitreviewid=" & oRs("permitreviewid") & "';"">" & oRs("permitreviewtype") & "</td>"
			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='permitrevieweredit.asp?permitreviewid=" & oRs("permitreviewid") & "';"">" & oRs("reviewstatus") & "</td>"
			response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='permitrevieweredit.asp?permitreviewid=" & oRs("permitreviewid") & "';"">"
			If CLng(oRs("revieweruserid")) > CLng(0) Then 
				response.write GetPermitReviewerName( CLng(oRs("revieweruserid")) )
			Else
				response.write "Unassigned"
			End If  
			response.write "</td>"
			response.write "</tr>"
			response.write "</tr>"
			oRs.MoveNext
		End If 
	Next 

	If sSearch <> "" Then 
		If CLng(iRowCount) = CLng(0) Then
			response.write vbcrlf & "<tr><td colspan=""6"">&nbsp;No Permits could be found that match your search criteria.</td></tr>"
		End If 
	Else 
		If CLng(iRowCount) = CLng(0) Then
			response.write vbcrlf & "<tr><td colspan=""6"">&nbsp;No Permits could be found.</td></tr>"
		End If 
	End If 
	%>
	</table>
     	<input type="button"<% if abspage <= 1 then response.write " disabled "%> name="prevRecordsButton" id="prevRecordsButton" value="<< Back" class="button ui-button ui-widget ui-corner-all" onclick="GoToPage(<%=abspage-1%>);"  />
       	<input type="button"<% if abspage >= pagecnt then response.write " disabled "%> name="nextRecordsButton" id="nextRecordsButton" value="Next >>" class="button ui-button ui-widget ui-corner-all" onclick="GoToPage(<%=abspage+1%>);"  />&nbsp;&nbsp;
	<%

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowPermitReviewers( iReviewerId )
'--------------------------------------------------------------------------------------------------
Sub ShowPermitReviewers( iReviewerId )
	Dim sSql, oRs

	sSql = "SELECT userid, firstname, lastname FROM users WHERE orgid = " & session("orgid") & " AND ispermitreviewer = 1 "
	sSql = sSQl & " ORDER BY lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select name=""revieweruserid"">"
	response.write vbcrlf & "<option value=""0"""
	If CLng(iReviewerId) = CLng(0) Then
		response.write " selected=""selected"" "
	End If 
	response.write ">All Reviewers</option>"

	Do While Not oRs.EOF
		response.write vbcrlf & "<option "
		If CLng(iReviewerId) = CLng(oRs("userid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write " value=""" & oRs("userid") & """>" & oRs("firstname") & " " & oRs("lastname") & "</option>"
		oRs.MoveNext
	Loop 
		
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowReviewStatuses( aReviewStatuses, bInitialLoad )
'--------------------------------------------------------------------------------------------------
Sub ShowReviewStatuses( ByRef aReviewStatuses, ByVal bInitialLoad )
	Dim sSql, oRs

	sSql = "SELECT reviewstatusid, reviewstatus, isinitialstatus FROM egov_reviewstatuses WHERE orgid = " & session("orgid")
	sSql = sSQl & " ORDER BY reviewstatusorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While NOT oRs.EOF
			response.write vbcrlf & "<input type=""checkbox"" name=""reviewstatusid"" value=""" & oRs("reviewstatusid") & """"
			If bInitialLoad Then
				If oRs("isinitialstatus") Then 
					response.write " checked=""checked"" "
				End If 
			Else
				For Each Item In aReviewStatuses
					If CLng(Item) = CLng(oRs("reviewstatusid")) Then
						response.write " checked=""checked"" "
					End If 
				Next 
			End If 
			response.write " />" & oRs("reviewstatus") & " &nbsp; "
			oRs.MoveNext
		Loop
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub DisplayLargeAddressList( sResidenttype, sStreetNumber, sStreetName, bFound )
'--------------------------------------------------------------------------------------------------
Sub DisplayLargeAddressList( ByVal sStreetNumber, ByVal sStreetName )
	Dim sSql, oAddressList, sCompareName

	sSQL = "SELECT DISTINCT sortstreetname, ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, "
	sSQL = sSQL & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSQL = sSQL & " FROM egov_residentaddresses "
	sSQL = sSQL & " WHERE orgid = " & session( "orgid" )
	sSQL = sSQL & " AND residentstreetname IS NOT NULL "
	sSQL = sSQL & " ORDER BY sortstreetname "
	
	Set oAddressList = Server.CreateObject("ADODB.Recordset")
	oAddressList.Open sSQL, Application("DSN"), 3, 1

	If NOT oAddressList.EOF Then 
		response.write "<input type=""text"" name=""residentstreetnumber"" value=""" & sStreetNumber & """ size=""8"" maxlength=""10"" /> &nbsp; "
		response.write "<select name=""streetname"">" 
		response.write vbcrlf & "<option value=""0000"">Choose street from dropdown...</option>"

		Do While Not oAddressList.EOF
			sCompareName = ""
			If oAddressList("residentstreetprefix") <> "" Then 
				sCompareName = oAddressList("residentstreetprefix") & " " 
			End If 

			sCompareName = sCompareName & oAddressList("residentstreetname")

			If oAddressList("streetsuffix") <> "" Then 
				sCompareName = sCompareName & " "  & oAddressList("streetsuffix")
			End If 

			If oAddressList("streetdirection") <> "" Then 
				sCompareName = sCompareName & " "  & oAddressList("streetdirection")
			End If 

			response.write vbcrlf & "<option value=""" & sCompareName & """"

			If sStreetName = sCompareName Then 
				response.write " selected=""selected"" "
			End If 

			response.write " >"
			response.write sCompareName & "</option>" 
			oAddressList.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oAddressList.Close
	Set oAddressList = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetDefaultReviewerUserId( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetDefaultReviewerUserId( iUserId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(userid) AS hits FROM users WHERE orgid = " & session("orgid") & " AND ispermitreviewer = 1 AND userid = " & iUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then 
			' IF they are a reviewer, set the default reviewer to them
			GetDefaultReviewerUserId = iUserId
		Else
			GetDefaultReviewerUserId = CLng(0)
		End If 
	Else 
		GetDefaultReviewerUserId = CLng(0)
	End If 

	oRs.CLose
	Set oRs = Nothing 
End Function 




%>

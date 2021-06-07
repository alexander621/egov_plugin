<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitaddresstypelist.asp
' AUTHOR: Steve Loar
' CREATED: 02/04/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of permit address types
'
' MODIFICATION HISTORY
' 1.0   02/04/2008	Steve Loar - INITIAL VERSION
' 1.1	04/01/2008	Steve Loar - Structure changed. No landvalue, totalvalue, tax district; added streetdirection
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, sWhere, sSearchField

sLevel = "../" ' Override of value from common.asp

PageDisplayCheck "permit addresses", sLevel	' In common.asp

If request("searchtext") = "" Then
	sSearch = ""
Else
	sSearch = request("searchtext")
End If 

If request("searchfield") = "" Then
	sSearchField = 1
Else
	sSearchField = request("searchfield")
End If 

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
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

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Permit Addresses</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->
			<form name="frmPermitSearch" method="post" action="permitaddresstypelist.asp">
				<div>
					<input type="text" name="searchtext" value="<%=Replace(sSearch,"""","&quot;")%>" size="50" maxlength="150" /> &nbsp; &nbsp;
					<select name="searchfield">
						<option value="1" <%If clng(sSearchField) = clng(1) Then response.write " selected=""selected"" "%>>All Search Fields</option>
						<option value="2" <%If clng(sSearchField) = clng(2) Then response.write " selected=""selected"" "%>>Street Name</option>
						<option value="3" <%If clng(sSearchField) = clng(3) Then response.write " selected=""selected"" "%>>Street Number</option>
						<option value="4" <%If clng(sSearchField) = clng(4) Then response.write " selected=""selected"" "%>>Parcel Id</option>
						<option value="5" <%If clng(sSearchField) = clng(5) Then response.write " selected=""selected"" "%>>Listed Owner</option>
					</select> &nbsp; &nbsp;
					<input type="submit" class="button ui-button ui-widget ui-corner-all" value="Search" />
					&nbsp; &nbsp; <input type="button" name="new" class="button ui-button ui-widget ui-corner-all" value="New Address" onclick="location.href='permitaddresstypeedit.asp?permitaddresstypeid=0&searchtext=<%=sSearch%>&searchfield=<%=sSearchField%>';" />
				</div>
			</form>

			<%	
			ShowPermitAddressTypes session("orgid"), sSearch, sSearchField
			%>
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
' Sub ShowPermitAddressTypes( iOrgid, sSearch, sSearchField )
'--------------------------------------------------------------------------------------------------
Sub ShowPermitAddressTypes( iOrgid, sSearch, sSearchField )
	Dim sSql, oRs, iRowCount, iPermitAddressTypeid, iPageSize, iAbspage, iPagecnt, iPagenum, intRec

	iRowCount = 0
	iPermitAddressTypeid = CLng(0)

	iPageSize = GetUserPageSize( Session("UserId") )

'	sSql = "SELECT permitaddresstypeid, streetnumber, ISNULL(streetprefix,'') AS streetprefix, ISNULL(streetname,'') AS streetname, "
'	sSql = sSql & " ISNULL(streettype,'') AS streettype, ISNULL(pin,'') AS pin FROM egov_permitaddresstypes "
	sSql = "SELECT residentaddressid, residentstreetnumber, ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, "
	sSql = sSql & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection, residentunit, parcelidnumber, residentcity, residentstate "
	sSql = sSql & " FROM egov_residentaddresses WHERE orgid = "& iOrgid 
	If sSearch <> "" Then
		'sSql = sSql & " AND (streetname LIKE '%" & dbsafe(sSearch) & "%' OR streetprefix LIKE '%" & dbsafe(sSearch) & "%' OR streettype LIKE '%" & dbsafe(sSearch) & "%' OR streetnumber LIKE '%" & dbsafe(sSearch) & "%' ) "
		
		Select Case sSearchField
			Case 2
				sSQl = sSql & " AND (residentstreetname LIKE '%" & sSearch & "%')"
			Case 3
				sSQl = sSql & " AND (residentstreetnumber LIKE '%" & sSearch & "%')"
			Case 4
				sSQl = sSql & " AND (parcelidnumber LIKE '%" & sSearch & "%')"
			Case 5
				sSQl = sSql & " AND (listedowner LIKE '%" & sSearch & "%')"
			Case Else 
				' Covers case of 1 which is all search fields
				sSQl = sSql & " AND (residentstreetname LIKE '%" & sSearch & "%' OR residentstreetnumber LIKE '%" & sSearch & "%' OR parcelidnumber LIKE '%" & sSearch & "%' OR listedowner LIKE '%" & sSearch & "%')"
		End Select 
	End If 
	'sSql = sSql & " ORDER BY sortstreetname, 3, 2, 5, 6"
	sSql = sSql & " ORDER BY sortStreetName, Cast(residentstreetnumber AS INT), residentcity"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.PageSize = iPageSize
	oRs.CacheSize = iPageSize
	oRs.CursorLocation = 3
	oRs.Open sSQL, Application("DSN"), 3, 1

	If request("pagenum") <> "" Then 
		iPagenum = CLng(request("pagenum"))
		sPagenum = "&pagenum=" & iPagenum
	Else 
		iPagenum = CLng(0)
		sPagenum = ""
	End If 

	If CLng(iPagenum) < CLng(1) And Not oRs.EOF Then
		oRs.AbsolutePage = 1
	ElseIf Not oRs.EOF Then 
		If iPagenum <= oRs.PageCount Then 
			oRs.AbsolutePage = iPagenum
		Else 
			oRs.AbsolutePage = 1
		End If 
	End If 

	iAbspage = oRs.AbsolutePage
	iPagecnt = oRs.PageCount

'	response.write "<p>iAbspage = " & iAbspage & "<br />"
'	response.write "iPagecnt = " & iPagecnt & "</p>"

	If Not oRs.EOF Then
		' Display the Paging elements 
		response.write vbcrlf & "<div style='font-size:10px; padding-bottom:10px;'>"
		If iAbspage > 1 Then
			response.write vbcrlf & "<a href=""permitaddresstypelist.asp?pagenum=" & (iAbspage - 1) & "&searchtext=" & sSearch & "&searchfield=" & sSearchField & """>"
		End If 
		response.write "<img src=""../images/arrow_back.gif"" align=""absmiddle"" border=""0"" />"
		response.write "&nbsp;" & langPrev & "&nbsp;" & iPageSize
		If iAbspage > 1 Then
			response.write "</a>"
		End If 
		response.write "&nbsp;&nbsp;"
		If iAbspage < iPagecnt Then 
			response.write "<a href=""permitaddresstypelist.asp?pagenum=" & (iAbspage + 1) & "&searchtext=" & sSearch & "&searchfield=" & sSearchField & """>"
		End If 
		response.write langNext & "&nbsp;" & iPageSize & "<img src=""../images/arrow_forward.gif"" align=""absmiddle"" border=""0"" />"
		If iAbspage < iPagecnt Then 
			response.write "</a>"
		End If 
		response.write vbcrlf & "</div>"

		' Display the addresses
		response.write vbcrlf & "<div class=""shadow"">"
		response.write vbcrlf & "<table class=""tablelist"" id=""addresstypes"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
		response.write vbcrlf & "<tr><th>Number</th><th>Street</th><th>Unit/<br />Suite</th><th>City</th><th>State</th><th>Parcel<br />Id No</th><th>Related<br />Permits</th><th>New<br />Permit</th></tr>"
		'Do While Not oRs.EOF
		For intRec = 1 To oRs.PageSize
			If Not oRs.EOF Then 
				If iPermitAddressTypeid <> CLng(oRs("residentaddressid")) Then 
					If iPermitAddressTypeid > CLng(0) Then
						response.write "</tr>"
					End If 
					iPermitAddressTypeid = CLng(oRs("residentaddressid"))
					iRowCount = iRowCount + 1
					response.write vbcrlf & "<tr id=""" & iRowCount & """"
					If iRowCount Mod 2 = 0 Then
						response.write " class=""altrow"" "
					End If 
					response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
					response.write "<td title=""click to edit"" onClick=""location.href='permitaddresstypeedit.asp?permitaddresstypeid=" & oRs("residentaddressid") & "&searchtext=" & sSearch & "&searchfield=" & sSearchField & sPagenum & "';"">&nbsp;" 
					response.write oRs("residentstreetnumber")
					response.write "</td>"
					response.write "<td title=""click to edit"" onClick=""location.href='permitaddresstypeedit.asp?permitaddresstypeid=" & oRs("residentaddressid") & "&searchtext=" & sSearch & "&searchfield=" & sSearchField & sPagenum & "';"" align=""center"">&nbsp;" 
					If oRs("residentstreetprefix") <> "" Then 
						response.write oRs("residentstreetprefix") & " " 
					End If 
					response.write oRs("residentstreetname")
					If oRs("streetsuffix") <> "" Then
						response.write " " & oRs("streetsuffix")
					End If
					If oRs("streetdirection") <> "" Then
						response.write " " & oRs("streetdirection")
					End If
					response.write "</td>"
					response.write "<td title=""click to edit"" onClick=""location.href='permitaddresstypeedit.asp?permitaddresstypeid=" & oRs("residentaddressid") & "&searchtext=" & sSearch & "&searchfield=" & sSearchField & sPagenum & "';"" align=""center"">&nbsp;" 
					response.write oRs("residentunit")
					response.write "</td>"
					response.write "<td title=""click to edit"" onClick=""location.href='permitaddresstypeedit.asp?permitaddresstypeid=" & oRs("residentaddressid") & "&searchtext=" & sSearch & "&searchfield=" & sSearchField & sPagenum & "';"" align=""center"">&nbsp;" 
					response.write oRs("residentcity")
					response.write "</td>"
					response.write "<td title=""click to edit"" onClick=""location.href='permitaddresstypeedit.asp?permitaddresstypeid=" & oRs("residentaddressid") & "&searchtext=" & sSearch & "&searchfield=" & sSearchField & sPagenum & "';"" align=""center"">&nbsp;" 
					response.write oRs("residentstate")
					response.write "</td>"
					response.write "<td title=""click to edit"" onClick=""location.href='permitaddresstypeedit.asp?permitaddresstypeid=" & oRs("residentaddressid") & "&searchtext=" & sSearch & "&searchfield=" & sSearchField & sPagenum & "';"" align=""center"">" 
					response.write oRs("parcelidnumber")
					response.write "</td>"

					If AddressHasPermits( oRs("residentaddressid") ) Then
						response.write "<td title=""click to view"" align=""center"" onClick=""location.href='addresspermitlist.asp?permitaddresstypeid=" & oRs("residentaddressid") & "&searchtext=" & sSearch & "&searchfield=" & sSearchFieldv & "';"">&nbsp;"
						response.write "view permits"
					Else
						response.write "<td title=""click to edit"" onClick=""location.href='permitaddresstypeedit.asp?permitaddresstypeid=" & oRs("residentaddressid") & "';"">&nbsp;"
					End If 
					response.write "<td align=""center""><a title=""click to create a permit for this address"" href=""newpermit.asp?permitaddresstypeid=" & oRs("residentaddressid") & """>Create</a></td>"
					response.write "</td>"
				End If 
				oRs.MoveNext 
			End If 
'		Loop 
		Next 
		response.write vbcrlf & "</tr>"
		response.write vbcrlf & "</table>"
		response.write vbcrlf & "</div>"
	Else
		If sSearch <> "" Then
			response.write vbcrlf & "<p>&nbsp;No Permit Addresses could be found that match your search criteria.</p>"
		Else 
			response.write vbcrlf & "<p>&nbsp;No Permit Addresses could be found. Click on the New Address button to start entering data.</p>"
		End If 
	End If  
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Function AddressHasPermits( iPermitAddressTypeId )
'--------------------------------------------------------------------------------------------------
Function AddressHasPermits( iPermitAddressTypeId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(permitaddressid) AS hits FROM egov_permitaddress "
	sSql = sSql & " WHERE residentaddressid = " & iPermitAddressTypeId
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then 
			AddressHasPermits = True 
		Else
			AddressHasPermits = False 
		End If 
	Else
		AddressHasPermits = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


%>

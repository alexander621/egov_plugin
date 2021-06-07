<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalslist.asp
' AUTHOR: Steve Loar
' CREATED: 08/13/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of rentals. From here you can create or edit rentals
'
' MODIFICATION HISTORY
' 1.0   08/13/2009	Steve Loar - INITIAL VERSION
' 1.1	03/24/2011	Steve Loar - deactivated functions added
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, iSearchItem, iLocationid, iCategoryId, sFrom, sRentalName, iOrderBy, sOrderBy, sLoadMsg
Dim iSupervisorUserId

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "create edit rentals", sLevel	' In common.asp

sSearch = ""
sFrom = ""

If request("s") = "d" Then
	sLoadMsg = "displayScreenMsg('The Rental Was Deleted');"
End If 

If request("locationid") <> "" Then 
	iLocationid = CLng(request("locationid"))
	If iLocationid > CLng(0) Then 
		sSearch = sSearch & " AND R.locationid = " & iLocationid
	End If 
Else
	iLocationid = 0
End If 

If request("recreationcategoryid") <> "" Then 
	iCategoryId = CLng(request("recreationcategoryid"))
	If iCategoryId > CLng(0) Then 
		sSearch = sSearch & " AND R.rentalid = C.rentalid AND C.recreationcategoryid = " & iCategoryId & " "
		sFrom = sFrom & ", egov_rentals_to_categories C "
	End If 
Else
	iCategoryId = 0
End If 

If request("supervisoruserid") <> "" Then 
	iSupervisorUserId = CLng(request("supervisoruserid"))
	If iSupervisorUserId > CLng(0) Then 
		sSearch = sSearch & " AND R.supervisoruserid = " & iSupervisorUserId
	End If 
Else
	iSupervisorUserId = 0
End If 

If request("rentalname") <> "" Then
	sRentalName = request("rentalname")
	sSearch = sSearch & " AND R.rentalname like '%" & dbsafe(sRentalName) & "%' "
Else
	sRentalName = ""
End If 

If request("orderby") <> "" Then
	iOrderBy = clng(request("orderby"))
	If iOrderBy = clng(1) Then 
		sOrderBy = " L.name, R.rentalname"
	Else
		sOrderBy = " R.rentalname, L.name"
	End If 
Else
	iOrderBy = 1
	sOrderBy = " L.name, R.rentalname"
End If 


%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalsstyles.css" />

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

	<script type="text/javascript" src="../scripts/fastinit.js"></script>
	<script language="Javascript" src="../scripts/tablesort2.js"></script>
	<script language="javascript" src="../scripts/modules.js"></script>

	<script language="Javascript">
	<!--

		function RefreshResults()
		{
			document.frmRentalsSearch.submit();
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
				<font size="+1"><strong>Rentals</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->
			<table id="screenMsgtable"><tr><td>
				<span id="screenMsg"></span>
			</td></tr></table>

			<!--BEGIN: FILTER SELECTION-->
			<div class="filterselection">
				<fieldset class="filterselection">
				   <legend class="filterselection">Search Options</legend>
					<p>
						<form name="frmRentalsSearch" method="post" action="rentalslist.asp">
							<table cellpadding="2" cellspacing="0" border="0">
								<tr>
									<td>Location:</td><td><% ShowLocationPicks iLocationid, true	' In rentalsguifunctions.asp %></td>
								</tr>
								<tr>
									<td>Category:</td><td><% ShowCategoryPicks iCategoryId %></td>
								</tr>
								<tr>
									<td>Supervisor:</td><td><% ShowRentalSupervisors iSupervisorUserId, "All Supervisors" %></td>
								</tr>
								<tr>
									<td>Name Like:</td><td><input type="text" id="rentalname" name="rentalname" value="<%=sRentalName%>" size="90" maxlength="90" /></td>
								</tr>
								<tr>
									<td>Order By:</td><td><% ShowOrderByPicks iOrderBy %></td>
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

			<input type="button" class="button" id="newrentalbutton" name="newrentalbutton" value="Create A New Rental" onclick="location.href='rentaledit.asp?rentalid=0';" /> &nbsp;&nbsp; 
			<input type="button" class="button" id="copyrentalbutton" name="copyrentalbutton" value="Copy To A New Rental" onclick="location.href='rentalcopy.asp';" />

<%				
			ShowRentals sSearch, sFrom, sOrderBy
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
' void ShowRentals( sSearch, sFrom, sOrderBy )
'--------------------------------------------------------------------------------------------------
Sub ShowRentals( ByVal sSearch, ByVal sFrom, ByVal sOrderBy )
	Dim sSql, oRs, iRowCount

	iRowCount = 0

	sSql = "SELECT R.rentalid, R.rentalname, L.name AS locationname, isdeactivated FROM egov_rentals R, egov_class_location L" & sFrom
	sSql = sSql & " WHERE R.orgid = " & session("orgid") & " AND R.locationid = L.locationid " & sSearch
	sSql = sSql & " ORDER BY " & sOrderBy

	'response.write "<br />" & sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<div id=""rentalslistshadow"" class=""shadow"">"
	response.write vbcrlf & "<table id=""rentalslist"" cellpadding=""1"" cellspacing=""0"" border=""0"" class=""sortable"">"
	response.write vbcrlf & "<tr><th>Rental</th><th>Location</th></tr>"

	Do While Not oRs.EOF
		iRowCount = iRowCount + 1
		response.write vbcrlf & "<tr id=""" & iRowCount & """"
		If iRowCount Mod 2 = 0 Then
			response.write " class=""altrow"" "
		End If 
		response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
		response.write "<td class=""firstcol"" align=""left"" title=""click to edit"" onClick=""location.href='rentaledit.asp?rentalid=" & oRs("rentalid") & "';"" nowrap=""nowrap"">"
		response.write oRs("rentalname")
		If oRs("isdeactivated") Then
			response.write " (deactivated)"
		End If 
		response.write "</td>"
		response.write "<td align=""center"" title=""click to edit"" onClick=""location.href='rentaledit.asp?rentalid=" & oRs("rentalid") & "';"" nowrap=""nowrap"">"
		response.write oRs("locationname") & "</td>"
		response.write "</tr>"
		oRs.MoveNext
	Loop 

	If sSearch <> "" Then 
		If CLng(iRowCount) = CLng(0) Then
			response.write vbcrlf & "<tr><td colspan=""2"">&nbsp;No Rentals could be found that match your search criteria.</td></tr>"
		End If 
	Else 
		If CLng(iRowCount) = CLng(0) Then
			response.write vbcrlf & "<tr><td colspan=""2"">&nbsp;No Rentals could be found.</td></tr>"
		End If 
	End If 

	response.write vbcrlf & "</table></div>"

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' void ShowCategoryPicks( iCategoryId )
'--------------------------------------------------------------------------------------------------
Sub ShowCategoryPicks( ByVal iCategoryId )
	Dim sSql, oRs, iCount

	iCount = 0
	sSql = "SELECT recreationcategoryid, categorytitle FROM egov_recreation_categories "
	sSql = sSql & "WHERE orgid = " & session("orgid")
	sSql = sSql & " AND isforrentals = 1 ORDER BY categorytitle"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write "<select name=""recreationcategoryid"">"
	response.write vbcrlf & vbtab & "<option value=""0"">All Categories</option>"

	Do While Not oRs.EOF
		response.write vbcrlf & vbtab & "<option value=""" & oRs("recreationcategoryid") & """ "
		If clng(oRs("recreationcategoryid")) = clng(iCategoryId) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("categorytitle") & "</option>"
		oRs.MoveNext 
	Loop

	response.write vbcrlf & "</select>"
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowOrderByPicks( iOrderBy )
'--------------------------------------------------------------------------------------------------
Sub ShowOrderByPicks( iOrderBy )

	response.write "<select name=""orderby"">"
	response.write vbcrlf & "<option value=""1"""
	If iOrderBy = clng(1) Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">Location, Rental Name</option>"
	response.write vbcrlf & "<option value=""2"""
	If iOrderBy = clng(2) Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">Rental Name, Location</option>"
	response.write "</select>"

End Sub 




%>

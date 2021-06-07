<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalscategorieslist.asp
' AUTHOR: Steve Loar
' CREATED: 09/10/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of rentals categories. From here you can create or edit categories
'
' MODIFICATION HISTORY
' 1.0   09/10/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, iSearchItem, iMaxCategories, sLoadMsg

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "edit rentals categories", sLevel	' In common.asp

If request("s") = "d" Then
	sLoadMsg = "displayScreenMsg('The Rental Category Was Deleted');"
End If 

'iMaxCategories = GetMaxCategories()

%>

<html lang="en">
<head>
	<meta charset="UTF-8">
	
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="rentalsstyles.css" />

	<script src="../scripts/jquery-1.7.2.min.js"></script>

	<script src="../scripts/modules.js"></script>
	<script src="../scripts/ajaxLib.js"></script>


	<script>
	<!--

		function SetUpPage()
		{
			<%=sLoadMsg%>
		}

		function displayScreenMsg(iMsg) 
		{
			if(iMsg!="") 
			{
				$("#screenMsg").html( "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;" );
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("#screenMsg").html("");
		}

		
		function updateSequence( iCategoryId )
		{
			//alert($("sequence" + iCategoryId).value);
			doAjax('updatecategorysequence.asp', 'categoryid=' + iCategoryId + '&sequence=' + $("sequence" + iCategoryId).value, '', 'get', '0');
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
				<font size="+1"><strong>Rentals Categories</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<table id="screenMsgtable"><tr><td>
				<span id="screenMsg"></span>
				<input type="button" class="button" id="newcategorybutton" name="newcategorybutton" value="Create New Category" onclick="location.href='rentalcategoryedit.asp?rc=0';" />
			</td></tr></table>

<%				'Pull the list here
			ShowRentalsCategories iMaxCategories
%>			

		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%
'--------------------------------------------------------------------------------------------------
' ShowRentalsCategories iMaxCategories 
'--------------------------------------------------------------------------------------------------
Sub ShowRentalsCategories( ByVal iMaxCategories )
	Dim sSql, oRs, iRowCount

	iRowCount = 0

	sSql = "SELECT recreationcategoryid, categorytitle, sequenceid FROM egov_recreation_categories"
	sSql = sSql & " WHERE orgid = " & session("orgid") & " AND isforrentals = 1 ORDER BY categorytitle"

	'response.write "<br />" & sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<table id=""rentalslist"" cellpadding=""1"" cellspacing=""0"" border=""0"">"
	response.write vbcrlf & "<tr><th>Rentals Categories</th></tr>"

	Do While Not oRs.EOF
		iRowCount = iRowCount + 1
		response.write vbcrlf & "<tr id=""" & iRowCount & """"
		If iRowCount Mod 2 = 0 Then
			response.write " class=""altrow"" "
		End If 
		response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
		response.write "<td class=""firstcol"" align=""left"" title=""click to edit"" onclick=""location.href='rentalcategoryedit.asp?rc=" & oRs("recreationcategoryid") & "';"" nowrap=""nowrap"">"
		response.write oRs("categorytitle") & "</td>"

		response.write "</tr>"
		oRs.MoveNext
	Loop 
	response.write vbcrlf & "</table>"

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' GetMaxCategories
'--------------------------------------------------------------------------------------------------
Function GetMaxCategories()
	Dim sSql, oRs

	sSql = "SELECT ISNULL(COUNT(recreationcategoryid),0) AS hits FROM egov_recreation_categories "
	sSql = sSql & "WHERE orgid = " & session("orgid") & " AND isforrentals = 1"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetMaxCategories = clng(oRs("hits"))
	Else
		GetMaxCategories = clng(0)	' This will never happen if there are any categories
	End If 

	oRs.Close
	Set oRs = Nothing
	
End Function 


'--------------------------------------------------------------------------------------------------
' ShowSequencePicks iMaxCategories, iRecreationCategoryId, iSequenceid
'--------------------------------------------------------------------------------------------------
Sub ShowSequencePicks( iMaxCategories, iRecreationCategoryId, iSequenceid )
	Dim x

	response.write vbcrlf & "<select id=""sequence" & iRecreationCategoryId & """ name=""sequence" & iRecreationCategoryId & """ onchange=""updateSequence(" & iRecreationCategoryId & ");"">"

	For x = 1 To iMaxCategories
		response.write vbcrlf & "<option value=""" & x & """ "
		If clng(iSequenceid) = clng(x) Then
			response.write "selected=""selected"" "
		End If 
		response.write ">" & x & "</option>"
	Next 

	response.write vbcrlf & "</select>"
End Sub 



%>

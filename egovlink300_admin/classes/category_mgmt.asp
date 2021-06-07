<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CATEGORY_MGMT.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 03/21/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   04/17/06   JOHN STULLENBERGER - INITIAL VERSION
' 1.1   04/26/06   TERRY FOSTER - MADE FUNCTIONAL
' 1.2	10/11/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iMaxOrder, sScreenMsg

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "categories" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

iMaxOrder = GetMaxCategoryOrder()
'response.write "iMaxOrder = " & iMaxOrder
'response.end

If request("msg") = "3" Then
	' Deleted
	sScreenMsg = "Category Deleted."
Else
	sScreenMsg = ""
End If


%>
<html lang="en">
<head>

 	<title>E-Gov Administration Console</title>

	 <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

 	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
 	<link rel="stylesheet" href="../global.css" />
 	<link rel="stylesheet" href="classes.css" />
 	<link rel="stylesheet" href="../recreation/facility.css" />

	<script src="../scripts/jquery-1.7.2.min.js"></script>
	<script src="tablesort.js"></script>
	
	<script>
	<!--
		function displayScreenMsg(iMsg) 
		{
			if(iMsg!="") 
			{
				$("#screenMsg").html( "*** " + iMsg + " ***" );
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("#screenMsg").html( "" );
		}
	
		function deleteconfirm(ID) 
		{
			if(confirm('Are you sure you want to delete this category?')) 
			{
				window.location="category_delete.asp?iCategoryid=" + ID;
			}
		}

		function ChangeOrder(categoryid,sequenceid,iDirection)
		{
			location.href='category_reorder.asp?sequenceid='+ sequenceid + '&categoryid=' + categoryid + '&iDirection=' + iDirection;
		}
		
		<% If sScreenMsg <> "" Then %>
		$( document ).ready(function() {
			displayScreenMsg( "<%= sScreenMsg %>" );
		});
		<% End If %>

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
	<font size="+1"><strong>Recreation: Category Management</strong></font><br />
	<span id="screenMsg"></span><br />
</p>
<!--END: PAGE TITLE-->


<!--BEGIN: CLASS LIST-->
	<% ListCategorys %> 
<!--END: CLASS LIST-->

	</div>
</div>
<!--END: PAGE CONTENT-->


<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' SUB LISTCATEGORYS()
'--------------------------------------------------------------------------------------------------
Sub ListCategorys()
	Dim sSql, oRs

	' GET ALL categoryS FOR ORG
	sSql = "SELECT * FROM EGOV_CLASS_CATEGORIES WHERE ORGID = " & SESSION("ORGID") & " AND isroot <> 1 ORDER BY sequenceid"
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	' DRAW LINK TO NEW category
	response.write "<div id=""functionlinks"">"
	'response.write "<a href=""category_edit.asp?categoryid=0""><img src=""../images/go.gif"" align=""absmiddle"" border=""0"">&nbsp;New Category</a>"
	response.write "<input type=""button"" class=""button"" value=""New Category"" onclick=""location.href='category_edit.asp?categoryid=0';"" />"
	response.write "</div>"

	If Not oRs.EOF Then
		iRowCount = 1
		' DRAW TABLE 
		response.write "<table cellpadding=""5"" cellspacing=""0"" border=""0"" class=""tableadmin"" width=""100%"" id=""categorylist"">"
		
		' HEADER ROW
		response.write "<tr>"
		response.write "<th>Title</th><th>Image</th><th>Description</th><th colspan=""2"">&nbsp;</th>"
		response.write "</tr>"

		
		' LOOP THRU AND DISPLAY ROWS
		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
			response.write "<tr>"
			response.write "<td valign=""top""><strong>" & oRs("categorytitle") & "</strong></td>"
			If oRs("imgurl") <> "" Then 
				response.write "<td valign=""top""><img alt=""" & oRs("imgalttag")& """ src=""" & oRs("imgurl") & """></td>"
			Else
				response.write "<td>&nbsp;</td>"
			End If 
			response.write "<td valign=""top"">" & oRs("categorydescription") & "</td>"
			
			response.write "<td nowrap=""nowrap"">"
			'response.write "<a href=category_edit.asp?categoryid=" & oRs("categoryid") & ">Edit</a>"
			response.write "<input type=""button"" class=""button"" value=""Edit"" onclick=""location.href='category_edit.asp?categoryid=" & oRs("categoryid") & "';"" />"
			response.write "&nbsp;<input type=""button"" class=""button"" value=""Delete"" onclick=""javascript:deleteconfirm(" & oRs("categoryid") & ");"" />"
			'response.write " | <a href=""javascript:deleteconfirm(" & oRs("categoryid") & ")"">Delete</a> | "
			response.write "</td>"
			response.write "<td nowrap=""nowrap"" align=""center"">"
			If iRowCount <> 2 Then 
				'response.write vbcrlf & vbtab & "&nbsp;<a href=""javascript:ChangeOrder(" & oRs("categoryid") & "," & oRs("sequenceid") & ",-1);"">Move Up</a>"
				response.write "<input type=""button"" class=""button"
				If clng(oRs("sequenceid")) <> iMaxOrder Then
					response.write " moveupbtn"
				End If
				response.write """ value=""Move Up"" onclick=""javascript:ChangeOrder(" & oRs("categoryid") & "," & oRs("sequenceid") & ",-1);"" />"
			End If 
			If clng(oRs("sequenceid")) <> iMaxOrder And iRowCount <> 2 Then
				response.write "<br />"
			End If 
			If clng(oRs("sequenceid")) <> iMaxOrder Then
				'response.write vbcrlf & vbtab & "<a href=""javascript:ChangeOrder(" & oRs("categoryid") & "," & oRs("sequenceid") & ",1);"">Move Down</a>  </td>"
				response.write "&nbsp;<input type=""button"" class=""button"" value=""Move Down"" onclick=""javascript:ChangeOrder(" & oRs("categoryid") & "," & oRs("sequenceid") & ",1);"" />"
			End If 
			response.write "</td>"

			response.write "</tr>"

			oRs.MoveNext
		Loop 

		' ClOSE TABLE AND FREE OBJECTS
		response.write "</table>"
		oRs.close
		Set oRs = Nothing 
	
	Else
		' NO categoryS WERE FOUND
		response.write "<font color=red><b>There are no categories to display.</b></font>"
	
	End If

End Sub


'--------------------------------------------------------------------------------------------------
' Function GetMaxCategoryOrder()
'--------------------------------------------------------------------------------------------------
Function GetMaxCategoryOrder()
	Dim sSql, oRs

	sSql = "SELECT MAX(sequenceid) AS maxOrder FROM EGOV_CLASS_CATEGORIES WHERE OrgID = " & Session("orgid") 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If IsNull(oRs("MaxOrder")) Then
		GetMaxCategoryOrder = clng(0)
	Else
		GetMaxCategoryOrder = clng(oRs("MaxOrder"))
	End If 

	oRs.close
	Set oRs = Nothing

End Function 


%>



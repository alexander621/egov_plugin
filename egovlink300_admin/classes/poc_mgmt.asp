<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: POC_MGMT.ASP
' AUTHOR: Steve Loar
' CREATED: 05/10/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of Point Of Content
'
' MODIFICATION HISTORY
' 1.0   05/10/06	Steve Loar - INITIAL VERSION
' 1.1	10/11/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "poc" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

%>


<html>
<head>
 	<title>E-Gov Administration Console</title>

	 <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

 	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
 	<link rel="stylesheet" href="../global.css" />
 	<link rel="stylesheet" href="../recreation/facility.css" />
 	<link rel="stylesheet" href="classes.css" />

 	<script src="tablesort.js"></script>

 	<script>
	<!--

		function mouseOverRow( oRow )
		{
			oRow.style.backgroundColor='#93bee1';
			oRow.style.cursor='pointer';
			oNextRow = document.getElementById(eval(parseInt(oRow.id) + 1));
			if (oNextRow)
			{
				oNextRow.style.backgroundRepeat="repeat-x";
				oNextRow.style.backgroundImage="url(../images/shadow.png)";
			}
		}

		function mouseOutRow( oRow )
		{	
			oRow.style.backgroundColor='';
			oRow.style.cursor='';
			oNextRow = document.getElementById(eval(parseInt(oRow.id) + 1));
			if (oNextRow)
			{
				oNextRow.style.backgroundImage="";
			}
		}

		function deleteconfirm(ID, sName) 
		{
			if(confirm('Do you wish to delete \'' + sName + '\'?')) 
			{
				window.location="poc_delete.asp?pocid=" + ID;
			}
		}

	//-->
	</script>

</head>

<body>

 
<%'DrawTabs tabRecreation,1%>
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	
<!--BEGIN: PAGE TITLE-->
<p>
	<font size="+1"><strong>Recreation: Point Of Contact Management</strong></font><br />
	<!--<a href="../recreation/default.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>-->
</p>
<!--END: PAGE TITLE-->


<!--BEGIN: CLASS LIST-->
	<% ListPOCs %> 
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
' SUB ListPOCs()
'--------------------------------------------------------------------------------------------------
Sub ListPOCs()
	Dim sSql, iRowCount, sClass

	iRowCount = 0
	' GET ALL INSTRUCTORS FOR ORG
	sSQL = "SELECT * FROM EGOV_CLASS_pointofcontact WHERE orgid = " & SESSION("orgid") & " Order by name"

	Set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), 0, 1

	' DRAW LINK TO NEW INSTRUCTOR
	response.write "<div id=""functionlinks""><a href=""poc_edit.asp?pocid=0""><img src=""../images/go.gif"" align=""absmiddle"" border=""0"">&nbsp;New Point of Contact</a></div>"

	If NOT oList.EOF Then

		' DRAW TABLE 
		response.write "<div class=""shadow""><table cellpadding=""5"" cellspacing=""0"" border=""0"" class=""instructortable style-alternate sortable-onload-2"" >" 
		
		' HEADER ROW
		response.write "<tr>"
		response.write "<th>Name</th><th>Email</th><th>Phone</th><th>Delete</th>"
		response.write "</tr>"

		
		' LOOP THRU AND DISPLAY ROWS
		Do While Not oList.EOF
			iRowCount = iRowCount + 1
			If iRowCount Mod 2 = 0 Then
				sClass = " class=""altrow"" "
			Else
				sClass = ""
			End If 

			response.write "<tr " & sClass & " id=""" & iRowCount & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
			'response.write "<a href=""poc_edit.asp?pocid=" & oList("pocid") & """>Edit</a> | "
			
			response.write "<td title=""click to edit"" onClick=""location.href='poc_edit.asp?pocid=" & oList("pocid") & "';"">" & oList("name") & "</td>"
			response.write "<td title=""click to edit"" onClick=""location.href='poc_edit.asp?pocid=" & oList("pocid") & "';"">" & oList("email") & "</td>"
			response.write "<td title=""click to edit"" onClick=""location.href='poc_edit.asp?pocid=" & oList("pocid") & "';"">" & FormatPhone(oList("phone")) & "</td>"
			response.write "<td><a title=""click to delete"" href=""javascript:deleteconfirm(" & oList("pocid") & ", '" & FormatForJavaScript(oList("name")) & "')"">Delete</a></td></tr>"
			oList.MoveNext
		Loop 

		' ClOSE TABLE AND FREE OBJECTS
		response.write "</table></div>"
		oList.close
		Set oList = Nothing 
	
	Else
		' NO INSTRUCTORS WERE FOUND
		response.write "<font color=red><b>There are no point of contacts to display.</b></font>"
	
	End If

End Sub


%>



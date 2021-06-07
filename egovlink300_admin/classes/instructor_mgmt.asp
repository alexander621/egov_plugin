<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: INSTRUCTOR_MGMT.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 03/21/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   03/21/06   JOHN STULLENBERGER - INITIAL VERSION
' 1.1	10/11/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "instructors" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />

	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>
	<script language="Javascript" src="tablesort.js"></script>

	<script language="Javascript">
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
				// Fire off AJAX check of instructor. Do not delete if they are the instructor of any classes
				doAjax('check_instructor_for_deletion.asp', 'instructorid=' + ID, 'InstructorCheckReturn', 'get', '0');
				//window.location="instructor_delete.asp?instructorid=" + ID;
			}
		}

		function InstructorCheckReturn( sResult )
		{
			//alert( sResult );
			if (sResult != "KEEP")
			{
				//alert('Successful');
				window.location="instructor_delete.asp?instructorid=" + sResult;
			}
			else 
			{
				alert("This instructor cannot be deleted because they are still listed as teaching classes.\nYou must remove them from all classes before you can delete them.");
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

	
<!--BEGIN: PAGE TITLE-->
<p>
	<font size="+1"><strong>Recreation: Instructor Management</strong></font><br />
</p>
<!--END: PAGE TITLE-->


<!--BEGIN: CLASS LIST-->
	<% ListInstructors %> 
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
' SUB LISTINSTRUCTORS()
'--------------------------------------------------------------------------------------------------
Sub ListInstructors()
	Dim sSql, iRowCount, sClass

	iRowCount = 0
	' GET ALL INSTRUCTORS FOR ORG
	sSQL = "SELECT * FROM egov_class_instructor WHERE orgid = " & SESSION("ORGID") & " ORDER BY lastname, firstname"

	Set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), 0, 1

	session("DISPLAYQUERY") = "SELECT firstname, lastname, email, phone, cellphone FROM egov_class_instructor WHERE ORGID = " & SESSION("ORGID") & " Order by lastname, firstname"

	' DRAW LINK TO NEW INSTRUCTOR
	response.write "<div id=""functionlinks""><a href=""instructor_edit.asp?instructorid=0""><img src=""../images/go.gif"" align=""absmiddle"" border=""0"">&nbsp;New Instructor</a>"
	response.write "&nbsp;&nbsp;<input type=""button"" class=""button"" value=""Download to Excel"" onClick=""location.href='../export/excel_export.asp'"" />"
	response.write "</div>"

	If NOT oList.EOF Then

		' DRAW TABLE 
		response.write "<div class=""shadow""><table cellpadding=""5"" cellspacing=""0"" border=""0"" class=""instructortable style-alternate sortable-onload-2"" >" 
		
		' HEADER ROW
		response.write "<tr>"
		response.write "<th>Instructor Name</th><th>Email</th><th>Phone</th><th>Cell Phone</th><th>Delete</th>"
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
			'response.write "<a href=""instructor_edit.asp?instructorid=" & oList("instructorid") & """>Edit</a> | "
			response.write "<td title=""click to edit"" onClick=""location.href='instructor_edit.asp?instructorid=" & oList("instructorid") & "';"">" & oList("lastname") & ", " & oList("firstname") & "</td>"
			response.write "<td title=""click to edit"" onClick=""location.href='instructor_edit.asp?instructorid=" & oList("instructorid") & "';"">" & oList("email") & "</td>"
			response.write "<td title=""click to edit"" onClick=""location.href='instructor_edit.asp?instructorid=" & oList("instructorid") & "';"">" & FormatPhone(oList("phone")) & "</td>"
			response.write "<td title=""click to edit"" onClick=""location.href='instructor_edit.asp?instructorid=" & oList("instructorid") & "';"">" & FormatPhone(oList("cellphone")) & "</td>"
			response.write "<td><a title=""click to delete"" href=""javascript:deleteconfirm(" & oList("instructorid") & ", '" & FormatForJavaScript(oList("firstname") & " " & oList("lastname")) & "')"">Delete</a></td></tr>"
			oList.MoveNext
		Loop 

		' ClOSE TABLE AND FREE OBJECTS
		response.write "</table></div>"
	
	Else
		' NO INSTRUCTORS WERE FOUND
		response.write "<font color=""red""><b>There are no instructors to display.</b></font>"
	
	End If

	oList.close
	Set oList = Nothing 

End Sub
%>



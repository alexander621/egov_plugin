<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CLASS_WAIVERS.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 04/10/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   04/10/06   JOHN STULLENBERGER - INITIAL VERSION
' 1.1	10/11/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "waivers" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

%>


<html>
<head>
	<title>E-Gov Class Waivers</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />

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

		function deleteconfirm(ID,sWaiverName) 
		{
			var msg = "Deleting this waiver removes it from all currently assigned class and events. \n\n Do you wish to remove " + sWaiverName + "?";
			if(confirm(msg)) 
			{
				window.location="class_waiver_delete.asp?iWaiverid=" + ID;
			}
		}

		function ChangeCheck(field, iWaiverid, iFacilityId)
		{
			if (field.checked == true)
			{
			//			alert("checked");
				location.href='waiver_include.asp?iWaiverId='+ iWaiverid + '&iFacilityId=' + iFacilityId;
			}
			else
			{
			//			alert("unchecked");
				location.href='waiver_remove.asp?iWaiverId='+ iWaiverid + '&iFacilityId=' + iFacilityId;
			}
		}

		function EditWaiver(iWaiverId, iFacilityId)
		{
			location.href='waiver_edit.asp?iWaiverId=' + iWaiverId + '&iFacilityId=' + iFacilityId
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
	
	<p>
		<font size="+1"><strong>Recreation: Class Waivers</strong></font>
	</p>

	<div id="functionlinks">
		<a href="class_waiver_edit.asp?iWaiverId=0&iFacilityId=<%=iFacilityId%>" id="new_waiver"><img src="../images/go.gif" align="absmiddle" border="0">&nbsp;Add New Waiver</a>&nbsp;&nbsp;<br/><br/>
	</div>

	<div class="shadow">
		<!--class="tableadmin"-->
		<table cellpadding="5" cellspacing="0" border="0" class="instructortable">
			<tr>
				<th align="left" size="5" nowrap="nowrap">Waiver Type</th>
				<th nowrap="nowrap" align="left" size="100">Waiver Name</th>
				<th align="left">Delete</th>
			</tr>
				<% ListWaivers Session("OrgID") %>
		</table>
	</div>

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
Sub ListWaivers(iorgid)
	Dim sSql, iRowCount, sClass

	iRowCount = 0

	sSQL = "SELECT * FROM egov_class_waivers WHERE orgid = " & iorgid & " ORDER BY waivertype, waivername"

	Set oWaiver = Server.CreateObject("ADODB.Recordset")
	oWaiver.Open sSQL, Application("DSN"), 0, 1
	
	' LIST ALL WAIVER FOR ORGANIZATION
	If Not oWaiver.EOF Then
	
		Do While Not oWaiver.EOF
			iRowCount = iRowCount + 1
			If iRowCount Mod 2 = 0 Then
				sClass = " class=""altrow"" "
			Else
				sClass = ""
			End If 

			response.write "<tr " & sClass & " id=""" & iRowCount & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
			
			' CHECK BOX FOR ASSIGNMENT
			response.write "<td title=""click to edit"" onClick=""location.href='class_waiver_edit.asp?waiverid=" & oWaiver("waiverid") & "';"">&nbsp;" & oWaiver("waivertype") & "</td>"

			' WAIVER NAME
			response.write "<td nowrap=""nowrap"" title=""click to edit"" onClick=""location.href='class_waiver_edit.asp?waiverid=" & oWaiver("waiverid") & "';"">" & oWaiver("waivername") & "</td>"

			' WAIVER ACTIONS
			'response.write "<td><a href='class_waiver_edit.asp?waiverid=" & oWaiver("waiverid") & "';>Edit</a> | "
			response.write "<td><a title=""click to delete"" href=""javascript:deleteconfirm(" & oWaiver("waiverid") & ",'" & FormatForJavaScript(oWaiver("waivername")) & "');"">Delete</a></td>"
			response.write "</tr>"

			oWaiver.MoveNext
		
		Loop 
	
	End If 
	
	' CLOSE AND CLEAR OBJECTS
	oWaiver.close
	Set oWaiver = nothing

End Sub


%>



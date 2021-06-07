<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: REQUESTMERGEFORMS_MANAGE.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/21/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Manages PDFs files used to merge with request data.
'
' MODIFICATION HISTORY
' 1.0   02/21/07	JOHN STULLENBERGER - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
if isFeatureOffline("action line") = "Y" then
   response.redirect "../../admin/outage_feature_offline.asp"
end if

' INITIALIZE AND DECLARE VARIABLES
' SPECIFY FOLDER LEVEL
sLevel = "../../" ' OVERRIDE OF VALUE FROM COMMON.ASP


' USER SECURITY CHECK
If Not UserHasPermission( Session("UserId"), "requestmergeforms" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 
%>



<html>
<head>
  <title>E-Gov Link Manage PDF Forms</title>

  <link rel="stylesheet" type="text/css" href="../../global.css" />
  <link rel="stylesheet" type="text/css" href="../../menu/menu_scripts/menu.css" />

  <script language="Javascript" src="../../scripts/modules.js"></script>

  <script language="JavaScript">
	<!--

		function confirm_delete(ipdfid)
		{
			if (confirm("Are you sure you want to delete this pdf?"))
				{ 
					// DELETE HAS BEEN VERIFIED
					location.href='pdf_delete.asp?ipdfid=' + ipdfid;
				}
		}

	//-->
 </script>

</head>


<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">


<% ShowHeader sLevel %>


<!--#Include file="../../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
		
			<h3>Manage PDF Forms</h3>
			<%
				Call subDisplayPDFUploadForm()
			%>

			<P>&nbsp;</P>

			<%
				Call subListPDFs() 
			%>

	</div>
</div>
<!--END: PAGE CONTENT-->


<!--#Include file="../../admin_footer.asp"-->  


</body>
</html>



<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYPDFUPLOADFORM()
'--------------------------------------------------------------------------------------------------
Sub subDisplayPDFUploadForm()
	
	' BEGIN: UPLOAD FORM
	response.write "<form  name=""frmAddpdf"" action=""pdf_save.asp"" method=""POST"" enctype=""multipart/form-data"">"
	response.write "<div class=shadow>"
	response.write "<table class=tablelist cellspacing=0 cellpadding=2 >"
	response.write "<tr><th class=corrections colspan=2 align=left>&nbsp;Upload New PDF Form</th></tr>"
	response.write "<tr><td align=right>&nbsp;</td></tr>"
	response.write "<tr><td colspan=2><p><ol><li>Press <b>Browse</b> to find the file to upload. <li>Enter a description for the file (Max 1024 characters).  <li>Press <b>Save</b>.</ol>  &nbsp;Note: It may take a fews minutes to upload depending on the file size and your internet connection.</td></tr>"
	response.write "<tr><td align=right>&nbsp;</td></tr>"
	response.write "<tr><td align=right><b>Name: </b></td>"
	response.write "<td><input style=""width:650px;"" name=""filAttachment"" type=""file"" ></td>"
	response.write "</tr>"
	response.write "<tr>"
	response.write "<td align=right valign=top><b>Description: </b></td>"
	response.write "<TD><textarea style=""width:575px;height:50px;"" name=""pdfdesc""></textarea></td>"
	response.write "</tr>"
	response.write "<tr>"
	response.write "<TD colspan=2 align=right><input type=submit value=""Save"" style=""""></td>"
	response.write "</tr>"
	response.write "</table>"
	response.write "</div>"
	response.write "</form>"
	' END: UPLOAD FORM

End Sub


'------------------------------------------------------------------------------------------------------------
' SUB SUBLISTPDFS() 
'------------------------------------------------------------------------------------------------------------
Sub subListPDFs() 
	
	' BEGIN: PDF LIST
	response.write "<P><div class=shadow>"
	response.write "<table class=tablelist cellspacing=0 cellpadding=2 >"
	response.write "<tr><th class=corrections colspan=3 align=left>&nbsp;Manage Existing PDF Forms</th></tr>"
	
	response.write "<tr>"
	response.write "<td><b>Date Added</b></td>"
	response.write "<td><b>Name</b></td>"
	response.write "<td><b>Action</b></td>"
	response.write "</tr>"


	' GET PDF FORMS FOR THIS ORGANIZATION
	sSQL = "SELECT pdfid,pdf_name,pdf_description,date_added,orgid,adminuserid FROM egov_action_request_pdfforms WHERE orgid='" & session("orgid") & "' ORDER BY isdefault, pdf_name"
	Set oPDFList = Server.CreateObject("ADODB.Recordset")
	oPDFList.Open sSQL,Application("DSN"),1,3

	' IF THERE ARE PDF FORMS DISPLAY THEM	
	If NOT oPDFList.EOF Then

		' LIST ALL PDF FORMS FOUND
		Do While NOT oPDFList.EOF

			' HANDLE ALTERNATING ROW COLORS
			If sBGColor = "#FFFFFF" Then
				sBGColor = "#E0E0E0"
			Else
				sBGColor = "#FFFFFF"
			End If
			
			' WRITE PDF ROW WITH LINKS TO PDF INFORMATION AND ACTIONS
			Response.Write "<tr style=""background-color:" & sBGColor & ";"">"
			response.write "<td>" & oPDFList("date_added") & " </td>"
			response.write "<td>" & oPDFList("pdf_name") & " </td>"
			response.write "<td>"
			response.write "<a href=""pdf_view.asp?pdfid=" & oPDFList("pdfid") & """> View </a> | <a href=""javascript:confirm_delete('" & oPDFList("pdfid") & "');""> Delete </a> | <a href=""pdf_view.asp?pdfid=" & oPDFList("pdfid") & """> Make Default </a>"
			response.write "</td>"
			response.write "</tr>"

			' WRITE PDF ROW WITH PDF DESCRIPTION
			Response.Write "<tr style=""border-bottom:solid 1px #000000;background-color:" & sBGColor & ";"">"
			response.write "<td colspan=3><i>" & oPDFList("pdf_description") & "</i></td>"
			response.write "</tr>"
			
			' NEXT ROW
			oPDFList.MoveNext
		
		Loop

		' CLOSE AND DESTROY RECORDSET
		oPDFList.Close
		Set oPDFList = NOTHING

	Else
		
		' NO PDFS FOUND
		response.write "<tr><td colpan=3>"
		response.write "<span style=""color:red;font-weight:bold;padding-left: 10px;""> No Attachments added.</span>"
		response.write "</td></tr>"


	End If


	response.write "</table>"
	response.write "</div></P>"
	' END: PDF LIST


End Sub
%>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<% 
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: MANAGE_LETTER_TO_FORM.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 03/5/07
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  DELETE ATTACHMENT
'
' MODIFICATION HISTORY
' 1.0	03/05/07	JOHN STULLENBERGER - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
if isFeatureOffline("action line") = "Y" then
   response.redirect "../admin/outage_feature_offline.asp"
end if

' INITIALIZE AND DECLARE VARIABLES
Dim sError
sLevel = "../" ' OVERRIDE OF VALUE FROM COMMON.ASP

' SET TIMEZONE INFORMATION INTO SESSION
Session("iUserOffset") = request.cookies("tz")

' PROCESS VARIABLES
If request.servervariables("REQUEST_METHOD") = "POST" Then
	' SAVES CHANGES
	Call SubSaveValues()
End If
%>



<html>

<head>

  <title>E-Gov Link Manage Merge Fields</title>

  <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />

  <script language="Javascript" src="../scripts/modules.js"></script>

  <script language="Javascript" > 
  <!--

	//Set timezone in cookie to retrieve later
	var d=new Date();
	if (d.getTimezoneOffset)
	{
		var iMinutes = d.getTimezoneOffset();
		document.cookie = "tz=" + iMinutes;
	}

  //-->
  </script>

  <STYLE>
		div.correctionsbox {border: solid 1px #336699;padding: 4px 0px 0px 4px ;}
		div.correctionsboxnotfound  {background-color:#e0e0e0;border: solid 1px #000000;padding: 10px;color:red;font-weight:bold;}
		td.correctionslabel {font-weight:bold;}
		th.corrections {background-color:#93bee1;font-size:12px;padding:5px;color:#000000; }
		input.correctionstextbox {border: solid 1px #336699;width:400px;}
		.savemsg {font-size:12px;padding:5px;color:#0000ff;font-weight:bold; }
  </STYLE>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" >


<% ShowHeader sLevel %>


<!--#Include file="../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
		

		<h3>Manage Merge Fields</h3>		
		<img align=absmiddle src="../../admin/images/arrow_2back.gif"> <a  href="manage_letter.asp?iletterid=<%=request("iletterid")%>&iorgid=<%=session("orgid")%>">Return to Edit Form Letter</a> 
		
		<%
		' DISPLAY FORM TO MANAGE LETTER TO FORM RELATIONSHIPS
		Call subDisplayLetterstoForms(request("iletterid"))

		%>


	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>



<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYLETTERSTOFORMS(ILETTERID)
'--------------------------------------------------------------------------------------------------
Sub subDisplayLetterstoForms(iLetterId)


	' CHECK FOR EMPTY OR MISSING USERID
	If IsNull(iLetterId) or iLetterId="" then
		response.write "<P><div class=correctionsboxnotfound>No information available for this form letter.</div></P>"
	Else
			' GET INFORMATION FOR SPECIFIED FORM
			sSQL = "SELECT FLtitle,Case blnAllMergeFields When 1 Then ' CHECKED ' Else ' UNCHECKED ' End as blnAllMergeFields FROM FormLetters  WHERE FLid=" & iLetterId & " AND orgid=" & session("orgid")
			
			' OPEN RECORDSET
			Set oLetter = Server.CreateObject("ADODB.Recordset")
			oLetter.Open sSQL, Application("DSN"), 3, 1
			
			If NOT oLetter.EOF Then
		
				response.write "<form name=""frmManage"" action=""manage_letter_to_forms.asp"" method=""POST"" >"
				response.write "<div class=shadow>"
				response.write "<table class=tablelist cellspacing=0 cellpadding=2 >"
				response.write "<tr><th class=corrections colspan=2 align=left>&nbsp; Letter Name: " & oLetter("FLtitle") & "</th></tr>"

				' DISPLAY INSTRUCTIONS
				response.write "<tr><td colspan=2><P class=instructions style=""padding:5px;""><b>Instructions.</b><br>1. Check <b>Merge Fields found on any Form</b> to be able to insert Merge Fields in this letter that were created on any form.<br>2. Or, check one or more <b>Form Name(s)</b> to be able to insert Merge Fields that were only created on the selected forms.<br>3. <b>Save</b> when finished making changes.<br><br>The fields will show up on the <b>Edit Form Letter</b> screen.</p></td></tr>"
				
				' BUTTON ROW
				response.write "<tr><td colspan=2 class=correctionslabel align=""left""><input type=submit value=""Save"">&nbsp;&nbsp;<input  type=button value=""Cancel"" onClick=""location.href='list_letter.asp';""></td></tr>"

				' ALL OR NONE SELECTIONS
				response.write "<tr><td colspan=2>&nbsp;</td></tr>"
				response.write "<tr>"
				response.write "<td  align=""left"">&nbsp;<b>This letter can use Merge Fields defined by:</b>"
				response.write "<P>"
				response.write "<br>&nbsp;&nbsp;&nbsp;&nbsp;<input name=""checkall"" type=checkbox value=""Y""" &  oLetter("blnAllMergeFields") & "> Merge Fields found on any Form."
					' LIST FORMS
					Call subDisplayFormList()
				response.write "</p>"
				response.write "</td>"
				response.write "</tr>"
				response.write "<tr><td colspan=2>&nbsp;</td></tr>"


				' BUTTON ROW
				response.write "<tr><td colspan=2 class=correctionslabel align=""left""><input type=submit value=""Save"">&nbsp;&nbsp;<input  type=button value=""Cancel"" onClick=""location.href='list_letter.asp';""></td></tr>"


				response.write "</table>"
				response.write "</div>"
				response.write "<input type=hidden value=""" & request("iletterid") & """ name=""iletterid"">"
				response.write "</form>"

			Else
				' NO MATCHING USER FOUND
				response.write "<P><div class=correctionsboxnotfound>No information available for this form letter.</div></P>"
			End If
		End If

End Sub


'--------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYFORMLIST()
'--------------------------------------------------------------------------------------------------
Sub subDisplayFormList()
	
	sSQL = "SELECT action_form_id,action_form_name,action_form_internal FROM egov_action_request_forms  WHERE (action_form_type <> 2) AND orgid=" & session("orgid") & "  order by action_form_name"

	Set oFormList = Server.CreateObject("ADODB.Recordset")
	oFormList.Open sSQL, Application("DSN") , 3, 1
	
	' LIST ALL AVAIALABLE FORMS
	If NOT oFormList.EOF Then

		' LOOP THRU FORMS
		Do while NOT oFormList.EOF 
			response.write "<br>&nbsp;&nbsp;&nbsp;&nbsp;<input onClick=""document.frmManage.checkall.checked=false;"" name=""chk" & oFormList("action_form_id") & """ type=checkbox value=""" & oFormList("action_form_id")& """ " & fnIsChecked(oFormList("action_form_id"),request("iletterid")) & "> (" & oFormList("action_form_id") & ") " & oFormList("action_form_name")

			If oFormList("action_form_internal") Then
				response.write  " - <font style=""color:red;font-size:10px;"">Internal Only</font>"
				blnInternal = 0
			Else
				response.write " - <font style=""color:blue;font-size:10px;"">Public</font>"
				blnInternal = 1
			End If
			oFormList.MoveNext
		Loop

	Else
		
		' NO FORMS FORM ORGANIZATION

	End If

	Set oFormList = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' SUB SUBSAVEVALUES()
'--------------------------------------------------------------------------------------------------
Sub SubSaveValues()

	' CLEAR OLD VALUES
 	sSQL = "DELETE FROM egov_letter_to_form WHERE letterid='" & request("iletterid") & "'"
	Set oDelete = Server.CreateObject("ADODB.Recordset")
	oDelete.Open sSQL, Application("DSN"), 3, 1
	Set oDelete = Nothing


		' SET NEW VALUES

		' SET GLOBAL VALUE
		sSQL = "SELECT blnAllMergeFields FROM FormLetters  WHERE FLid='" & request("iletterid") & "' AND orgid='" & session("orgid") & "'"
		Set oLetter = Server.CreateObject("ADODB.Recordset")
		oLetter.Open sSQL, Application("DSN"), 3, 3
					
		If NOT oLetter.EOF Then
			
			If request.form("checkall") = "Y" Then 
				' TURN ON ALL
				oLetter("blnAllMergeFields") = 1
			Else
				' TURN OFF ALL
				oLetter("blnAllMergeFields") = 0
			End If
			
			' UPDATE AND CLOSE
			oLetter.Update
			oLetter.Close
			
		End If

		Set oLetter = Nothing

		

		' SET THE INDIVIDUAL FORM VALUES
		If request.form("checkall") <> "Y" Then 
			
			sSQL = "SELECT formid,letterid FROM egov_letter_to_form WHERE 1=2"
			Set oUpdate = Server.CreateObject("ADODB.Recordset")
			oUpdate.Open sSQL, Application("DSN"), 3, 3
			
			For Each oItem in Request.Form
				If Left(oItem,3) = "chk" Then
					' UPDATE LETTER TO FORM TABLE
					oUpdate.AddNew
					oUpdate("formid") = request(oItem)
					oUpdate("letterid") = request("iletterid")
					oUpdate.Update
				Else
					' IGNORE
				End If
			Next
			
			oUpdate.Close
			Set oUpdate = Nothing

		End If


End Sub


'--------------------------------------------------------------------------------------------------
' FUNCTION FNISCHECKED(IFORMID,ILETTERID) 
'--------------------------------------------------------------------------------------------------
Function fnIsChecked(iformid,iletterid) 

	sReturnValue = " UNCHECKED "
	
	sSQL = "SELECT formid,letterid FROM egov_letter_to_form WHERE formid='" & iformid & "' AND letterid='" & iletterid & "'"
	Set oChecked = Server.CreateObject("ADODB.Recordset")
	oChecked.Open sSQL, Application("DSN"), 3, 1
			
	If NOT oChecked.EOF Then
		sReturnValue = " CHECKED "
	End If

	Set oChecked = Nothing

	' RETURN VALUE
	fnIsChecked = sReturnValue

End Function
%>



<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->

<%
iLetterID = request("iLetterID")
iOrgID = session("orgId")
If request.servervariables("REQUEST_METHOD") = "POST" Then
	Select Case request("TASK")

		Case "new_question"
			' ADD NEW QUESTION TO FORM
			Call subAddQuestion(iorgid,iLetterID)
		Case Else
			' DEFAULT ACTION

	End Select
End If
%>

<HTML>
<HEAD>
<TITLE> E-GovLink Forms Management </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<link href="../global.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
function confirm_delete(ifieldid,sname)
{
	input_box=confirm("Are you sure you want to delete the question that begins (" + sname + ")? \nAll values for this question will be lost.");

	if (input_box==true)
		{ 
			// DELETE HAS BEEN VERIFIED
			location.href='delete_field.asp?iLetterID=<%=iLetterID%>&ifieldid='+ ifieldid;
		}
	else
		{
			// CANCEL DELETE PROCESS
		}
}
</script>
</HEAD>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%DrawTabs tabActionline,1%>

  <table border="0" cellpadding="6" cellspacing="0" class="start" width="100%">
    <!--<tr>
      <td><font size="+1"><b>Form Builder</b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.go(-1)"><%=langBackToStart%></a></td>
    </tr>-->
    <tr>
      <td valign="top">



<%
	' GET FORM GENERAL INFORMATION
	Dim sTitle
	Dim sIntroText
	Dim sFooterText
	Dim sMask
	subGetFormInformation(iLetterID)
	
%>



<div style="margin-top:20px; margin-left:20px;" >

<p><font class=label>Edit Form Letter </font> <small>[<a class=edit href="copy_form.asp?task=copyme&iLetterID=<%=iLetterID%>&iorgid=<%=session("orgid")%>">Copy This Form</a>]</small> <small>[<a class=edit href="../action_line/edit_form.asp?task=name&control=<%=iLetterID%>&iorgid=<%=session("orgid")%>">Manage This Form</a>]</small> <small>[<a class=edit href="list_forms.asp">Return to Form List</a>]</small><hr size="1" width="600px;" style="text-align:left; color:#000000;"></P>

<p><small>[<a class=edit href="edit_form.asp?task=name&iLetterID=<%=iLetterID%>&iorgid=<%=iorgid%>">Edit Name</a>]</small> - <font class=subtitle><%=sTitle%></font></p>



<div class=group>


<div class="orgadminboxf">


	<!--BEGIN: INTRO INFORMATION -->
		<P><small>[<a class=edit href="edit_form.asp?task=intro&iLetterID=<%=iLetterID%>&iorgid=<%=iorgid%>">Edit Intro</a>]</small></P>
		<P>
		<%If sIntroText <> "" Then
			response.write sIntroText
		Else
			response.write " - <i> Introduction text is currently blank </i> -"
		End If
		%></P>
	<!--END: INTRO INFORMATION -->


	
	
	<!--BEGIN: FORM FIELD INFORMATION -->
		<p> 
		<small>[<a class=edit href="add_field.asp?iLetterID=<%=iLetterID%>&iorgid=<%=iorgid%>" >Add New Question</a>]</small> </P>
		
		<P><% Call subDisplayQuestions(iLetterID) %> </P>

		<p><font color=red>*</font><B><i>Required Field</i></b></P>
	<!--END: FORM FIELD INFORMATION -->
	
	
	<!--BEGIN: ENDING NOTES -->
		<P><small>[<a class=edit href="edit_form.asp?task=footer&iLetterID=<%=iLetterID%>&iorgid=<%=iorgid%>">Edit Footer</a>]</small></P>  
		<P>
		<%If sFooterText <> "" Then
			response.write sFooterText
		Else
			response.write " - <i> Footer text is currently blank </i> -"
		End If
		%>
		</P>
	<!--END: ENDING NOTES -->

</div>


<!--include file="bottom_include.asp"-->

      </td>
       
    </tr>
  </table>
</BODY>
</HTML>



<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB SUBGETFORMINFORMATION(iLetterID)
'--------------------------------------------------------------------------------------------------
Sub subGetFormInformation(iLetterID)
	
	sSQL = "SELECT * FROM FormLetters WHERE FLid=" & iLetterID

	Set oForm = Server.CreateObject("ADODB.Recordset")
	oForm.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oForm.EOF Then
		' POPULATE DATA FROM RECORDSET
		sTitle = oForm("action_form_name")
		sIntroText = oForm("action_form_description")
		sFooterText = oForm("action_form_footer")
		sMask = oForm("action_form_contact_mask")

	End If

	Set oForm = Nothing 

End Sub



'------------------------------------------------------------------------------------------------------------
' FUNCTION DBSAFE( STRDB )
'------------------------------------------------------------------------------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function


'------------------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYQUESTIONS(iLetterID)
'------------------------------------------------------------------------------------------------------------
Sub subDisplayQuestions(iLetterID)

	sSQL = "SELECT * FROM egov_action_form_questions WHERE formid=" & iLetterID & " ORDER BY sequence"

	Set oQuestions = Server.CreateObject("ADODB.Recordset")
	oQuestions.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oQuestions.EOF Then
	
		response.write "<table>"
	
	Do While NOT oQuestions.EOF 
		
		' ENUMERATE QUESTIONS
		
		' DETERMINE IF REQUIRED
		sIsrequired = oQuestions("isrequired")
		If sIsrequired = True Then
			sIsrequired = " <font color=red>*</font> "
		Else
			sIsrequired = ""
		End If

		response.write "<tr><TD><small>[<a class=edit href=""edit_field.asp?iLetterID=" & iLetterID & "&iorgid=" & iorgid & " &ifieldid=" & oQuestions("questionid") & """>Edit</a>] [<a class=edit href=# onclick=""confirm_delete('" & oQuestions("questionid") & "','" & JSsafe(Left(oQuestions("prompt"),25)) & "...');"">Delete</a>] [<a class=edit href=""order_field.asp?direction=UP&iLetterID=" & iLetterID & "&iorgid=" & iorgid & " &ifieldid=" & oQuestions("questionid") & """>Move Up</a>] [<a class=edit href=""order_field.asp?direction=down&iLetterID=" & iLetterID & "&iorgid=" & iorgid & " &ifieldid=" & oQuestions("questionid") & """>Move Down</a>] [<a class=edit href=""order_field.asp?direction=top&iLetterID=" & iLetterID & "&iorgid=" & iorgid & " &ifieldid=" & oQuestions("questionid") & """>Move to Top</a>] [<a class=edit href=""order_field.asp?direction=bottom&iLetterID=" & iLetterID & "&iorgid=" & iorgid & " &ifieldid=" & oQuestions("questionid") & """>Move to Bottom</a>]</small></td></tr>"

		Select Case oQuestions("fieldtype")

			Case "2"
			' BUILD RADIO QUESTION
			response.write "<tr><td class=question>" & sIsrequired & oQuestions("prompt")& "</td></tr>"
			arrAnswers = split(oQuestions("answerlist"),chr(10))
			
			For alist = 0 to ubound(arrAnswers)
				response.write "<tr><td><input name=""question" & oQuestions("questionid") & """ class=formradio type=radio>" & arrAnswers(alist) & "</td></tr>"
			Next


			Case "4"
			' BUILD SELECT QUESTION
			response.write "<tr><td class=question>" & sIsrequired  & oQuestions("prompt")& "</td></tr>"
			arrAnswers = split(oQuestions("answerlist"),chr(10))
			
			response.write "<tr><td><select class=formselect>"
			For alist = 0 to ubound(arrAnswers)
				response.write "<option>" & arrAnswers(alist) & "</option>" 
			Next
			response.write "</select></td></tr>"


			Case "6"
			' BUILD CHECKBOX QUESTION
			response.write "<tr><td class=question>" & sIsrequired  & oQuestions("prompt")& "</td></tr>"
			arrAnswers = split(oQuestions("answerlist"),chr(10))
			
			For alist = 0 to ubound(arrAnswers)
				response.write "<tr><td><input class=formcheckbox type=checkbox>" & arrAnswers(alist) & "</td></tr>"
			Next



			Case "8"
			' BUILD TEXT QUESTION
			response.write "<tr><td class=question>" & sIsrequired  & oQuestions("prompt")& "</td></tr>"
			response.write "<tr><td><input value="""" type=""text"" style=""width:300px;"" maxlength=""100""></td></tr>"


			Case "10"
			' BUILD TEXTAREA QUESTION
			response.write "<tr><td class=question>" & sIsrequired  & oQuestions("prompt")& "</td></tr>"
			response.write "<tr><td><textarea class=formtextarea></textarea></td></tr>"
			

			Case Else

		End Select 

		response.write "<tr><TD>&nbsp;</td></tr>"

		oQuestions.MoveNext
	Loop

		response.write "</table>"
	End If

	Set oQuestions = Nothing 

End Sub

'------------------------------------------------------------------------------------------------------------
' FUNCTION ISREQUIRED(SMASK,IFIELD)
'------------------------------------------------------------------------------------------------------------
Function IsRequired(sMASK,iField)
	sValue = Mid(sMask,iField,1)
	
	If sValue = "2" Then
		sReturnValue = " <font color=red>*</font> "
	Else
		sReturnValue = ""
	End If

	IsRequired = sReturnValue
End Function

'------------------------------------------------------------------------------------------------------------
' FUNCTION ISDISPLAY(SMASK,IFIELD)
'------------------------------------------------------------------------------------------------------------
Function IsDisplay(sMASK,iField)
	sValue = Mid(sMask,iField,1)
	
	If sValue = "1" or sValue = "2" Then
		sReturnValue = True
	Else
		sReturnValue = False
	End If

	IsDisplay = sReturnValue
End Function

'------------------------------------------------------------------------------------------------------------
'Function JSsafe( strDB )
'------------------------------------------------------------------------------------------------------------
Function JSsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  strDB = Replace( strDB, "'", "\'" )
  strDB = Replace( strDB, chr(34), "\'" )
  strDB = Replace( strDB, ";", "\;" )
  strDB = Replace( strDB, "-", "\-" )
  JSsafe = strDB
End Function

%>
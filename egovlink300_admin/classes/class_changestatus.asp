<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_changestatus.asp
' AUTHOR: Steve Loar
' CREATED: 04/26/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   04/26/06   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	Dim iClassId, bIsParent, sCancelReason, iStatusId, sClassName

	iClassId = request("classid")
	iStatusId = clng(request("statusid"))
	sCancelReason = DBsafe( CStr(request("cancelreason")) )

	' Get isParent
	bIsParent = GetClassIsParent( iClassId )

	' update egov_class
	Change_Class_Status iClassId, iStatusId, sCancelReason 

	' send email to class for cancels
	If LCase(request("emailclass")) = "on" And iStatusId = 2 Then
		sClassName = GetClassName( iClassId )
		SendCancelMail iClassId, sCancelReason, sClassName
	End If 

	' update any children
	If bIsParent Then
		Change_Children_Status iClassId, iStatusId, sCancelReason
		If request("emailclass") = "on" And iStatusId = 2 Then
			' send email to children
			SendChildCancelMail iClassId, sCancelReason, sClassName
		End If 
	End If 

	response.redirect "edit_class.asp?classid=" & iClassId


'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Function GetClassInfo( iClassId, bIsParent )
'--------------------------------------------------------------------------------------------------
Function GetClassIsParent( iClassId )
	Dim sSql, oInfo

	sSql = "Select isparent from egov_class where classid = " & iClassId
	GetClassIsParent = False 

	Set oInfo = Server.CreateObject("ADODB.Recordset")
	oInfo.Open sSQL, Application("DSN"), 0, 1

	If Not oInfo.EOF Then
		GetClassIsParent = oInfo("isparent")
	End If 

	oInfo.close
	Set oInfo = Nothing
End Function  


'--------------------------------------------------------------------------------------------------
' Sub Cancel_Class( iClassId, sCancelReason )
'--------------------------------------------------------------------------------------------------
Sub Change_Class_Status( iClassId, iStatusId, sCancelReason )
	Dim oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")

		' Update the class table
		.CommandText = "Update egov_class Set statusid = " & iStatusId & ", cancelreason = '" & sCancelReason & "', canceldate = GetDate() Where classid = " & iClassId
		.Execute
	End With
	Set oCmd = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub Cancel_Children( iParentClassId, sCancelReason )
'--------------------------------------------------------------------------------------------------
Sub Change_Children_Status( iParentClassId, iStatusId, sCancelReason )
	Dim oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")

		' Update the class table
		.CommandText = "Update egov_class Set statusid = " & iStatusId & ", cancelreason = '" & sCancelReason & "', canceldate = Getdate() Where parentclassid = " & iParentClassId
		.Execute
	End With
	Set oCmd = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub SendCancelMail( iClassId, sCancelReason, sClassName )
'--------------------------------------------------------------------------------------------------
Sub SendCancelMail( iClassId, sCancelReason, sClassName )
	Dim sSql, oEmails

	sSql = "SELECT useremail FROM egov_class C, egov_class_list L, egov_users U "
	sSql = sSql & " WHERE useremail IS NOT NULL AND useremail != '' AND C.classid = L.classid AND L.userid = U.userid "
	sSql = sSql & " AND C.classid = " & iClassId 

	Set oEmails = Server.CreateObject("ADODB.Recordset")
	oEmails.Open sSQL, Application("DSN"), 0, 1

	Do While Not oEmails.EOF
		SendCancelNotification oEmails("useremail"), sCancelReason, sClassName
		oEmails.MoveNext
	Loop 

	oEmails.Close
	Set oEmails = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub SendChildCancelMail iParentClassId, sCancelReason, sClassName
'--------------------------------------------------------------------------------------------------
Sub SendChildCancelMail( iParentClassId, sCancelReason, sClassName )
	Dim sSql, oEmails

	sSql = "Select useremail from egov_class C, egov_class_list L, egov_users U where C.classid = L.classid and L.userid = U.userid "
	sSql = sSql & " and C.parentclassid = " & iParentClassId 

	Set oEmails = Server.CreateObject("ADODB.Recordset")
	oEmails.Open sSQL, Application("DSN"), 0, 1

	Do While Not oEmails.EOF
		SendCancelNotification oEmails("useremail"), sCancelReason, sClassName
		oEmails.movenext
	Loop 

	oEmails.close
	Set oEmails = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub SendCancelNotification( sToEmail, sCancelReason, sClassName )
'--------------------------------------------------------------------------------------------------
Sub SendCancelNotification( sToEmail, sCancelReason, sClassName )
	Dim sSubject, sFrom, sMsg2, objMail2, sOrgName, sBody, sCityUrl, sPhone, oCdoMail, oCdoConf


	' Get Org name
	sOrgName = GetOrgName( Session("orgid") )

	' Get Org URL
	sCityUrl = GetOrgUrl( Session("orgid") )
	'sCityUrl = "www.montgomeryohio.org"

	' Get Org Default Phone
	sPhone = GetDefaultPhone( Session("orgid") )
	'sPhone = "5138912424"

	' Get From address
	'sFrom = GetDefaultEmail( session("orgid") )
	sFromEmail   = "noreplies@eclink.com"
	sFromName = sOrgName & " E-GOV WEBSITE"

	' Subject
	sSubject = "Cancellation of " & sClassName


	sBody = "<p>Thank you for showing an interest in " & sClassName & " offered through the " & sOrgName & ". "
	sBody = sBody & "This event has been cancelled for the following reason. </p>" & vbcrlf & vbcrlf 
	sBody = sBody & "<p>" & sCancelReason & "</p>" & vbcrlf & vbcrlf 
	sBody = sBody & "<p>Your refund is being processed and should reach your house within two weeks time.</p>" & vbcrlf & vbcrlf 
	sBody = sBody & "<p>I encourage you to consider participation in one of the other programs being offered that may meet with "
	sBody = sBody & "your interests, needs and schedule.   A complete list of current programming available is listed on our website at " & sCityUrl 
	sBody = sBody & " with online registration services available for your convenience.</p>" & vbcrlf & vbcrlf 
	sBody = sBody & "<p>Our goal is to offer a variety of programs that meet with the needs and expectations of our customers.  "
	sBody = sBody & "Any feedback you have on the types, times, and other elements of the programs is appreciated and useful in "
	sBody = sBody & "the development of future classes. "   '   An evaluation form is attached and your comments are valued.
	sBody = sBody & "If you have any questions, please contact us at "
	sBody = sBody & FormatPhone( sPhone ) & "</p>" & vbcrlf & vbcrlf
	sBody = sBody & "<p>Thank you.</p>"


	sendEmail "", sToEmail, "", sSubject, sBody, clearHTMLTags( sBody ), "N"

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetFromAddress()
'--------------------------------------------------------------------------------------------------
Function GetFromAddress( iOrgId )
	Dim sSql, oFrom

	' get the email to send the admin message to
	sSQL = "SELECT assigned_email FROM dbo.egov_paymentservices where orgid = " & iOrgId & " and paymentservice_type = 4" 

	Set oFrom = Server.CreateObject("ADODB.Recordset")
	oFrom.Open sSQL, Application("DSN"), 0, 1
	
	If oFrom("assigned_email") = "" Or isNull(oFrom("assigned_email")) Then 
		GetFromAddress = "jstullenberger@eclink.com" ' NEED TO HAVE A DEFAULT INSTITUTION EMAIL ADDRESS
	Else 
		GetFromAddress = oFrom("assigned_email") ' ASSIGNED ADMIN USER EMAIL
	End If
	
	oFrom.close
	Set oFrom = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetOrgUrl( iClassId )
'--------------------------------------------------------------------------------------------------
Function GetOrgUrl( iOrgId )
	Dim sSql, oName

	sSql = "SELECT OrgEgovWebsiteURL FROM organizations WHERE orgid = " & iOrgId

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN"), 0, 1

	If Not oName.EOF Then
		GetOrgUrl = "<a href=""" & oName("OrgEgovWebsiteURL") & "/classes/class_list.asp"">" & oName("OrgEgovWebsiteURL") & "/classes/class_list.asp</a>"
	Else 
		GetOrgUrl = ""
	End If 

	oName.Close
	Set oName = Nothing

End Function 



%>

<!--#Include file="class_global_functions.asp"-->  

<!-- #include file="../includes/common.asp" //-->

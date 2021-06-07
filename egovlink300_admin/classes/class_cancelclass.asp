<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_cancelclass.asp
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
	Dim iClassId, bIsParent, sCancelReason

	iClassId = request("classid")
	sCancelReason = DBsafe( CStr(request("cancelreason")) )

	' Get isParent
	bIsParent = GetClassIsParent( iClassId )

	' update egov_class
	Cancel_Class iClassId, sCancelReason 

	' send email to class
	If request("emailclass") = "on" Then
		'sendCancelMail( iClassId )
	End If 

	' update any children
	If bIsParent Then
		Cancel_Children iClassId, sCancelReason
		' send email to children
		'sendChildCancelMail( iClassId )
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
Sub Cancel_Class( iClassId, sCancelReason )
	Dim oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")

		' Update the class table
		.CommandText = "Update egov_class Set statusid = 2, cancelreason = '" & sCancelReason & "', canceldate = GetDate() Where classid = " & iClassId
		.Execute
	End With
	Set oCmd = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub Cancel_Children( iParentClassId, sCancelReason )
'--------------------------------------------------------------------------------------------------
Sub Cancel_Children( iParentClassId, sCancelReason )
	Dim oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")

		' Update the class table
		.CommandText = "Update egov_class Set statusid = 2, cancelreason = '" & sCancelReason & "', canceldate = Getdate() Where parentclassid = " & iParentClassId
		.Execute
	End With
	Set oCmd = Nothing

End Sub 



%>

<!--#Include file="class_global_functions.asp"-->  


<!-- #include file="../includes/common.asp" -->
<!-- #include file="class_global_functions.asp" -->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: MultiClassEmailSend.asp
' AUTHOR: Steve Loar
' CREATED: 02/19/2013
' COPYRIGHT: Copyright 2013 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module sends emails to several class/event rosters
'
' MODIFICATION HISTORY
' 1.0	02/19/2013	Steve Loar - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iSelectedClassIds, sMessageBody, sSubject, sFromEmail, sFromName, iSentCount, sWhere

iSelectedClassIds = request("selectedclassids") ' this is a list of selected classes separated by commas
sMessageBody = request("messagebody")
sSubject     = request("subject")
sFromEmail   = request("fromemail")
sFromName    = request("fromname")
iSentCount   = 0
sWhere       = ""

Select Case clng(request("sendto"))
	Case 1
		sWhere = " AND L.status = 'ACTIVE' "
	Case 2
		sWhere = " AND L.status = 'WAITLIST' "
	Case 3
		sWhere = " AND (L.status = 'ACTIVE' OR L.status = 'WAITLIST') "
End Select 

iSentCount = SendClassMail( iSelectedClassIds, sMessageBody, sSubject, sFromEmail, sFromName, sWhere )

response.redirect "MultiClassEmail.asp?selectedclassids=" & iSelectedClassIds & "&sentcount=" & iSentCount & "&success=SS&fromemail=" & sFromEmail & "&fromname=" & sFromName & "&subject=" & sSubject & "&messagebody=" & sMessageBody & "&sendto=" & request("sendto") & "&classcount=" & request("classcount")
'response.write "done!"




'------------------------------------------------------------------------------
Function SendClassMail( ByVal iSelectedClassIds, ByVal sMessageBody, ByVal sSubject, ByVal sFromEmail, ByVal sFromName, ByVal sWhere )
	Dim sSql, oRs, iSentCount

	iSentCount = 0

	sSql = "SELECT DISTINCT ISNULL(U.useremail, U2.useremail) AS useremail "
	sSql = sSql & "FROM egov_class_list L, egov_users U, egov_users U2 "
	sSql = sSql & "WHERE L.attendeeuserid = U.userid AND L.userid = U2.userid "
	sSql = sSql & "AND L.classid IN ( " & iSelectedClassIds & " ) "
	sSql = sSql & sWhere
	sSql = sSql & "AND ( U.emailnotavailable = 0 OR U2.emailnotavailable = 0 ) "
	sSql = sSql & "AND ( U.useremail IS NOT NULL OR U2.useremail IS NOT NULL )"
	'response.write "sql = " & sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		If oRs("useremail") <> "" Then 
			'response.write "sending to " & oRs("useremail") & "<br /><br />"
			SendAnEmail sFromName, sFromEmail, sSubject, sMessageBody , oRs("useremail")
			iSentCount = iSentCount + 1
		End If 
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

	SendClassMail = iSentCount

End Function 


'------------------------------------------------------------------------------
Sub SendAnEmail( ByVal sFromName, ByVal sFromEmail, ByVal sSubject, ByVal sHTMLBody, ByVal sSendToEmail )
	Dim lcl_from_email

	If sSendToEmail <> "" Then 
		lcl_from_email = sFromName & " <" & sFromEmail & ">"

		'Send the email
		sendEmail lcl_from_email, sSendToEmail, "", sSubject, sHTMLBody, "", "Y"
	End If 

End Sub 


'------------------------------------------------------------------------------
Sub dtb_debug(p_value)
	Dim sSql, oDTB

	sSql = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"

	set oDTB = Server.CreateObject("ADODB.Recordset")
	oDTB.Open sSql, Application("DSN"), 0, 1

End Sub



%>

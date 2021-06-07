<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitnumberupdate.asp
' AUTHOR: Steve Loar
' CREATED: 08/25/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This updates the permit number. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   08/25/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, iPermitNumberYear, iPermitNumber, sSql

iPermitId = CLng(request("permitid"))
iPermitNumberYear = CLng(request("permitnumberyear"))
iPermitNumber = CLng(request("permitnumber"))

sSql = "UPDATE egov_permits SET permitnumberyear = " & iPermitNumberYear & ", permitnumber = " & iPermitNumber & " WHERE permitid = " & iPermitId
RunSQL sSql

' Push out the expiration date and reset the expiration flag
PushOutPermitExpirationDate( iPermitId )

'return back the permit number so the calling page can update it's parent page
response.write GetPermitNumber( iPermitId )


%>
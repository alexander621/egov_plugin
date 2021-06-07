<!-- #include file="../includes/common.asp" //-->

<%
Dim iWaiverId, sName, sType, sDescription, sURL, bRequired, sBody

iWaiverId = CLng(request("iWaiverId"))
sName = request("sName")
sType = request("sType")
sDescription = request("sDescription")
sURL = request("sURL")
bRequired = request("bRequired")
sBody = request("sBody")


sName = DBsafe( sName )
sDescription = DBsafe( sDescription )
sBody = DBsafe( sBody )
sURL = DBsafe( sURL )

If bRequired = "on" Then 
	bRequired = 1
Else
	bRequired = 0
End If

If CLng(iWaiverId) = CLng(0) Then
	' Insert new records
	sSql = "INSERT INTO egov_class_waivers (OrgID,waivername,waivertype,waiverdescription,waiverurl,isrequired,waiverbody ) Values (" & Session("OrgID") & ",'" & sName & "','" & sType & "','" & sDescription & "','" & sURL & "'," & bRequired & ", '" & sBody & "')"
Else 
	' Update existing records
	sSql = "UPDATE egov_class_waivers SET waivername='" & sName & "', waivertype='" & sType & "', waiverdescription='" & sDescription & "', waiverurl='" & sURL & "', isrequired=" & bRequired & ", waiverbody = '" & sBody & "' WHERE Waiverid = " & iWaiverId 
End If

'	response.write sSQL

RunSQLStatement sSql

' REDIRECT TO waiver management page
response.redirect "class_waivers.asp"


%>

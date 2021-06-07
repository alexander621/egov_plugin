<%
Function ProcessRecords()
	set oCmd = server.createobject("adodb.connection")
	set oCmd2 = server.createobject("adodb.connection")
	Set oTableRowIDs = CreateObject("Scripting.Dictionary")
	Set oRs = Server.CreateObject("ADODB.Recordset")
	rowid=0
	for each fieldname in Request.Form	
		oCmd.open Application("DSN")
		origfieldname = fieldname
		'get the field value
		fieldvalue = DBsafe(request(fieldname))
		'response.write fieldname &"="& fieldvalue & "<br>"

		'check and see if this form value is important
		tablenameid = ""

		isinstring = false
		if instr(fieldname,"ef:") or (right(fieldname,2) = "id" AND right(fieldname,5) <> "orgid") or left(fieldname,5) = "skip_" then
			isinstring = true
			'DEBUG CODE: response.write fieldname & "=TRUE" & "<br>"
		end if
		if Instr(fieldname,"_") and isinstring = false  then
			'split the field name into two important parts
			adbaddress = split(fieldname,"_")
			fieldname = ""
			tablename = ""
			islowertable = False
			idcolumn = request.form("columnnameid")

			for i = 0 to UBOUND(adbaddress) - 1
				'response.write adbaddress(i) & "<BR>"
				if tablename = "" then
					tablename = adbaddress(i)
				else
					tablename = tablename & "_" & adbaddress(i)
				end if
			next
			fieldname = adbaddress(i)
			tablenameid = tablename & fieldname
			tablenameid = idcolumn

			'response.write tablenameid


			'extract tablerowid if it exists
			tablerowid = 0
			if oTableRowIDs.Item(tablenameid) <> "" then
				tablerowid = clng(oTableRowIDs.Item(tablenameid))
			end if
		
		
			
			'NOW DO DO DB STUFF

			'UPDATE ROWS CASE
			if request.form(tablenameid) <> "" or tablerowid <> 0 or lowertablerowid <> 0 then
				'check for row id
				if request.form(tablenameid) <> "" then
					rowid = clng(request.form(tablenameid))
				elseif tablerowid <> 0 then
					rowid = tablerowid
				else
					rowid = lowertablerowid
				end if
				'if so then update it
				sSQL = "UPDATE " & tablename & " SET " & fieldname & "='" & fieldvalue & "' WHERE " & idcolumn & "=" & rowid
				'response.write sSQL & "<br>"
				'response.end
				oCmd.Execute(sSQL)

			'INSERT ROWS CASE
			elseif fieldvalue <> "" and fieldvalue <> "off" and fieldvalue <> "0" then
				tableidrequest = adbaddress(0) & "id"
				parenttableid = 0
				if islowertable = true then
					parenttableid = clng(oTableRowIDs.Item(tableidrequest))
				end if

				'if isprovider = true then
					sSQL = "INSERT INTO " & tablename & " ("& fieldname &") VALUES('" & fieldvalue &  "')"
				'elseif islowertable = False then
					'sSQL = "INSERT INTO " & tablename & " (userid,"& fieldname &") VALUES('" & request.cookies("userid") & "','" & fieldvalue & "')"
				'else
					'sSQL = "INSERT INTO " & tablename & " ("& adbaddress(0) &"id,"& fieldname &") VALUES('"
					'if request(tableidrequest) <> "" then
 						'sSQL = sSQL & request(tableidrequest)
					'else
 						'sSQL = sSQL & parenttableid
					'end if
 					'sSQL = sSQL & "','" & fieldvalue & "')"
				'end if
				'response.write sSQL & "<br>"
				'response.end
				if fieldvalue <> "" and fieldvalue <> "off" then
					oCmd.Execute(sSQL)
				end if
				sSQL2 = "SELECT @@IDENTITY AS NewID"
				set rs = oCmd.Execute(sSQL2)
				rowid = rs.Fields("NewID").value
				if lowertable = true and newlowertablerowid <> 0 then
					tablenameid = tablename & "id" & newlowertablerowid
				end if

				oTableRowIDs.Item(tablenameid) = rowid
				ProcessRecords = rowid
			end if

		end if
		oCmd.close
	next
end function



'--------------------------------------------------------------------------------------------------
' Function DBsafe( strDB )
'--------------------------------------------------------------------------------------------------
Function DBsafe( strDB )
	Dim sNewString
	If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
	sNewString = Replace( strDB, "'", "''" )
	sNewString = Replace( sNewString, "<", "&lt;" )
	sNewString = Replace( sNewString, ">", "&gt;" )
	DBsafe = sNewString
End Function


'--------------------------------------------------------------------------------------------------
' Sub AddFamilyMember( iBelongsToUserId, sFirstName, sLastName, sRelationship, sBirthDate )
'--------------------------------------------------------------------------------------------------
Sub AddFamilyMember( iBelongsToUserId, sFirstName, sLastName, sRelationship, sBirthDate )
	' This function adds family members to the egov_familymembers table
	Dim sSql, oCmd
	
	sSql = "Insert Into egov_familymembers (firstname, lastname, birthdate, belongstouserid, relationship, userid) values ('"
	If sBirthDate <> "NULL" Then
		sSql = sSql & sFirstName & "', '" & sLastName & "', '" & sBirthDate & "', " & iBelongsToUserId & ", '" & sRelationship & "', " & iBelongsToUserId & " )"
	Else
		sSql = sSql & sFirstName & "', '" & sLastName & "', " & sBirthDate & ", " & iBelongsToUserId & ", '" & sRelationship & "', " & iBelongsToUserId & " )"
	End If 

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing
End Sub 


'-------------------------------------------------------------------------------------------------
' Function RunIdentityInsert( sInsertStatement )
'-------------------------------------------------------------------------------------------------
Function RunIdentityInsert( ByVal sInsertStatement )
	Dim sSql, iReturnValue, oInsert, oCmd

	iReturnValue = 0

	'response.write "<p>" & sInsertStatement & "</p><br /><br />"
	'response.flush

	'INSERT NEW ROW INTO DATABASE AND GET ROWID
	sSql = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"
	session("RunIdentityInsertSQL") = sSql 

	Set oCmd = Server.CreateObject("ADODB.Connection")
	oCmd.open Application("DSN")

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert = oCmd.Execute(sSql)

	iReturnValue = oInsert("ROWID")

	Set oInsert = Nothing
	Set oCmd = Nothing

	session("RunIdentityInsertSQL") = ""

	RunIdentityInsert = iReturnValue

End Function


'-------------------------------------------------------------------------------------------------
' Sub RunSQL( sSql )
'-------------------------------------------------------------------------------------------------
Sub RunSQL( ByVal sSql )
	Dim oCmd

'	response.write "<p>" & sSql & "</p><br /><br />"
'	response.flush

	session("RunSQL") = sSql

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing
	session("RunSQL") = ""

End Sub 


'-------------------------------------------------------------------------------------------------
' Function StateNotValid( sUserstate )
'-------------------------------------------------------------------------------------------------
Function StateNotValid( sUserstate )
	Dim bNotValid

	bNotValid = True  
	
	Select Case sUserstate
		Case "AL", "AK", "AR", "AZ", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID", "IL", "IN", "IA", "KS"
			bNotValid = False
		Case "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY"
			bNotValid = False
		Case "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY"
			bNotValid = False
		Case Else
			bNotValid = True 
	End Select 

	StateNotValid = bNotValid
End Function 


%>


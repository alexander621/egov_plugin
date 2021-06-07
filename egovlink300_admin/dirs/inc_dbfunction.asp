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
		
			if origfieldname = "egov_users_password" then
				fieldvalue = createHashedPassword(fieldvalue)
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

				if fieldname = "residenttype" and blnSpecialAddress then
					fieldvalue = strResidenttype
				end if


				'if so then update it
				If fieldvalue = "NULL" Then
					sSQL = "UPDATE " & tablename & " SET " & fieldname & " = NULL WHERE " & idcolumn & " = " & rowid
				Else 
					sSQL = "UPDATE " & tablename & " SET " & fieldname & " = '" & fieldvalue & "' WHERE " & idcolumn & " = " & rowid
				End If 
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
  If Not VarType( strDB ) = 8 Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function


'--------------------------------------------------------------------------------------------------
' Function AddFamilyMember( iBelongsToUserId, sFirstName, sLastName, sRelationship, sBirthDate, iUserId )
'--------------------------------------------------------------------------------------------------
Sub AddFamilyMember( iBelongsToUserId, sFirstName, sLastName, sRelationship, sBirthDate, iUserId )
	' This function adds family members to the egov_familymembers table
	Dim sSql

	' Handle the passing of blank birthdates
	If sBirthDate = "" Then
		sBirthDate = "NULL"
	End If 
	
	sSql = "Insert Into egov_familymembers (firstname, lastname, birthdate, belongstouserid, relationship, userid) values ('"
	If sBirthDate <> "NULL" Then 
		sSql = sSql & DBsafe( sFirstName ) & "', '" & DBsafe( sLastName ) & "', '" & sBirthDate & "', " & iBelongsToUserId & ", '" & sRelationship & "', " & iUserId & " )"
	Else
		sSql = sSql & DBsafe( sFirstName ) & "', '" & DBsafe( sLastName ) & "', " & sBirthDate & ", " & iBelongsToUserId & ", '" & sRelationship & "', " & iUserId & " )"
	End If 

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub FamilyMemberUpdate( iUserId, sFirstName, sLastName, sRelationship, ByVal sBirthDate )
'--------------------------------------------------------------------------------------------------
Sub FamilyMemberUpdate( iUserId, sFirstName, sLastName, sRelationship, ByVal sBirthDate )
	Dim sSql, oCmd

	' Handle the passing of blank birthdates
	If sBirthDate = "" Then
		sBirthDate = "NULL"
	End If 

	If sBirthDate <> "NULL" Then
		sBirthDate = "'" & sBirthDate & "'"
	End If 
	
	sSql = "Update egov_familymembers Set firstname = '" & DBsafe( sFirstName )
	sSql = sSql & "', lastname = '" & DBsafe( sLastName )
	sSql = sSql & "', birthdate = " & sBirthDate
	sSql = sSql & ",  relationship = '" & sRelationship
	sSql = sSql & "'  Where userid = " & iUserId

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

End Sub

%>


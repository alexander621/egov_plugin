<%
Function SQLText( VarName )
Dim val

  val = Request.Form(VarName)
  If val & "" = "" Then
    SQLText = "NULL"
  Else
	  val = Replace(val, vbCrLf, "<br>")
		SQLText = "'" & Replace(val, "'", "''") & "'"
  End If
End Function

Function URLVerify( VarName )
Dim val

  val = Request.Form(VarName)
	If val & "" = "" Then
		URLVerify = "NULL"
	Else
		If Left(val, 7) <> "http://" Then val = "http://" & val
		URLVerify = "'" & Replace(val, "'", "''") & "'"
	End If
End Function


Dim cnn, sql, userid

userid = Request("userid")

sql = "UPDATE Users SET FirstName = " & SQLText("FirstName") & ", LastName = " & SQLText("LastName") & "," _
    & "EmailName = " & SQLText("EmailName") & " WHERE UserID = " & userid

Set cnn = Server.CreateObject("ADODB.Connection")
cnn.Open Application("ECMDConnectionString")
cnn.Execute (sql)
cnn.Close

sql = "SELECT JobTitle FROM UserInformation WHERE UserID = " & userid
cnn.Open Application("dbCnnStr_Tools")
Set rst = cnn.Execute (sql)
f = rst.EOF
rst.Close
Set rst = Nothing

If NOT f Then
	sql = "UPDATE UserInformation SET JobTitle = " & SQLText("JobTitle") & ", Department = " & SQLText("Department") & "," _
			& "Nickname = " & SQLText("NickName") & ", HomeAddress = " & SQLText("HomeAddress") & ", BusinessAddress = " & SQLText("BusinessAddress") & "," _
			& "HomeNumber = " & SQLText("HomeNumber") & ", BusinessNumber = " & SQLText("BusinessNumber") & "," _
			& "MobileNumber = " & SQLText("MobileNumber") & ", PagerNumber = " & SQLText("PagerNumber") & ", FaxNumber = " & SQLText("FaxNumber") & "," _
			& "WebPage = " & URLVerify("WebPage") & ", Birthday = " & SQLText("Birthday") & ", SpouseName = " & SQLText("SpouseName") & "," _
			& "SpouseBirthday = " & SQLText("SpouseBirthday") & ", ManagerName = " & SQLText("ManagerName") & "," _
			& "AOLScreenName = " & SQLText("AOLScreenName") & ", ICQNumber = " & SQLText("ICQNumber") & ", StartDate = " & SQLText("StartDate") & "," _
			& "WorkStatus = " & SQLText("WorkStatus") & ", Custom1Name = " & SQLText("Custom1Name") & ", Custom1Value = " & SQLText("Custom1Value") & "," _
			& "Custom2Name = " & SQLText("Custom2Name") & ", Custom2Value = " & SQLText("Custom2Value") & "," _
			& "Custom3Name = " & SQLText("Custom3Name") & ", Custom3Value = " & SQLText("Custom3Value") & " " _
			& "WHERE UserID = " & userid
Else
  sql = "INSERT INTO UserInformation(UserID, JobTitle, Department, Nickname, HomeAddress, BusinessAddress, HomeNumber, BusinessNumber, "_
	    & "MobileNumber, PagerNumber, FaxNumber, WebPage, Birthday, SpouseName, SpouseBirthday, ManagerName, AOLScreenName, "_
			& "ICQNumber, StartDate, WorkStatus, Custom1Name, Custom1Value, Custom2Name, Custom2Value, Custom3Name, Custom3Value) " _
			& "VALUES(" & userid & "," & SQLText("JobTitle") & "," & SQLText("Department") & "," & SQLText("Nickname") & "," & SQLText("HomeAddress") _
			& "," & SQLText("BusinessAddress") & "," & SQLText("HomeNumber") & "," & SQLText("BusinessNumber") & "," & SQLText("MobileNumber") _
			& "," & SQLText("PagerNumber") & "," & SQLText("FaxNumber") & "," & URLVerify("WebPage") & "," & SQLText("Birthday") _
			& "," & SQLText("SpouseName") & "," & SQLText("SpouseBirthday") & "," & SQLText("ManagerName") & "," & SQLText("AOLScreenName") _
			& "," & SQLText("ICQNumber") & "," & SQLText("StartDate") & "," & SQLText("WorkStatus") & "," & SQLText("Custom1Name") _
			& "," & SQLText("Custom1Value") & "," & SQLText("Custom2Name") & "," & SQLText("Custom2Value") & "," & SQLText("Custom3Name") _
			& "," & SQLText("Custom3Value") & ")"
End If

cnn.Execute (sql)
cnn.Close
Set cnn = Nothing

Response.Redirect "userinfo.asp?id=" & userid
%>
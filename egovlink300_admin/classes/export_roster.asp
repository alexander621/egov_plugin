<!-- #include file="class_global_functions.asp" //-->
<%
Dim fromDate, toDate

'SET UP PAGE OPTIONS
' sDate = Month(Date()) & Day(Date()) & Year(Date())
sDate = year(date()) & month(date()) & day(date())
sTime = hour(time()) & minute(time()) & second(time())
server.scripttimeout = 4800
response.ContentType = "application/msexcel"
response.AddHeader "Content-Disposition", "attachment;filename=ROSTER_EXPORT_" & sDate & "_" & sTime & ".xls"

'Determine if the list should be limited by a specific time
iTimeID = request("timeid")

If Request("fromDate") <> "" Then 
	fromDate = Request("fromDate")
Else
	fromDate = ""
End If 

If Request("toDate") <> "" Then 
	toDate = Request("toDate")
Else
	toDate = ""
End If 


'CREATE CSV FILE FOR DOWNLOAD
 CreateDownload iTimeID, fromDate, toDate

'------------------------------------------------------------------------------------------------------------
' CreateDownload p_timeid, fromDate, toDate 
'------------------------------------------------------------------------------------------------------------
 Sub CreateDownload( ByVal p_timeid, ByVal fromDate, ByVal toDate )
	Dim sWhereClause, sSql, oRs

	If toDate <> "" And fromDate <> "" Then
		sWhereClause = " AND signupdate >= '" & fromDate & "' AND signupdate < '" & DateAdd( "d", 1, CDate(toDate) ) & "' "
	Else
		sWhereClause = ""
	End If 

	'WRITE COLUMN HEADINGS
	response.write "<table border=""1""><tr>"
	response.write "<th>FIRST NAME ENROLLEE</th>"
	response.write "<th>LAST NAME</th>"           
	response.write "<th>AGE</th>"                 
	response.write "<th>HOME PHONE</th>"          
	response.write "<th>EMERGPHONE</th>"          
	response.write "<th>EMERGENCYCONTACT</th>"    
	response.write "<th>ADDRESS</th>"             
	response.write "<th>CITY</th>"                
	response.write "<th>STATE</th>"               
	response.write "<th>ZIP</th>"                 
	response.write "<th>EMAIL</th>"               
	response.write "<th>PAYEE NAME</th>" 
	response.write "<th>CLASS/EVENT</th>"
	response.write "<th>ACTIVITY #</th>"                
	response.write "<th>FEEPAID</th>"             
	response.write "<th>RECEIPT#</th>"            
	response.write "<th>DATEPAID</th>"            
	response.write "<th>RESIDENCY</th>"           
	response.write "<th>STATUS</th>"              
	response.write "<th>INSTRUCTOR</th>"          
	Response.write "</tr>"
	response.flush

	'LOOP THRU EACH CLASS CHECKED AND DISPLAY CLASS INFORMATION AND ROSTER
	For Each item In request("classid")
 
		sSql = "SELECT "
		sSql = sSql & " firstname, "
		sSql = sSql & " lastname, "
		sSql = sSql & " birthdate, "
		sSql = sSql & " userhomephone, "
		sSql = sSql & " (SELECT isnull(emergencyphone,'') FROM egov_users WHERE userid = familymemberuserid) as emergencyphone, "
		sSql = sSql & " (SELECT isnull(emergencycontact,'') FROM egov_users WHERE userid = familymemberuserid) as emergencycontact, "
		sSql = sSql & " (SELECT isnull(dbo.fn_BuildAddress("
		sSql = sSql &                                     "isnull(userstreetnumber,''),"
		sSql = sSql &                                     "isnull(userstreetprefix,''),"
		sSql = sSql &                                     "isnull(useraddress,''),"
		sSql = sSql &                                     "'',''"
		sSql = sSql &                                    ")"
		sSql = sSql &                 ",'') FROM egov_users where userid = familymemberuserid"
		sSql = sSql &  ") as user_address, "
		sSql = sSql & " (SELECT isnull(usercity,'') FROM egov_users where userid = familymemberuserid) as user_city, "
		sSql = sSql & " (SELECT isnull(userstate,'') FROM egov_users where userid = familymemberuserid) as user_state, "
		sSql = sSql & " (SELECT isnull(userzip,'') FROM egov_users where userid = familymemberuserid) as user_zip, "
		sSql = sSql & " useremail, "
		sSql = sSql & " userfname + ' ' + userlname as payee_name, "
		sSql = sSql & " classname, activityno, "
		sSql = sSql & " amount, "
		sSql = sSql & " paymentid, "
		sSql = sSql & " signupdate, "
		sSql = sSql & " description, "
		sSql = sSql & " status, "
		sSql = sSql & " ltrim(rtrim(instructor_firstname + ' ' + instructor_lastname)) as instructor "
		sSql = sSql & " FROM egov_class_roster "
		sSql = sSql & " WHERE (classid = " & item & " ) "

		If p_timeid <> "" Then 
			sSql = sSql & " AND classtimeid = " & p_timeid
		End If

		' filter by registration date
		sSql = sSql & sWhereClause

		'sSql = sSql & " ORDER BY status,signupdate,userlname, userfname"
		'sSql = sSql & " ORDER BY signupdate, userlname, userfname"
		sSql = sSql & " ORDER BY userlname, userfname, signupdate"

		'response.write "<br /><br />" & sSql & "<br /><br />"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.eof Then 
			'WRITE DATA
			Do While Not oRs.EOF
				Response.write "<tr>"
				For Each fldLoop In oRs.Fields

					'Birthdate
					If fldLoop.name = "birthdate" Then 
						If Not IsNull(oRs("birthdate")) AND IsDate(oRs("birthdate")) Then 
							iMonths = DateDiff("m", oRs("birthdate"), Now())
							If iMonths = 0 Then 
								iMonths = 1 
							End If 
							iAge = FormatNumber(iMonths / 12, 1)
						Else
							iAge = 21
						End If 

						If iAge >= 18 Then 
							iAge = "Adult"
						End If 
							sFieldValue = iAge

					'Amount
					ElseIf fldLoop.name = "amount" Then 
						If fldLoop.value <> "" Then 
							iAmount = FormatCurrency(fldLoop.value)
						Else 
							iAmount = ""
						End If 

						sFieldValue = iAmount

					'Emergency Contact Info
					ElseIf fldLoop.name = "emergencyphone" Or fldLoop.name = "userhomephone" Then 
						sFieldValue = FormatPhone(fldLoop.value)
					Else 
						sFieldValue = Trim(fldLoop.Value)
					End If 

					'REMOVE LINE BREAKS
					If Not IsNull(sFieldValue) Then
						sFieldValue = Replace(sFieldValue,Chr(10),"")
						sFieldValue = Replace(sFieldValue,Chr(13),"")
						sFieldValue = Replace(sFieldValue,"default_novalue","")
						sFieldValue = Replace(sFieldValue,"<p><b>","")
						sFieldValue = Replace(sFieldValue,"</b><br></p>"," [] ")
						sFieldValue = Replace(sFieldValue,"</b><br>"," [")
						sFieldValue = Replace(sFieldValue,"</p>","] ")
					End If 

					response.write "<td>" & sFieldValue & "</td>"

				Next 

				Response.write "</tr>"
				response.flush
				
				oRs.MoveNext
			Loop 

		End If

		oRs.Close 
		Set oRs = Nothing

	Next 

End Sub


%>

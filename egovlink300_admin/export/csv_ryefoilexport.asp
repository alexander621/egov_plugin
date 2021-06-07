<!--#include file="../../egovlink300_global/includes/inc_rye.asp"-->
<%
'SET UP PAGE OPTIONS
' sDate = Month(Date()) & Day(Date()) & Year(Date())
sDate = year(date()) & month(date()) & day(date())
sTime = hour(time()) & minute(time()) & second(time())
server.scripttimeout = 4800
response.ContentType = "application/msexcel"
response.AddHeader "Content-Disposition", "attachment;filename=EGOV_FOIL_EXPORT_" & sDate & "_" & sTime & ".CSV"

'CREATE CSV FILE FOR DOWNLOAD
' CreateDownload

'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
' FUNCTION CREATEDOWNLOAD()
'------------------------------------------------------------------------------------------------------------
' Sub CreateDownload()

Set oSchema = Server.CreateObject("ADODB.Recordset")

'response.write request.cookies("User")("UserID")
'response.end

sSQL = session("DISPLAYQUERY")

sSQL = "SELECT CAST([action_autoid] as varchar(50)) + RIGHT('00'+CAST(DATEPART(hh, submit_date) as varchar(2)),2) + RIGHT('00'+CAST(DATEPART(n, submit_date) as varchar(2)),2) as TrackingNumber, " & vbcrlf _
	& " submit_date as [Submission Date], comment AS [Records Requested], " & vbcrlf _
	& " (SELECT TOP 1 u.FirstName + ' ' + u.LastName FROM egov_action_responses ar INNER JOIN Users u ON u.UserID = ar.action_userid where ar.action_autoid = egov_action_request_view.action_autoid and ar.action_userid <> 11068 ORDER BY action_editDate) as [Department Assigned], " & vbcrlf _
	& "  userfname + ' ' + userlname as [Requestor],complete_date As  [Completion Date], status as [Status],[assigned_userID], [assignedname], due_date as [Current Due Date] " & vbcrlf _
	& mid(sSQL, instr(sSQL,"FROM egov_action_request_view"))

'sSQL = replace(sSQL, "FROM egov_action_request_view ", ",(SELECT TOP 1 u.FirstName + ' ' + u.LastName FROM egov_action_responses ar INNER JOIN Users u ON u.UserID = ar.action_userid where ar.action_autoid = egov_action_request_view.action_autoid and ar.action_userid <> 11068 ORDER BY action_editDate) as Department FROM  egov_action_request_view	")

'response.write sSQL
'response.end
oSchema.Open sSQL, Application("DSN"), 3, 1

If Not oSchema.EOF Then

	'WRITE COLUMN HEADINGS
	For Each fldLoop in oSchema.Fields
		if fldLoop.Name <> "assigned_userID" and fldLoop.Name <> "assignedname" then response.write chr(34) & fldLoop.Name & chr(34) & ","
	Next
	response.write """Original Due Date"""
	response.write vbcrlf
	response.flush

	'WRITE DATA
	do while not oSchema.eof
		for each fldLoop in oSchema.Fields

				sFieldValue = trim(fldLoop.Value)

				'REMOVE LINE BREAKS
			if not isnull(sFieldValue) then
					sFieldValue = replace(sFieldValue,chr(10),"")
					sFieldValue = replace(sFieldValue,chr(13),"")
					sFieldValue = replace(sFieldValue,chr(34),"'")
					sFieldValue = replace(sFieldValue,"default_novalue","")
					sFieldValue = replace(sFieldValue,"<p><b>","")
					sFieldValue = replace(sFieldValue,"</b><br></p>"," [] ")
					sFieldValue = replace(sFieldValue,"</b><br>"," [")
					sFieldValue = replace(sFieldValue,"</p>","] ")
			
					if fldLoop.Name = "Records Requested" and instr(sFieldValue,"Describe records being sought - One request per submission. [") > 0  then
						sFieldValue = mid(sFieldValue, instr(sFieldValue, "Describe records being sought - One request per submission. [") + 61, len(sFieldValue))
						'on error resume next
						sFieldValue = left(sFieldValue, instr(sFieldValue, "] Please indicate your preference:") - 1)
						'if err.number <> 0 then
							'sFieldValue = "ERROR: " & sFieldValue
						'end if
						'on error goto 0
					end if

			end if
			if fldLoop.Name = "Department Assigned" then
				'sFieldValue = "TEST" & (isnull(sFieldValue) or sFieldValue = "") & " = " & oSchema("assignedname") & " => "
				if (isnull(sFieldValue) or sFieldValue = "") and oSchema("assignedname") <> "Rye Foil" then
					sFieldValue = oSchema("assignedname")
				end if
			end if

			if fldLoop.Name = "Current Due Date" and (isnull(sFieldValue) or sFieldValue = "") then
				DueDate = oSchema("Submission Date")
				sFieldValue = GetFOILDueDate(DueDate)
			end if


'LEFT OFF HERE!!!!!
'             if sFieldValue <> "" then
'                sFieldValue = replace(sFieldValue,"&quot;","""")
'             end if

			if fldLoop.Name <> "assigned_userID" and fldLoop.Name <> "assignedname" then	response.write chr(34) & sFieldValue & chr(34) & ","
		next

		DueDate = oSchema("Submission Date")
		DueDate = GetFOILDueDate(DueDate)
		response.write """" & DueDate & """"
		response.write vbcrlf
		response.flush

		oSchema.MoveNext
	Loop
Else

 'NO DATA

End If
Set oSchema = Nothing

' End Sub

sub dtb_debug(p_value)

  if p_value <> "" then
     lcl_value = replace(p_value,"'","''")
  else
     lcl_value = ""
  end if

  sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & lcl_value & "')"
  set rs = Server.CreateObject("ADODB.Recordset")
  rs.Open sSQL, Application("DSN"), 3, 1

end sub
%>

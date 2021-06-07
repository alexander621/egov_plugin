<!-- #include file="../includes/common.asp" //-->
<%
'SET UP PAGE OPTIONS
' sDate = Month(Date()) & Day(Date()) & Year(Date())
 sDate = year(date()) & month(date()) & day(date())
 sTime = hour(time()) & minute(time()) & second(time())
 server.scripttimeout = 4800
 response.ContentType = "application/msexcel"
 response.AddHeader "Content-Disposition", "attachment;filename=EGOV_EXPORT_ACTLOG_" & sDate & "_" & sTime & ".CSV"

'CREATE CSV FILE FOR DOWNLOAD
 if OrgHasFeature("activity_log_download") AND UserHasPermission( session("userid"), "activity_log_download" ) then
    Call CreateDownload()
 else
  	 response.redirect "../permissiondenied.asp"
 end if

'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
' FUNCTION CREATEDOWNLOAD()
'------------------------------------------------------------------------------------------------------------
 Sub CreateDownload()
  	Set oSchema = Server.CreateObject("ADODB.Recordset")

   sSQL = session("ACTLOGQUERY")

'response.write cleanUpHTMLTags(sSQL) & vbcrlf

  	oSchema.Open sSQL, Application("DSN"), 3, 1

  	If NOT oSchema.EOF Then
	
     'WRITE COLUMN HEADINGS
    		response.write "E-Gov Link Tracking #" & ", "
      response.write "Status - Date/Time"    & ","
      response.write "Activity Log"          & ", "
      response.write vbcrlf
      response.flush

      lcl_action_autoid = oSchema("action_autoid")

    	Do While NOT oSchema.EOF
       		if CLng(lcl_action_autoid) <> CLng(oSchema("action_autoid")) then
            		if lcl_action_form_resolved_status = "Y" then
               			lcl_status = "RESOLVED"
            		else
               			lcl_status = "SUBMITTED"
            		end if

			'DISPLAY SUBMIT DATE TIME AND USER
       			response.write lcl_tracking_number & ", "
       			response.write sSubmitName & " - " & UCASE(lcl_status) & " - " & lcl_submit_date
	            response.write "," & vbcrlf
		    response.flush
       		end if

        	'Build the status line
         	lcl_status_line = buildStatusLine(oSchema("users_first_name"),oSchema("users_last_name"),oSchema("action_status"),oSchema("status_name"),oSchema("action_editdate"))

        	'Build the Internal, External Comments, and Note to Citizen
         	lcl_external_comment = "Note to Citizen: "
         	lcl_internal_comment = "Internal Note: "

         	if trim(oSchema("user_name")) <> "" then
            		lcl_action_citizen   = trim(oSchema("user_name")) & ": " 
         	else
            		lcl_action_citizen = "Note to Citizen: "
         	end if

         	if oSchema("action_externalcomment") <> "" then
            		lcl_external_comment = cleanUpHTMLTags(lcl_external_comment & oSchema("action_externalcomment"))
         	else
            		lcl_external_comment = ""
         	end if
		
         	if oSchema("action_citizen") <> "" then
            		lcl_action_citizen = cleanUpHTMLTags(lcl_action_citizen & oSchema("action_citizen"))
         	else
            		lcl_action_citizen = ""
         	end if
		
         	if oSchema("action_internalcomment") <> "" then
            		lcl_internal_comment = cleanUpHTMLTags(lcl_internal_comment & oSchema("action_internalcomment"))
         	else
            		lcl_internal_comment = ""
         	end if

        	'Get the employee or citizen that submitted the request.
         	if oSchema("employeesubmitid") < 0 OR IsNull(oSchema("employeesubmitid")) OR oSchema("employeesubmitid") = "" then
           		'Use Citizen Name as Submitter
            		sSubmitName = oSchema("user_name") & " (Citizen)"
         	else
           		'User Employee Name as Submitter
            		sSubmitName = oSchema("EmployeeSubmitName") & " (Admin Employee)"
         	end if

        	'Display the data
    		response.write oSchema("E-Gov Link Tracking #") & ", "
         	response.write lcl_status_line & ","
	
         	if trim(lcl_external_comment) <> "" then
            		response.write lcl_external_comment & ","
         	end if

         	if trim(lcl_action_citizen) <> "" then
            		response.write lcl_action_citizen & ","
         	end if

         	if trim(lcl_internal_comment) <> "" then
            		response.write lcl_internal_comment & ","
         	end if

         	response.write vbcrlf
		response.flush

        	'Set the "previous" variables to the current values.
        	'This will group the Activiy Log records by the action_autoid by putting a blank row in betwen the groups.
         	lcl_action_autoid               = oSchema("action_autoid")
         	lcl_tracking_number             = oSchema("E-Gov Link Tracking #")
         	lcl_action_form_resolved_status = oSchema("action_form_resolved_status")
         	lcl_submit_date                 = oSchema("submit_date")

      		oSchema.MoveNext
    	Loop

      if lcl_action_autoid <> "" AND oSchema.eof then
         if lcl_action_form_resolved_status = "Y" then
            lcl_status = "RESOLVED"
         else
            lcl_status = "SUBMITTED"
         end if

     			'DISPLAY SUBMIT DATE TIME AND USER
      			response.write lcl_tracking_number & ", "
         response.write sSubmitName & " - " & UCASE(lcl_status) & " - " & lcl_submit_date & ","
         response.write vbcrlf
	 response.flush
      end if

  	Else

     'NO DATA

   End If
   Set oSchema = Nothing

 End Sub

'---------------------------------------------------------------------------------
 function cleanUpHTMLTags(p_value)
    lcl_return = p_value

				lcl_return = replace(lcl_return,chr(10),"")
 			lcl_return = replace(lcl_return,chr(13),"")
    lcl_return = replace(lcl_return,"default_novalue","")
    lcl_return = replace(lcl_return,"<p><b>","")
    lcl_return = replace(lcl_return,"</b><br></p>"," [] ")
    lcl_return = replace(lcl_return,"</b><br>"," [")
    lcl_return = replace(lcl_return,"</p>","] ")
    lcl_return = replace(lcl_return,"<br>"," ")
    lcl_return = replace(lcl_return,"<BR>"," ")
    lcl_return = replace(lcl_return,"<br />"," ")
    lcl_return = replace(lcl_return,"<BR />"," ")
    lcl_return = replace(lcl_return,"<strong>","")
    lcl_return = replace(lcl_return,"</strong>","")
    lcl_return = replace(lcl_return,vbcrlf,"")
    lcl_return = replace(lcl_return,"""","""""")

    cleanUpHTMLTags = chr(34) & lcl_return & chr(34)

 end function

'-----------------------------------------------------------------------------------------
 function buildStatusLine(p_firstname,p_lastname,p_status,p_substatus,p_submitdate)
   'Get the sub-status and build it into the status line
    lcl_substatus_name = p_substatus

    if lcl_substatus_name <> "" then
    			lcl_substatus_name = " - " & lcl_substatus_name
    else
       lcl_substatus_name = ""
    end if

   'Format the status line
    lcl_status_line = ""

   'First Name
    if trim(p_firstname) <> "" then
       lcl_status_line = trim(p_firstname)
    end if

   'Last Name
    if trim(p_lastname) <> "" then
       if lcl_status_line <> "" then
          lcl_status_line = lcl_status_line & " " & trim(p_lastname)
       else
          lcl_status_line = trim(p_lastname)
       end if
    end if

   'Status and Sub-Status
    if trim(p_status) <> "" then
       if lcl_status_line <> "" then
          lcl_status_line = lcl_status_line & " - " & trim(p_status) & "[sub_status]"
       else
          lcl_status_line = trim(p_status) & "[sub_status]"
       end if
    end if

   'Submit Date
    if trim(p_submitdate) <> "" then
       if lcl_status_line <> "" then
          lcl_status_line = lcl_status_line & " - " & FormatDateTime(trim(p_submitdate),0)
       else
          lcl_status_line = FormatDateTime(trim(p_submitdate),0)
       end if
    end if

    if instr(lcl_status_line,"[sub_status]") > 0 then
       lcl_status_line = replace(lcl_status_line,"[sub_status]",lcl_substatus_name)
    end if

    buildStatusLine = lcl_status_line

 end function

sub dtb_debug(p_value)
  sSQLi = "INSERT INTO my_table_dtb(notes) VALUES ('" & REPLACE(p_value,"'","''") & "')"

  set rsi = Server.CreateObject("ADODB.Recordset")
  rsi.Open sSQLi, Application("DSN"), 3, 1
end sub
%>

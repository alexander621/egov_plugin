<!--#include file="../../includes/common.asp"//-->
<!--#include file="../action_line_global_functions.asp"//-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: CORRECTION_CONTACT_INFO_CGI.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/12/07
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  SAVE CONTACT INFORMATION
'
' MODIFICATION HISTORY
' 1.0	02/12/07	John Stullenberger - Initial Version
' 1.1 07/22/09 David Boyer - Miscellaneous
'                            a. Caught code up to our current standards
'                            b. Added action_line_global_functions.asp include file for the AddTaskComment routine
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

 lcl_requestid = request("irequestid")

 Call subSaveContactInfo()

'------------------------------------------------------------------------------
sub subSaveContactInfo()

	sSQL = "SELECT u.userfname, u.userlname, u.userbusinessname, u.useremail, u.userhomephone, u.userfax, "
 sSQL = sSQL & " u.useraddress, u.usercity, u.userstate, u.userzip, r.contactmethodid "
 sSQL = sSQL & " FROM egov_actionline_requests r "
 sSQL = sSQL &      " INNER JOIN egov_users u ON r.userid = u.userid "
 sSQL = sSQL & " WHERE r.action_autoid = " & request("irequestid")

	set oSave = Server.CreateObject("ADODB.Recordset")
	oSave.Open sSQL, Application("DSN"), 1, 2

 'oSave.CursorLocation = 3
 'oSave.Open sSQL, Application("DSN"), 1, 2

	if not oSave.eof then
  		sLogEntry               = ""
    lcl_new_contactmethodid = ""
		
		 'Check for different values and save
  		for each oColumn in oSave.fields
			
    			'Compare values - If changed then update record and log
     			if trim(oColumn.Value) <> trim(request(oColumn.Name)) OR IsNull(oColumn.Value) then

      				'Log changes
       				if oColumn.Name <> "contactmethodid" then

        					'All changes other than contact id
         					sFieldName = GetFieldDisplayName(oColumn.Name)

         					if sLogEntry = "" then
           						sLogEntry = "Edit Contact Information<br />" & sFieldName & ": " & sLogEntry & chr(34) & trim(oColumn.Value) & chr(34) & " changed to " & chr(34) & trim(request(oColumn.Name)) & chr(34)
              else
           						sLogEntry = sLogEntry & "<br /> " & sFieldName & ": " & chr(34) & trim(oColumn.Value) & chr(34) & " changed to " & chr(34) & trim(request(oColumn.Name)) & chr(34)
              end if

              oSave(oColumn.Name) = trim(request(oColumn.Name))

           else
        					'Special handling to get display name for contact method
         					if sLogEntry = "" then
           						sLogEntry = "Edit Contact Information<br /> Contact Method:" & sLogEntry & chr(34) & DisplayContactMethod(trim(oColumn.Value)) & chr(34) & " changed to " & chr(34) & DisplayContactMethod(trim(request(oColumn.Name))) & chr(34)
         					else
                 if  trim(request(oColumn.Name)) <> "0" _
                 AND trim(oColumn.Value) <> trim(request(oColumn.Name)) then
               						sLogEntry = sLogEntry & "<br /> Contact Method:" & chr(34) & DisplayContactMethod(trim(oColumn.Value)) & chr(34) & " changed to " & chr(34) & DisplayContactMethod(trim(request(oColumn.Name))) & chr(34)
                 end if
              end if

             'Set the contactmethodid value
              if trim(request(oColumn.Name)) <> "" then
                 lcl_new_contactmethodid = trim(request(oColumn.Name))
              end if
           end if
        end if
		  next

   	oSave.Update
 end if

	oSave.close
	set oSave = nothing

'Record save activity in log ONLY if something has changed
	if sLogEntry <> "" then
  		AddCommentTaskComment sLogEntry, sExternalMsg, request("status"), request("irequestid"), session("userid"), session("orgid"), request("substatus"), "", ""

  'Update the task sub-status
   if request("substatus") = "" then
      lcl_sub_status = 0
   else
      lcl_sub_status = request("substatus")
   end if

   sSQL = "UPDATE egov_actionline_requests "
   sSQL = sSQL & " SET sub_status_id = "   & lcl_sub_status
   sSQL = sSQL & " WHERE action_autoid = " & request("irequestid")

   set oUpdate2 = Server.CreateObject("ADODB.Recordset")
   oUpdate2.Open sSQL, Application("DSN") , 3, 1
   set oUpdate2 = nothing
 end if

'Update the contact method
 if lcl_new_contactmethodid <> "" then
    sSQL = "UPDATE egov_actionline_requests SET "
    sSQL = sSQL & " contactmethodid = "     & lcl_new_contactmethodid
    sSQL = sSQL & " WHERE action_autoid = " & request("irequestid")

    set oUpdateContactMethod = Server.CreateObject("ADODB.Recordset")
    oUpdateContactMethod.Open sSQL, Application("DSN"), 3, 1

    set oUpdateContactMethod = nothing
 end if

	response.redirect "../action_respond.asp?control=" & request("irequestid") & "&r=save&status="&request("status")

end sub

'------------------------------------------------------------------------------
function DBsafe( strDB )

  if not VarType( strDB ) = vbString then DBsafe = strDB : exit function
  DBsafe = Replace( strDB, "'", "''" )

end function

'------------------------------------------------------------------------------
function DisplayContactMethod(iValue)

	sSQL = "SELECT * FROM egov_contactmethods WHERE rowid = '" & iValue & "'"

	set oMethods = Server.CreateObject("ADODB.Recordset")
	oMethods.Open sSQL, Application("DSN") , 3, 1
	
	if not oMethods.eof then
  		iReturnValue = oMethods("contactdescription") 
	else
  		iReturnValue = "NOT SPECIFIED"
	end if

	set oMethods = nothing
	
	DisplayContactMethod = iReturnValue

end function

'------------------------------------------------------------------------------
function GetFieldDisplayName(sValue)

	sReturnValue = sValue

	arrNames = Array("userfname","userlname","userbusinessname","useremail","userhomephone","userfax","useraddress","usercity","userstate","userzip","contactmethodid")
	arrDisplayNames = Array("First Name","Last Name","Business Name","Email","Home Phone","Fax","Address","City","State","Zip","Contact Method")

'Loop thru arrays and find match
	for iLoop = 0 to UBOUND(arrNames)
		   if cstr(trim(arrNames(iLoop))) = cstr(trim(sValue)) then
     			sReturnValue = arrDisplayNames(iLoop)
     			exit for
   		end if
	next

	GetFieldDisplayName = sReturnValue

end function
%>
<!-- #include file="../includes/common.asp" //-->
<%
 lcl_formid = 0
 lcl_task   = ""

 if request("formid") <> "" then
    if not containsApostrophe(request("formid")) then
       lcl_formid = clng(request("formid"))
    end if
 end if

 if trim(request("task")) <> "" then
    lcl_task = trim(request("task"))
    lcl_task = ucase(lcl_task)
 end if

 if lcl_formid > 0 then
    updateFormOption lcl_formid, _
                     lcl_task
 else
    response.write "Failed to update form option - Error in AJAX Routine"
 end if

'------------------------------------------------------------------------------
sub updateFormOption(iFormID, _
                     iTask)

   dim sFormID, sTask

   sFormID              = 0
   sTask                = ""
   sDBColumnName        = ""
   sNewValue_isInternal = ""

   if iFormID <> "" then
      if not containsApostrophe(iFormID) then
         sFormID = clng(iFormID)
      end if
   end if

   if trim(iTask) <> "" then
      sTask = trim(iTask)
      sTask = ucase(sTask)
   end if

   if sFormID > 0 then

      sNewValue_isInternal = getNewFormValue(sTask, _
                                             sFormID)

      if sTask = "INTERNAL" then
         sDBColumnName = "action_form_internal"
      elseif sTask = "ENABLEFORM" then
         sDBColumnName = "action_form_enabled"
      elseif sTask = "DISPLAYONLIST" then
         sDBColumnName = "action_form_displayOnList"
      end if

      sSQL = "UPDATE egov_action_request_forms SET "
      sSQL = sSQL & sDBColumnName & " = " & sNewValue_isInternal
      sSQL = sSQL & " WHERE action_form_id = " & sFormID

     	set oMaintainFormOptions = Server.CreateObject("ADODB.Recordset")
    	 oMaintainFormOptions.Open sSQL, Application("DSN"), 3, 1

      set oMaintainFormOptions = nothing

      lcl_success = "SU"
   else
      lcl_success   = "ERROR"
      lcl_isAjaxMsg = "ERROR"
   end if

   response.write lcl_success

end sub

'------------------------------------------------------------------------------
function getNewFormValue(iTask, _
                         iFormID)

  dim lcl_return, sSQL, sTask, sFormID
  dim sDBColumn

  sDBColumn  = ""
  sTask      = ""
  sFormID    = 0

  if trim(iTask) <> "" then
     sTask = ucase(iTask)
  end if

  if iFormID <> "" then
     if not containsApostrophe(iFormID) then
        sFormID = clng(iFormID)
     end if
  end if

  if sTask = "INTERNAL" then
     sDBColumn = "isnull(action_form_internal, 0)"
  elseif sTask = "ENABLEFORM" then
     sDBColumn = "isnull(action_form_enabled, 0)"
  elseif sTask = "DISPLAYONLIST" then
     sDBColumn = "isnull(action_form_displayOnList, 0)"
  end if

  sSQL = "SELECT " & sDBColumn & " as dbColumnValue "
  sSQL = sSQL & " FROM egov_action_request_forms "
  sSQL = sSQL & " WHERE action_form_id = " & sFormID

 	set oGetNewFormValue = Server.CreateObject("ADODB.Recordset")
	 oGetNewFormValue.Open sSQL, Application("DSN"), 3, 1

  if not oGetNewFormValue.eof then
     if oGetNewFormValue("dbColumnValue") <> "" then

       'This is the current value
        if sTask = "INTERNAL" OR sTask = "ENABLEFORM" OR sTask = "DISPLAYONLIST" then
           lcl_return = "1"

          'We need to SWITCH the value to FALSE
           if oGetNewFormValue("dbColumnValue") then
              lcl_return = "0"
           end if
        end if
     end if
  end if

  oGetNewFormValue.close
  set oGetNewFormValue = nothing

  getNewFormValue = lcl_return

end function

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  if p_value <> "" then
     sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"
    	set oDTB = Server.CreateObject("ADODB.Recordset")
   	 oDTB.Open sSQL, Application("DSN"), 3, 1
  end if

end sub
%>
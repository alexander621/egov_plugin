<%
'------------------------------------------------------------------------------
function setupScreenMsg(iSuccess)

  lcl_return = ""

  if iSuccess <> "" then
     iSuccess = UCASE(iSuccess)

     if iSuccess = "SU" then
        lcl_return = "Successfully Updated..."
     elseif iSuccess = "SA" then
        lcl_return = "Successfully Created..."
     elseif iSuccess = "SD" then
        lcl_return = "Successfully Deleted..."
    elseif iSuccess = "RSS_SUCCESS" then
       lcl_msg = "Successfully Sent to RSS..."
    elseif iSuccess = "RSS_ERROR" then
       lcl_msg = "ERROR: Failed to send to RSS..."
    elseif iSuccess = "AJAX_ERROR" then
       lcl_msg = "ERROR: An error has during the AJAX routine..."

     end if
  end if

  setupScreenMsg = lcl_return

end function

'------------------------------------------------------------------------------
sub dtb_debug(p_value)
  sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"
  runsql(sSQL)
end sub
%>
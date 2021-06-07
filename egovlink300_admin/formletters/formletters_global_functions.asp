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
     elseif iSuccess = "MOVE_UP" then
        lcl_return = "Successfully Moved Up..."
     elseif iSuccess = "MOVE_DOWN" then
        lcl_return = "Successfully Moved Down..."
     elseif iSuccess = "MOVE_TOP" then
        lcl_return = "Successfully Moved to the Top..."
     elseif iSuccess = "MOVE_BOTTOM" then
        lcl_return = "Successfully Moved to the Bottom..."
     elseif iSuccess = "RSS_SUCCESS" then
        lcl_return = "Successfully Sent to RSS..."
     elseif iSuccess = "RSS_ERROR" then
        lcl_return = "ERROR: Failed to send to RSS..."
     elseif iSuccess = "AJAX_ERROR" then
        lcl_return = "ERROR: An error has during the AJAX routine..."
     end if
  end if

  setupScreenMsg = lcl_return

end function
%>
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
     elseif iSuccess = "EU" then
        lcl_return = "This Bid Posting has sub-categories associated to it.  Category cannot be modified."
     elseif iSuccess = "SNJ" then
        lcl_return = "Job Posting Successfully Created..."
     elseif iSuccess = "SNB" then
       lcl_return = "Bid Posting Successfully Created..."
     end if
  end if

  setupScreenMsg = lcl_return

end function
%>
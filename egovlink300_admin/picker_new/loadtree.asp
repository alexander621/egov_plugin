<%Response.buffer = true%>
<!-- #include file="URLDecode.asp" //-->
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="picker_global_functions.asp" //-->
<%
  strPath            = URLDecode(request("path"))
  session("curpath") = strPath

 'Get City Document Location
  sLocationName = trim(GetVirtualName(Session("OrgID")))

 'Display Folder Tree
  strList = LoadFolder("/public_documents300/custom/pub/" & sLocationName,strPath)
  sFolder = Right(strPath, Len(strPath) - InStrRev(strPath,"/"))

  if sFolder = "pub" then
     sFolder = "Documents"
  end if

  if Left(sFolder,2) = "z." then
     sFolder = Mid(sFolder, 3)
  end if

 'Determine which sections to display
  lcl_displayDocuments  = "N"
  lcl_displayActionLine = "N"
  lcl_displayPayments   = "N"
  lcl_displayURL        = "N"

  if request("displayDocuments") = "Y" then
     lcl_displayDocuments = "Y"
  end if

  if request("displayActionLine") = "Y" then
     lcl_displayActionLine = "Y"
  end if

  if request("displayPayments") = "Y" then
     lcl_displayPayments = "Y"
  end if

  if request("displayURL") = "Y" then
     lcl_displayURL = "Y"
  end if

 'Build the BODY onload
  lcl_onload = ""

  if lcl_displayActionLine = "Y" then
     lcl_onload = lcl_onload & "parent.document.frmAddArticle.currentfolderpath.value='" & strPath & "';"
  end if

  lcl_onload = lcl_onload & "parent.document.all.currentfolder.value='" & sFolder & "';"
  lcl_onload = lcl_onload & "parent.document.all.currentpath.value='"   & strPath & "';"

 'Build the page
  response.write "<html>" & vbcrlf
  response.write "<head>" & vbcrlf
  response.write "  <link rel=""stylesheet"" type=""text/css"" href=""menu.css"" />" & vbcrlf
  response.write "</head>" & vbcrlf
  response.write "<body leftmargin=""0"" topmargin=""0"" onload=""" & lcl_onload & """>" & vbcrlf
  response.write "  <div id=""menu2"">" & vbcrlf
  response.write "    <table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"" id=""menu"">" & vbcrlf
  response.write "      <tr>" & vbcrlf
  response.write "          <td valign=""top"" width=""100%"">" & vbcrlf
  response.write "              <nobr>" & vbcrlf
  response.write "              <ul id=""ulRoot"" style=""display: block;"">" & vbcrlf

  if strList = "" then
     response.write "&nbsp;<font color=""#003366""><i>Permission Denied</i></font>" & vbcrlf
  else
     response.write strList & vbcrlf
  end if

  response.write "              </ul>" & vbcrlf
  response.write "              </nobr>" & vbcrlf
  response.write "          </td>" & vbcrlf
  response.write "      </tr>" & vbcrlf
  response.write "    </table>" & vbcrlf
  response.write "    <iframe name=""hiddenframe"" src="""" width=""0"" height=""0""></iframe>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf
%>
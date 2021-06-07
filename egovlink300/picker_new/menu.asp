<!-- #include file="../include_top_functions.asp" //-->
<!-- #include file="../class/classOrganization.asp" //-->
<!-- #include file="picker_global_functions.asp" //-->
<%
  Dim oPickerMenu

  set oPickerMenu = New classOrganization

 'Get City Document Location
  sLocationName = trim(GetVirtualName(iorgid))

 'Get the starting folder
  lcl_folderStart = ""

  if request("folderStart") <> "" then
     lcl_folderStart = request("folderStart")
  'else
  '   lcl_folderStart = "published_documents"
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

 'Determine if we default the URL to open
  lcl_onload = ""

  if  lcl_displayDocuments  = "N" _
  and lcl_displayActionLine = "N" _
  and lcl_displayPayments   = "N" _
  and lcl_displayURL        = "Y" then
      lcl_onload = lcl_onload & "document.getElementById('newurl').click();"
  end if
%>
<html>
<head>
  <style type="text/css">
  <!--
    td      { font-family: MS Sans Serif,Tahoma,Arial; font-size:10px; color:#ffffff;}
    .sel    { border: 1px solid #cccccc; cursor:hand; padding:5px 0px;}
    .notsel { border: 1px solid #666666; cursor:hand; padding:5px 0px;}
  //-->
  </style>
<script language="javascript">
<!--
function MakeActive(o) {
<%
  if lcl_displayDocuments = "Y" then
     response.write "  document.all.exdoc.className  = ""notsel"";" & vbcrlf
  end if

  if lcl_displayActionLine = "Y" then
     response.write "  document.all.newdoc.className = ""notsel"";" & vbcrlf
  end if

  if lcl_displayPayments = "Y" then
     response.write "  document.all.newpay.className = ""notsel"";" & vbcrlf
  end if

  if lcl_displayURL = "Y" then
     response.write "  document.all.newurl.className = ""notsel"";" & vbcrlf
  end if
%>
  o.className = "sel";
  parent.MakeActive(o.id);
}

function clickMenuOption(iID,iOption) {
  MakeActive(iOption);

  if(iID == "exdoc") {
  <%
    lcl_folder_location = sLocationName

    if lcl_folderStart <> "CITY_ROOT" then
       lcl_folder_location = lcl_folder_location & lcl_folderStart
    end if
       
    response.write "parent.explorer.window.location.href='loadtree.asp?path=/public_documents300/custom/pub/" & lcl_folder_location & "';"
  %>
     //parent.explorer.window.location.href='loadtree.asp?path=/public_documents300/custom/pub/<%=sLocationName%><%=lcl_folderStart%>';
  }else if(iID == "newdoc") {
     parent.explorer.window.location.href='listforms.asp';
  }else if(iID == "newpay") {
     parent.explorer.window.location.href='listpayments.asp';
  }
}
//-->
</script>
</head>
<body bgcolor="#666666" topmargin="0" leftmargin="0" onload="<%=lcl_onload%>">
<%
  response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""85"">" & vbcrlf

 'Documents
  if lcl_displayDocuments = "Y" then
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""center"">" & vbcrlf
     response.write "          <div class=""sel"" id=""exdoc"" onclick=""clickMenuOption('exdoc',this);"">" & vbcrlf
     response.write "          <img src=""images/existdoc.gif"" /><br />Link Document" & vbcrlf
     response.write "          </div>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

 'Action Line
  if lcl_displayActionLine = "Y" then
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""center"">" & vbcrlf
     response.write "          <div class=""notsel"" id=""newdoc"" onclick=""clickMenuOption('newdoc',this);"">" & vbcrlf
     response.write "          <img src=""images/newdoc.gif"" /><br />Link Action Form" & vbcrlf
     response.write "          </div>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

 'Payments
  if lcl_displayPayments = "Y" then
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""center"">" & vbcrlf
     response.write "          <div class=""notsel"" id=""newpay"" onclick=""clickMenuOption('newpay',this);"">" & vbcrlf
     response.write "          <img src=""images/newdoc.gif"" /><br />Link Payment Form" & vbcrlf
     response.write "          </div>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

 'URL
  if lcl_displayURL = "Y" then
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""center"">" & vbcrlf
     response.write "          <div class=""notsel"" id=""newurl"" onclick=""clickMenuOption('newurl',this);"">" & vbcrlf
     response.write "          <img src=""images/newurl.gif"" /><br />Link to URL<br /><br />" & vbcrlf
     response.write "          </div>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

  response.write "</table>" & vbcrlf
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf
%>
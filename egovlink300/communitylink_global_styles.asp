<%
'------------------------------------------------------------------------------
sub buildStyles_CL_Misc()

  response.write "  .orgLogo { padding: 5px; }" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub buildStyles_CL_TopBar(p_topbar_bgcolor, p_topbar_fontcolor, p_topbar_fonttype, p_topbar_fontcolorhover)

  response.write "  .topBar td " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "      padding:          5px;" & vbcrlf
  response.write "      background-color: #" & p_topbar_bgcolor   & ";" & vbcrlf
  response.write "      color:            #" & p_topbar_fontcolor & ";" & vbcrlf
  response.write "    }" & vbcrlf

  response.write "  .topBarOption:link, .topBarOption:visited " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "      font-family: "  & p_topbar_fonttype  & ";" & vbcrlf
  response.write "      color:       #" & p_topbar_fontcolor & ";" & vbcrlf
  response.write "      font-size:   10px;" & vbcrlf
  response.write "    }" & vbcrlf

  response.write "  .topBarOption:hover " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "      font-family:      " & p_topbar_fonttype       & ";" & vbcrlf
  response.write "      color:           #" & p_topbar_fontcolorhover & ";" & vbcrlf
  response.write "      font-size:       10px;" & vbcrlf
  response.write "      text-decoration: underline;" & vbcrlf
  response.write "    }" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub buildStyles_CL_Footer(p_footer_bgcolor, p_footer_fontcolor, p_footer_fonttype, p_footer_fontcolorhover)

  response.write "  .footer td " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "      padding:          5px;" & vbcrlf
  response.write "      background-color: #" & p_footer_bgcolor   & ";" & vbcrlf
  response.write "      color:            #" & p_footer_fontcolor & ";" & vbcrlf
  response.write "    }" & vbcrlf

  response.write "  .footerOption:link, .footerOption:visited " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "      font-family: "  & p_footer_fonttype  & ";" & vbcrlf
  response.write "      color:       #" & p_footer_fontcolor & ";" & vbcrlf
  response.write "      font-size:   10px;" & vbcrlf
  response.write "    }" & vbcrlf

  response.write "  .footerOption:hover " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "      font-family:      " & p_footer_fonttype       & ";" & vbcrlf
  response.write "      color:           #" & p_footer_fontcolorhover & ";" & vbcrlf
  response.write "      font-size:       10px;" & vbcrlf
  response.write "      text-decoration: underline;" & vbcrlf
  response.write "    }" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub buildStyles_CL_SideMenuBar(p_blnMenuOn, p_showsidemenubar, p_sidemenuoption_bgcolor, p_sidemenuoption_fonttype, _
                               p_sidemenuoption_fontcolor, p_sidemenuoption_fontcolorhover)

  if p_blnMenuOn AND p_showsidemenubar then
     response.write "  .sideMenuBar" & vbcrlf
     response.write "    { " & vbcrlf
     response.write "      padding:          5px;" & vbcrlf
     response.write "      cursor:           pointer;" & vbcrlf
     response.write "      border-bottom:    1pt solid #ffffff;" & vbcrlf
     response.write "      background-color: #" & p_sidemenuoption_bgcolor & ";" & vbcrlf
     'response.write "      color: #" & p_topbarfontcolor & ";" & vbcrlf
     response.write "    }" & vbcrlf

     response.write "  .sideMenuBarOption:link, .sideMenuBarOption:visited " & vbcrlf
     response.write "    { " & vbcrlf
     response.write "      font-size:   12px;" & vbcrlf
     response.write "      font-family:  " & p_sidemenuoption_fonttype  & ";" & vbcrlf
     response.write "      color:       #" & p_sidemenuoption_fontcolor & ";" & vbcrlf
     response.write "    }" & vbcrlf

     response.write "  .sideMenuBarOption:hover " & vbcrlf
     response.write "    { " & vbcrlf
     response.write "      font-size: 12px;" & vbcrlf
     response.write "      font-family:  " & p_sidemenuoption_fonttype       & ";" & vbcrlf
     response.write "      color:       #" & p_sidemenuoption_fontcolorhover & ";" & vbcrlf
     response.write "      text-decoration: underline;" & vbcrlf
     response.write "    }" & vbcrlf
  end if

end sub

'------------------------------------------------------------------------------
sub buildStyles_CL_PageHeader(p_pageheader_bgcolor, p_pageheader_fonttype, p_pageheader_fontcolor, p_pageheader_fontsize)

  response.write "  .pageHeader "
  response.write "    { " & vbcrlf
  response.write "     	font-size:         " & p_pageheader_fontsize & "px; " & vbcrlf
  response.write "      background-color: #" & p_pageheader_bgcolor  & "; " & vbcrlf
  response.write "      border-bottom: 1pt solid #808080;" & vbcrlf
  response.write "      padding: 5px;" & vbcrlf
  response.write "    } " & vbcrlf

  response.write "  .pageHeader_welcome " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "     	font-family:    " & p_pageheader_fonttype  & "; " & vbcrlf
  response.write "     	color:         #" & p_pageheader_fontcolor & "; " & vbcrlf
  response.write "     	font-size:     12px; " & vbcrlf
  response.write "     	font-weight:   bold; " & vbcrlf
  response.write "     	padding:       0; " & vbcrlf
  response.write "     	margin-bottom: 2px; " & vbcrlf
  response.write "    } " & vbcrlf

  response.write "  .pageHeader_welcomeSubMsg " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "     	font-family:  " & p_pageheader_fonttype  & "; " & vbcrlf
  response.write "     	color:       #" & p_pageheader_fontcolor & "; " & vbcrlf
  response.write "      font-size:   10px; " & vbcrlf
  response.write "      font-weight: bold; " & vbcrlf
  response.write "    } " & vbcrlf

  response.write "  .pageHeader_homePageMsg " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "     	font-family:  " & p_pageheader_fonttype  & "; " & vbcrlf
  response.write "     	color:       #" & p_pageheader_fontcolor & "; " & vbcrlf
  response.write "     	font-size:    " & p_pageheader_fontsize  & "px; " & vbcrlf
  response.write "      padding-top: 5px; " & vbcrlf
  response.write "    } " & vbcrlf

end sub

'------------------------------------------------------------------------------
sub buildStyles_CL_LinkOptions()

  response.write "  .communitylink_table " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "      border: 1pt solid #000000; " & vbcrlf
  response.write "    } " & vbcrlf

  response.write "  .communitylink_table th " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "      font-weight:      bold; " & vbcrlf
  response.write "      border-bottom:    1pt solid #000000; " & vbcrlf
  response.write "      background-color: #cccccc; " & vbcrlf
  response.write "    } " & vbcrlf

  response.write "  .communityLink_viewLinks:link, .communityLink_viewLinks:visited, .communityLink_viewLinks:active " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "      color: #800000; " & vbcrlf
  response.write "    } " & vbcrlf

  response.write "  .communityLink_viewLinks:hover " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "      color: #800000; " & vbcrlf
  response.write "      text-decoration: underline; " & vbcrlf
  response.write "    } " & vbcrlf

  response.write "  .communityLink_bottomborder " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "      border-bottom: 1pt dotted #808080; "  & vbcrlf
  response.write "    } " & vbcrlf

end sub
%>
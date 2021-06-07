<%
'------------------------------------------------------------------------------
function LoadFolder(vpath,dbpath)
  Dim sSql, oRst
  Dim objFSO, objDir, objFound, url, strNav, temp, pos, fullPath, blnFound, newPath, name, icon, bCanSeeSecure

 'Only tranlate to a local path if a virtual path is specified
  if InStr(1, path, ":") < 1 Then
     fullPath = Server.MapPath(vpath)
  else
     fullPath = vpath
  end if

 'Determine Physical Path
  ppath    = Server.MapPath(dbpath)
  ppath    = replace(ppath,"\custom\pub\custom\pub","\custom\pub")
  fullpath = ppath

  'bCanSeeSecure = True
  bCanSeeSecure = HasPermission("CanEditDocuments")

  set objFSO = Server.CreateObject("Scripting.FileSystemObject")

  blnFound = objFSO.FolderExists(fullPath)
  if blnFound then
     set objDir = objFSO.GetFolder(fullPath)
  end if

  strNav = ""

  if IsObject(objDir) AND Err.Number = 0 then
     sSQL = "EXEC ListFolders " & Session("OrgID") & ", " & Session("UserID") & ", '" & dbpath & "'"

     set oRst = Server.CreateObject("ADODB.Recordset")
     oRst.Open sSQL, Application("DSN"), 3, 1

     if not oRst.eof then
        do while not oRst.eof
           newpath = Server.URLEncode(oRst("FolderPath"))
           name = oRst("FolderName")

           if Left(name,1) <> "_" then  'if dir starts with underscore that means it is hidden
              if Left(name,2) <> "z." then 'if dir starts with z. that means it at the bottom of the list (sorting option)
                 if oRst("Secure") > 0 And bCanSeeSecure then
                    icon = "images/locked_folder_closed.gif"
                 else
                    icon = "images/folder_closed.gif"
                 end if
              else
                 name = Right(name,Len(name)-2)
                 icon = "images/help.gif"
              end if

              strNav = strNav & "<li style=""position:relative;"" nodeType=""c"" id=""" & newpath & """ class=""kid"">" & vbcrlf
              strNav = strNav & "  <img style=""filter:alpha(opacity=100);"" src=""" & icon &""" width=""18"" height=""18"" align=""absmiddle""> " & vbcrlf
              strNav = strNav & "  <a title=""" & name & """ href=""loadtree.asp?path="& newPath & """ target=""explorer"">" & name & "</a>" & vbcrlf
              strNav = strNav & "  <ul class=""hdn""></ul>" & vbcrlf
              strNav = strNav & "</li>" & vbcrlf
              'strNav = strNav & "<LI style=""position:relative;"" nodeType=""c"" id=""" & newpath & """ class=kid><IMG style=""filter:alpha(opacity=100);"" src=""" & icon &""" width=""18"" height=""18"" align=""absmiddle""> <A title=""" & name & """ href=""loadtree.asp?path="& newPath & """ target=""hiddenframe"">" & name & "</A><UL class=hdn></UL></LI>" & vbCrLf
           end if

           oRst.movenext
        loop

        oRst.Close
     end if

     set oRst = nothing

     for each objFound In objDir.Files
        name = objFound.Name

        url = dbpath & "/" & name
        url = Replace(url, " ", "%20")
        'url = replace(url,"/custom/pub","")

        imgSrc = "images/document.gif"

        pos = InStr(1, name, ".")

        if pos > 0 then
           select Case Mid(name,pos+1,3)
             case "doc"
                imgSrc = "images/msword.gif"
             case "xls"
                imgSrc = "images/msexcel.gif"
             case "ppt"
                imgSrc = "images/msppt.gif"
             case "htm"
                imgSrc = "images/msie.gif"
             case "pdf"
                imgSrc = "images/pdf.gif"
           end select

           temp = Left(name,pos-1)
        else
           temp = name
        end if

       'Code for Files
        sLaunchURL = replace(url,"custom/pub","")
        'strNav = strNav & "<LI style=""position:relative; filter: Alpha(opacity=50);"" nodeType=""a"" id=""" & url & """><IMG class=clsNoHand style=""filter:alpha(opacity=100);"" src=""" & imgSrc & """ width=""18"" height=""18"" align=""absmiddle""> <A title=""" & temp & """ target=""fraTopic"" href=""" & sLaunchURL & """>" & temp & "</A></LI>"
         strNav = strNav & "<li style=""position:relative; filter: Alpha(opacity=50);"" nodeType=""a"" id=""" & url & """>" & vbcrlf
         strNav = strNav & "  <img class=""clsNoHand"" style=""filter:alpha(opacity=100);"" src=""" & imgSrc & """ width=""16"" height=""16"" align=""absmiddle"" /> " & vbcrlf
         strNav = strNav & "  <a title=""" & temp & """ href=""#"" onclick=""parent.document.all.FilePath.value='" & name & "'"">" & temp & "</a>" & vbcrlf
         strNav = strNav & "</li>" & vbcrlf
     Next

     set objDir = nothing
     set objFSO = nothing

     if strNav = "" then
        strNav = "<li>" & vbcrlf
        strNav = strNav & "  <font color=""#003366"">" & vbcrlf
        strNav = strNav & "  <img src=""images/spacer.gif"" width=""0"" height=""18"" align=""absmiddle"" />" & vbcrlf
        strNav = strNav & "  &nbsp;&nbsp;&nbsp;&nbsp;<i>(Empty)</i></font>" & vbcrlf
        strNav = strNav & "</li>" & vbcrlf
        'If strNav = "" Then strNav = "<li style=""margin-top:-13px;""></li>"
     end if

     LoadFolder = strNav
  else
     set objDir = nothing
     set objFSO = nothing

     response.write "Error Occured."
  end if

end function

'------------------------------------------------------------------------------
function GetVirtualName(iorgid)
  
  lcl_return = "UNKNOWN"

  if iorgid <> "" then
     sSQL = "SELECT OrgVirtualSiteName "
     sSQL = sSQL & " FROM Organizations "
     sSQL = sSQL & " WHERE orgid = '" &  iorgid & "'"

     set oVirtualName = Server.CreateObject("ADODB.Recordset")
     oVirtualName.open sSQL,Application("DSN"),3,1
  
     if not oVirtualName.eof then
        lcl_return = oVirtualName("OrgVirtualSiteName")
     end if
  end if

  GetVirtualName = lcl_return

end function

'------------------------------------------------------------------------------
sub dtb_debug(p_value)
  sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"

  set oDTB = Server.CreateObject("ADODB.Recordset")
  oDTB.open sSQL,Application("DSN"),3,1

  set oDTB = nothing

end sub
%>
<!-- #include file="../includes/common.asp" -->
<%
Response.Buffer = True
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

response.write Session("curpath") & " is current path.<br>"
response.write request("foldername") & " is the current folder name."

strDir = Session("curpath") & "/" &request("foldername")
%>

<%

    If Not objFSO.FolderExists(Server.MapPath(strDir)) Then
      objFSO.CreateFolder(Server.MapPath(strDir))

      '---BEGIN: Update DB fields--------------------------------
      Set oCmd = Server.CreateObject("ADODB.Command")
      With oCmd
      .ActiveConnection = Application("DSN")
      .CommandText = "NewFolder"
      .CommandType = adCmdStoredProc
      .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
      .Parameters.Append oCmd.CreateParameter("CreatorID", adInteger, adParamInput, 4, Session("UserID"))
      .Parameters.Append oCmd.CreateParameter("FolderPath", adVarChar, adParamInput, 300, strDir)
      .Execute
      End With
      Set oCmd = Nothing
      '---END: Update DB fields----------------------------------

       Response.Write "<img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='main.asp' style='font-family: Verdana, Arial, Helvetica; font-weight: bold; font-size: 8pt;' >Back To Documents Main</a><br><br>"
      Response.Write "<b>Folder was added successfully.</b>"
	  Response.redirect("picker/default.asp?task=return&path="&Server.URLEncode(Session("curpath")))

    Else
	   Response.Write "<img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='main.asp' style='font-family: Verdana, Arial, Helvetica; font-weight: bold; font-size: 8pt;' >Back To Documents Main</a><br><br>"
      Session("ECapture_Error") = "The category you attempted to add already exists.<br>Please try again."
     
    End If
%>
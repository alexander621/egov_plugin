<!-- #include file="../includes/common.asp" //-->
<%
'if not logged in dont allow access, no matter what
If Session("UserID") = 0 Then Response.Redirect RootPath

Dim oCmd, sTo, sCC, sGroups, sUsers, sUsersCC, sGroupsCC, index, nextSemi, sSql, oRst, sSent, sNotFound, iSendNow, oldID
index=1
oldID=Request.Form("OldMailID")
if oldID & "" = "" then oldID=0
Response.Write("---" & oldID & "---")

sTo=Trim(Request.Form("To"))
sCC=Trim(Request.Form("CC"))

iSendNow = 0
If Request.Form("SendNow") & "" = "1" Then
  iSendNow = 1
End If

'------------------------------------------------------------------------------------
'Begin Formatting for the To field
'------------------------------------------------------------------------------------
do while (instr(index,sTo,";")>0)
  nextSemi=instr(index,sTo,";")
  AddToListTo trim(mid(sTo, index, nextSemi-index))
  index=nextSemi+1
loop
if len(mid(sTo,index))>0 then AddToListTo trim(mid(sTo,index))
'if len(sGroups) > 0 then sGroups=Left(sGroups,len(sGroups)-1)
'if len(sUsers) > 0 then sUsers=Left(sUsers,len(sUsers)-1)

public sub AddToListTo(sName)
  dim sFirst, sLast
  if instr(1, sName, ",") >0 then 'must be user
    sFirst=trim(mid(sName,1,instr(1, sName, ",")-1))
    sLast=trim(mid(sName,instr(1, sName, ",")+1))
    sUsers=sUsers+sFirst+ "_" +sLast + ","
  else 'must be group
    sGroups=sGroups+sName +","
  end if
end sub
'------------------------------------------------------------------------------------
'End Formatting for the To field
'------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------
'Begin Formatting for the CC field
'------------------------------------------------------------------------------------
index=1
do while (instr(index,sCC,";")>0)
  nextSemi=instr(index,sCC,";")
  AddToListCC trim(mid(sCC, index, nextSemi-index))
  index=nextSemi+1
loop
if len(mid(sCC,index))>0 then AddToListCC trim(mid(sCC,index))

public sub AddToListCC(sName)
  dim sFirst, sLast
  if instr(1, sName, ",") >0 then 'must be user
    sFirst=trim(mid(sName,1,instr(1, sName, ",")-1))
    sLast=trim(mid(sName,instr(1, sName, ",")+1))
    sUsersCC=sUsersCC+sFirst+ "_" +sLast + ","
  else 'must be group
    sGroupsCC=sGroupsCC+sName +","
  end if
end sub
'------------------------------------------------------------------------------------
'End Formatting for the CC field
'------------------------------------------------------------------------------------


Response.Write("{" & Request.Form("To") & "} <br>")
Response.Write("{" & sUsers & "} <br>")
Response.Write("{" & sGroups & "} <br>")
Response.Write("{" & sGroupsCC & "} <br>")

'sSql = "EXEC SendMailUser " & Session("OrgID") & "," & Session("UserID") & ",'" & sUsers  & "','" & sUsersCC & "','" & sGroups  & "','" & sGroupsCC & "','" & Replace(Request.Form("Subject"),"'", "") & "'" & ",'" & Replace(Request.Form("Message"),"'", "") & "'" & "," & iSendNow
'Response.Write(sSql)

Set oCmd = Server.CreateObject("ADODB.Command")
With oCmd
  .ActiveConnection = Application("DSN")
  .CommandText = "SendMailUser"
  .CommandType = adCmdStoredProc
  .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID"))
  .Parameters.Append oCmd.CreateParameter("UserID", adInteger, adParamInput, 4, Session("UserID"))
  .Parameters.Append oCmd.CreateParameter("sTo", adVarChar, adParamInput, 2000, sUsers)
  .Parameters.Append oCmd.CreateParameter("sCc", adVarChar, adParamInput, 2000, sUsersCC)
  .Parameters.Append oCmd.CreateParameter("sGroupsTo", adVarChar, adParamInput, 2000, sUsers)
  .Parameters.Append oCmd.CreateParameter("sGroupsCc", adVarChar, adParamInput, 2000, sUsersCC)
  .Parameters.Append oCmd.CreateParameter("Subject", adVarChar, adParamInput, 100, Request.Form("Subject"))
  .Parameters.Append oCmd.CreateParameter("Message", adVarChar, adParamInput, 5000, Request.Form("Message"))
  .Parameters.Append oCmd.CreateParameter("SendNow", adInteger, adParamInput, 4, Request.Form("SendNow"))
  .Parameters.Append oCmd.CreateParameter("OldMailID", adInteger, adParamInput, 4, oldID)
End With

Set oRst = Server.CreateObject("ADODB.Recordset")
With oRst
  .CursorLocation = adUseClient
  .CursorType = adOpenStatic
  .LockType = adLockReadOnly
  .Open oCmd
End With

Set oCmd = Nothing

Do while not oRst.EOF
  if oRst("uID") = -1 then
    sNotFound=sNotFound & oRst("LastName") & ", " & oRst("FirstName") & "; "
  else
    sSent=sSent & oRst("LastName") & ", " & oRst("FirstName") & "; "
  end if
  oRst.MoveNext()
Loop

if sSent <> "" then
  Response.Write("The message was successfully sent to the following people:" & "<br>" & sSent & "<br>")
end if

if len(sNotFound) > 0 then 
  Response.Write("The following people could not be found:" & "<br>" & sNotFound)
end if

if oRst.state=1 then oRst.Close


If iSendNow = 1 Then
  Response.Redirect "sentmail.asp?success"
Else
  Response.Redirect "drafts.asp"
End If
%>
<!-- #include file="../includes/common.asp" //-->
<%
On Error Resume Next

'INITIATILIZE OBJECTS AND VARIABLES
Dim sDelete, oCmd, Item, strIDList


'LOOP THRU EACH ITEM IN FORM TO DETERMINE IF IT IS AN ANNOTATION TO DELETE
For Each Item In Request.Form
  If Left(Item,4) = "del_" Then
   	strIDList = strIDList & Mid(Item,5) & ","
  End If
Next

'DEBGUG CODE: response.write strIDList


'CALL PROCEDURE TO DELETE ANNOTATIONS
Set oCmd = Server.CreateObject("ADODB.Command")
With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "DelAnnotations"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("AnnotationIDs", adVarChar, adParamInput, 255, strIDList)
    .Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, Session("OrgID") )
    .Execute
End With
  

' DESTROY OBJECTS
Set oCmd = Nothing
  

' REDIRECT TO ANNOTATION PAGE WITH MESSAGE
Response.Redirect "annotatearticle.asp?task=deleted"
%>
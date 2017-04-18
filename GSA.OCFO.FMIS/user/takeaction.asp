<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/fmisuservb.asp" -->
<%
Set takeaction = Server.CreateObject("ADODB.Recordset")
takeaction.ActiveConnection = MM_fmisuservb_STRING

If Request.QueryString("method") = "mgmt" Then
takeaction.Source = "UPDATE users SET mgmt_approver = '" &Request.QueryString("mgmt_approver")& "', mgmt_app_date = '" &Request.QueryString("mgmt_app_date")& "' WHERE ref_num = " &Request.QueryString("ref_num")
takeaction.CursorType = 0
takeaction.CursorLocation = 2
takeaction.LockType = 1
takeaction.Open()
takeaction_numRows = 0
Response.Redirect("newrequests.asp")
End If

If Request.QueryString("method") = "sec" Then
takeaction.Source = "UPDATE users SET sec_approver = '" &Request.QueryString("sec_approver")& "', sec_app_date = '" &Request.QueryString("sec_app_date")& "' WHERE ref_num = " &Request.QueryString("ref_num")
takeaction.CursorType = 0
takeaction.CursorLocation = 2
takeaction.LockType = 1
takeaction.Open()
takeaction_numRows = 0
Response.Redirect("newrequests.asp")
End If

If Request.QueryString("method") = "userid" Then
takeaction.Source = "UPDATE users SET userid = '" &Request.QueryString("userid")& "' WHERE ref_num = " &Request.QueryString("ref_num")
takeaction.CursorType = 0
takeaction.CursorLocation = 2
takeaction.LockType = 1
takeaction.Open()
takeaction_numRows = 0
Response.Redirect "ex_singleuser.asp?ref_num=" &Request.QueryString("ref_num")
End If
%>

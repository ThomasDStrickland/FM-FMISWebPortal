<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%> 
<!--#include file="../Connections/fmisuservb.asp" -->
<%
Dim action
Dim action_numRows

Set action = Server.CreateObject("ADODB.Recordset")
action.ActiveConnection = MM_fmisuservb_STRING
action.Source = "SELECT * FROM users"
action.CursorType = 0
action.CursorLocation = 2
action.LockType = 1
action.Open()

action_numRows = 0
%>
<html>
<head>
<title>Management Actions</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="fmis.css" rel="stylesheet" type="text/css">
</head>
<body>
<p>&nbsp;</p>

  
<form action="takeaction.asp" method="get" name="form1">
  <table width="80%" align="left">
    <tr valign="baseline"> 
      <td width="39%" align="right" nowrap><div align="left">User Requesting ID 
          :</div></td>
      <td width="61%"><%=(action.Fields.Item("user").Value)%> <input name="method" type="hidden" id="method" value="<%=Request.QueryString("method")%>">
        <input name="ref_num" type="hidden" id="ref_num" value="<%=(action.Fields.Item("ref_num").Value)%>">
      </td>
    </tr>
    <%
	If Request.QueryString("method") = "mgmt" Then
	%>
    <tr valign="baseline"> 
      <td nowrap align="right"><div align="left">BCAC Approving Manager :</div></td>
      <td> <input type="text" name="mgmt_approver" size="32"></td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"><div align="left">BCAC Management Approval Date 
          :</div></td>
      <td> <input type="text" name="mgmt_app_date" value="<%=Date()%>" size="32"> 
      </td>
    </tr>
    <%
	End If 
	If Request.QueryString("method") = "sec" Then
	%>
    <tr valign="baseline"> 
      <td nowrap align="right"><div align="left">Security Manager Approver :</div></td>
      <td> <input type="text" name="sec_approver" size="32"> </td>
    </tr>
    <tr valign="baseline"> 
      <td nowrap align="right"><div align="left">Security Managerment Approval 
          Date :</div></td>
      <td> <input name="sec_app_date" type="text" value="<%=Date()%>" size="32"> 
      </td>
    </tr>
    <%
	End If
	If Request.QueryString("method") = "userid" Then
	%>
    <tr valign="baseline"> 
      <td nowrap align="right"><div align="left">Userid Assigned :</div></td>
      <td> <input type="text" name="userid" value="<%=(action.Fields.Item("userid").Value)%>" size="32"></td>
    </tr>
    <%
	End If
	%>
    <tr valign="baseline"> 
      <td nowrap align="right">&nbsp;</td>
      <td> <input type="submit" value="Update Record"> </td>
    </tr>
  </table>
</form>
<p>&nbsp;</p>
</body>
</html>
<%
action.Close()
Set action = Nothing
%>

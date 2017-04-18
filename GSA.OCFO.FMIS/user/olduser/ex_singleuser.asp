<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/fmisuservb.asp" -->
<%
Dim userinfo__MMColParam
userinfo__MMColParam = "1"
If (Request.QueryString("ref_num") <> "") Then 
  userinfo__MMColParam = Request.QueryString("ref_num")
End If
%>
<%
Dim userinfo
Dim userinfo_numRows

Set userinfo = Server.CreateObject("ADODB.Recordset")
userinfo.ActiveConnection = MM_fmisuservb_STRING
userinfo.Source = "SELECT *  FROM users  WHERE ref_num = " + Replace(userinfo__MMColParam, "'", "''") + ""
userinfo.CursorType = 0
userinfo.CursorLocation = 2
userinfo.LockType = 1
userinfo.Open()

userinfo_numRows = 0
%>
<html>
<head>
<title>UserID Request Form : Printer Friendly</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="fmis.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="650" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td><p class="headtext"><%=(userinfo.Fields.Item("system").Value)%> UserID Request Form</p>
      <p>&nbsp;</p></td>
  </tr>
  <tr> 
    <td><span class="headtext">Date ID Requested</span> <strong>:</strong> <%=(userinfo.Fields.Item("date_requested").Value)%></td>
  </tr>
  <tr>
    <td><span class="headtext">Agency</span> <strong>:</strong> <%=(userinfo.Fields.Item("ex_agency").Value)%></td>
  </tr>
  <tr> 
    <td> <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <caption align="top" class="headtext">
        Approving Director / Manager Information 
        </caption>
        <tr> 
          <th width="24%" scope="row"> 
            <div align="right">Name : </div></th>
          <td><%=(userinfo.Fields.Item("mgrname").Value)%> </td>
        </tr>
        <tr> 
          <th width="24%" scope="row"> 
            <div align="right">Phone # : </div></th>
          <td><%=(userinfo.Fields.Item("mgrphone").Value)%> </td>
        </tr>
        <tr> 
          <th width="24%" scope="row"> 
            <div align="right">Title : </div></th>
          <td><%=(userinfo.Fields.Item("mgrtitle").Value)%> </td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <caption align="top" class="headtext">
        User Information 
        </caption>
        <tr> 
          <th width="24%" scope="row"> <div align="right">Name : </div></th>
          <td><%=(userinfo.Fields.Item("user").Value)%> </td>
        </tr>
        <tr> 
          <th width="24%" scope="row"> <div align="right">Phone # : </div></th>
          <td><%=(userinfo.Fields.Item("userphone").Value)%> </td>
        </tr>
        <tr> 
          <th width="24%" scope="row"> <div align="right">Email :</div></th>
          <td><%=(userinfo.Fields.Item("ex_email").Value)%></td>
        </tr>
        <tr> 
          <th width="24%" scope="row"> <div align="right">Reason ID Needed :</div></th>
          <td><%=(userinfo.Fields.Item("purpose").Value)%> </td>
        </tr>
        <tr> 
          <th scope="row"><div align="right">IP Address :</div></th>
          <td><%=(userinfo.Fields.Item("ex_ip").Value)%></td>
        </tr>
        <tr> 
          <th scope="row"><div align="right">LAN Contact :</div></th>
          <td><%=(userinfo.Fields.Item("ex_lan").Value)%></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <caption align="top" class="headtext">
        Confirmation Information 
        </caption>
        <tr> 
          <th width="24%" scope="row"> 
            <div align="right">Question :</div></th>
          <td><%=(userinfo.Fields.Item("confques").Value)%></td>
        </tr>
        <tr> 
          <th width="24%" scope="row"> 
            <div align="right">Answer :</div></th>
          <td><%=(userinfo.Fields.Item("confansw").Value)%></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><hr size="1" noshade></td>
  </tr>
  <tr> 
    <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <caption align="top" class="headtext">
        User Agreement 
        </caption>
        <tr> 
          <th width="3%" scope="row"> <p>&nbsp;</p></th>
          <td width="97%"><p>The following guidelines are established by the Office 
              of Finance for all FMIS System Users.<br>
              As a system user, it is your responsibility to ensure :</p>
            <ol>
              <li>The confidentiality of you password.</li>
              <li>That the userid will be used for OFFICIAL BUISNESS ONLY.</li>
              <li>That proper care will be exercized to protect all assets while 
                performing your duties.<br>
              </li>
            </ol></td>
        </tr>
        <tr> 
          <th width="3%" scope="row">&nbsp;</th>
          <td>I accept the responsibilities described above : <%=(userinfo.Fields.Item("useragree").Value)%></td>
        </tr>
      </table>
      <hr size="1" noshade></td>
  </tr>
  <tr> 
    <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <caption align="left" class="headtext">
        Configuration Management 
        </caption>
        <tr> 
          <th width="34%" scope="row"> <div align="right">BCAC Mgmt Approver :</div></th>
          <td width="33%"><%=(userinfo.Fields.Item("mgmt_approver").Value)%></td>
          <td width="33%"><strong>Date : </strong><%=(userinfo.Fields.Item("mgmt_app_date").Value)%></td>
        </tr>
        <tr> 
          <th scope="row"> <div align="right">Security Administrator :</div></th>
          <td><%=(userinfo.Fields.Item("sec_approver").Value)%></td>
          <td><strong>Date :</strong> <%=(userinfo.Fields.Item("sec_app_date").Value)%> </td>
        </tr>
      </table>
      <hr size="1" noshade></td>
  </tr>
  <tr> 
    <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <caption class="headtext">
        UserID Information 
        </caption>
        <tr> 
          <th width="22%" scope="row"> <div align="right">UserID Assigned :</div></th>
          <td width="78%"><%=(userinfo.Fields.Item("userid").Value)%></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
<%
userinfo.Close()
Set userinfo = Nothing
%>

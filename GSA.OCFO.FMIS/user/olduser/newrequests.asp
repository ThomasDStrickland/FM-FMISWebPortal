<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/fmisuservb.asp" -->
<%
Dim newusers
Dim newusers_numRows

Set newusers = Server.CreateObject("ADODB.Recordset")
newusers.ActiveConnection = MM_fmisuservb_STRING
newusers.Source = "SELECT *  FROM users  WHERE userid IS NULL"
newusers.CursorType = 0
newusers.CursorLocation = 2
newusers.LockType = 1
newusers.Open()

newusers_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
newusers_numRows = newusers_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim newusers_total
Dim newusers_first
Dim newusers_last

' set the record count
newusers_total = newusers.RecordCount

' set the number of rows displayed on this page
If (newusers_numRows < 0) Then
  newusers_numRows = newusers_total
Elseif (newusers_numRows = 0) Then
  newusers_numRows = 1
End If

' set the first and last displayed record
newusers_first = 1
newusers_last  = newusers_first + newusers_numRows - 1

' if we have the correct record count, check the other stats
If (newusers_total <> -1) Then
  If (newusers_first > newusers_total) Then
    newusers_first = newusers_total
  End If
  If (newusers_last > newusers_total) Then
    newusers_last = newusers_total
  End If
  If (newusers_numRows > newusers_total) Then
    newusers_numRows = newusers_total
  End If
End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = newusers
MM_rsCount   = newusers_total
MM_size      = newusers_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
newusers_first = MM_offset + 1
newusers_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (newusers_first > MM_rsCount) Then
    newusers_first = MM_rsCount
  End If
  If (newusers_last > MM_rsCount) Then
    newusers_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = MM_keepMove & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<html>
<head>
<title>New User Requests</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="fmis.css" rel="stylesheet" type="text/css">
</head>

<body>
<% 
While ((Repeat1__numRows <> 0) AND (NOT newusers.EOF)) 
%>
<table width="600"  border="1" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC" summary="UserID Request Information">
  <tr> 
    <td colspan="2">User Requesting ID : <%=(newusers.Fields.Item("user").Value)%><br> <font size="-2">User Phone : <%=(newusers.Fields.Item("userphone").Value)%></font></td>
  </tr>
  <tr> 
    <td width="50%">System ID Requested For : <%=(newusers.Fields.Item("system").Value)%></td>
    <td width="50%">Date ID Requested : <%=(newusers.Fields.Item("date_requested").Value)%></td>
  </tr>
  <tr> 
    <td colspan="2">Purpose ID Needed : <%=(newusers.Fields.Item("purpose").Value)%></td>
  </tr>
  <tr> 
    <td colspan="2">User Confirmation Question : <%=(newusers.Fields.Item("confques").Value)%></td>
  </tr>
  <tr> 
    <td colspan="2">User Confirmation Answer : <%=(newusers.Fields.Item("confansw").Value)%></td>
  </tr>
  <tr> 
    <td colspan="2">User Agreed to Terms : <%=(newusers.Fields.Item("useragree").Value)%></td>
  </tr>
  <tr> 
    <td colspan="2">Approval Information</td>
  </tr>
  <tr> 
    <td colspan="2">User's Manager Name : <%=(newusers.Fields.Item("mgrname").Value)%><br>
      User's Manager Title : <%=(newusers.Fields.Item("mgrtitle").Value)%><br>
      User's Manager Phone : <%=(newusers.Fields.Item("mgrphone").Value)%></td>
  </tr>
  <tr> 
    <td colspan="2"> <%
	If isNull(newusers.Fields.Item("mgmt_approver").Value) And isNull(newusers.Fields.Item("sec_approver").Value) Then
	%> <a href="action.asp?ref_num=<%=(newusers.Fields.Item("ref_num").Value)%>&method=mgmt">Click 
      here to add management approval for this user.</a><br> <a href="action.asp?ref_num=<%=(newusers.Fields.Item("ref_num").Value)%>&method=sec">Click 
      here to add a security approval for this user.</a><br> <%
	End If
	If isNull(newusers.Fields.Item("sec_approver").Value) And Not isNull(newusers.Fields.Item("mgmt_approver").Value) Then
	%>
      BCAC Mangement Approver : <%=(newusers.Fields.Item("mgmt_approver").Value)%><br>
      Management Apporoval Date : <%=(newusers.Fields.Item("mgmt_app_date").Value)%><br> <a href="action.asp?ref_num=<%=(newusers.Fields.Item("ref_num").Value)%>&method=sec">Click 
      here to add a security approval for this user.</a><br> <%
	End If
	If Not isNull(newusers.Fields.Item("sec_approver").Value) And isNull(newusers.Fields.Item("mgmt_approver").Value) Then
	%>
      Security Approver : <%=(newusers.Fields.Item("sec_approver").Value)%><br>
      Security Approval Date : <%=(newusers.Fields.Item("sec_app_date").Value)%><br> <a href="action.asp?ref_num=<%=(newusers.Fields.Item("ref_num").Value)%>&method=mgmt">Click 
      here to add a BCAC Management approval for this user.</a><br> <%
	End If
	If Not isNull(newusers.Fields.Item("sec_approver").Value) And Not isNull(newusers.Fields.Item("mgmt_approver").Value) And isNull(newusers.Fields.Item("userid").Value) Then
	%>
      BCAC Mangement Approver : <%=(newusers.Fields.Item("mgmt_approver").Value)%><br>
      Management Apporoval Date : <%=(newusers.Fields.Item("mgmt_app_date").Value)%><br>
      Security Approver : <%=(newusers.Fields.Item("sec_approver").Value)%><br>
      Security Approval Date : <%=(newusers.Fields.Item("sec_app_date").Value)%><br> <a href="action.asp?ref_num=<%=(newusers.Fields.Item("ref_num").Value)%>&method=userid">Click 
      here to add a userid for this user.</a><br> <%
	End If
	%> </td>
  </tr>
</table>
<br>
<br>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  newusers.MoveNext()
Wend
%>
<table border="0" width="50%" align="left">
  <tr> 
    <td width="23%" align="center"> <% If MM_offset <> 0 Then %>
      <a href="<%=MM_moveFirst%>">First</a> 
      <% End If ' end MM_offset <> 0 %> </td>
    <td width="31%" align="center"> <% If MM_offset <> 0 Then %>
      <a href="<%=MM_movePrev%>">Previous</a> 
      <% End If ' end MM_offset <> 0 %> </td>
    <td width="23%" align="center"> <% If Not MM_atTotal Then %>
      <a href="<%=MM_moveNext%>">Next</a> 
      <% End If ' end Not MM_atTotal %> </td>
    <td width="23%" align="center"> <% If Not MM_atTotal Then %>
      <a href="<%=MM_moveLast%>">Last</a> 
      <% End If ' end Not MM_atTotal %> </td>
  </tr>
</table>

</body>
</html>
<%
newusers.Close()
Set newusers = Nothing
%>

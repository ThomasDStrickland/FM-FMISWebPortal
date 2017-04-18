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
<html><!-- InstanceBegin template="/Templates/fmis_template.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" --><title>FMIS</title><!-- InstanceEndEditable -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-- Fireworks MX Dreamweaver MX target.  Created Fri Dec 20 08:54:26 GMT-0500 (Eastern Standard Time) 2002-->
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}
function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
 var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
   var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
   if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<link href="../fmis.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../images/about_f2.jpg','../images/training_f2.jpg','../images/faq_f2.jpg','../images/support_f2.jpg','../images/contacts_f2.jpg','../images/email_f2.jpg','../images/cfo_insite_f2.jpg','../images/enter_f2.jpg')">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0"><!--DWLayoutTable-->
  <!-- fwtable fwsrc="fmis_full.png" fwbase="fmis_template.jpg" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
  <tr> 
    <td width="166" height="98"><img name="banner_left" src="../images/banner_left.jpg" width="166" height="98" border="0" alt="FMIS Banner"></td>
    <td width="59" height="98"><img name="banner_left_2" src="../images/banner_left_2.jpg" width="59" height="98" border="0" alt="FMIS Banner"></td>
    <td width="100%" height="98" background="../images/banner_bg.jpg">&nbsp;</td>
    <td width="181" height="98"> <div align="right"><img name="banner_right" src="../images/banner_right.jpg" width="181" height="98" border="0" alt="FMIS Banner"></div></td>
  </tr>
  <tr> 
    <td width="166" height="20"><img name="menu_top" src="../images/menu_top.jpg" width="166" height="20" border="0" alt="Menu Top"></td>
    <td rowspan="11" colspan="3" valign="top" bgcolor="#ffffff"><!-- InstanceBeginEditable name="Body" --><strong>New 
      User Requests</strong><br>
      <br>
      <% 
While ((Repeat1__numRows <> 0) AND (NOT newusers.EOF)) 
%>
      <table width="80%"  border="1" align="center" bordercolor="#CCCCCC" summary="UserID Request Information">
        <tr> 
          <td colspan="2">User Requesting ID : <%=(newusers.Fields.Item("user").Value)%><br> 
            <font size="-2">User Phone : <%=(newusers.Fields.Item("userphone").Value)%></font></td>
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
          <td colspan="2"> 
            <%
	If isNull(newusers.Fields.Item("mgmt_approver").Value) And isNull(newusers.Fields.Item("sec_approver").Value) Then
	%>
            <a href="action.asp?ref_num=<%=(newusers.Fields.Item("ref_num").Value)%>&method=mgmt">Click 
            here to add management approval for this user.</a><br> <a href="action.asp?ref_num=<%=(newusers.Fields.Item("ref_num").Value)%>&method=sec">Click 
            here to add a security approval for this user.</a><br> 
            <%
	End If
	If isNull(newusers.Fields.Item("sec_approver").Value) And Not isNull(newusers.Fields.Item("mgmt_approver").Value) Then
	%>
            BCAC Mangement Approver : <%=(newusers.Fields.Item("mgmt_approver").Value)%><br>
            Management Apporoval Date : <%=(newusers.Fields.Item("mgmt_app_date").Value)%><br> 
            <a href="action.asp?ref_num=<%=(newusers.Fields.Item("ref_num").Value)%>&method=sec">Click 
            here to add a security approval for this user.</a><br> 
            <%
	End If
	If Not isNull(newusers.Fields.Item("sec_approver").Value) And isNull(newusers.Fields.Item("mgmt_approver").Value) Then
	%>
            Security Approver : <%=(newusers.Fields.Item("sec_approver").Value)%><br>
            Security Approval Date : <%=(newusers.Fields.Item("sec_app_date").Value)%><br> 
            <a href="action.asp?ref_num=<%=(newusers.Fields.Item("ref_num").Value)%>&method=mgmt">Click 
            here to add a BCAC Management approval for this user.</a><br> 
            <%
	End If
	If Not isNull(newusers.Fields.Item("sec_approver").Value) And Not isNull(newusers.Fields.Item("mgmt_approver").Value) And isNull(newusers.Fields.Item("userid").Value) Then
	%>
            BCAC Mangement Approver : <%=(newusers.Fields.Item("mgmt_approver").Value)%><br>
            Management Apporoval Date : <%=(newusers.Fields.Item("mgmt_app_date").Value)%><br>
            Security Approver : <%=(newusers.Fields.Item("sec_approver").Value)%><br>
            Security Approval Date : <%=(newusers.Fields.Item("sec_app_date").Value)%><br> 
            <a href="action.asp?ref_num=<%=(newusers.Fields.Item("ref_num").Value)%>&method=userid">Click 
            here to add a userid for this user.</a><br> 
            <%
	End If
	%>
          </td>
        </tr>
      </table>
      <br>
      <hr width="80%" size="1" noshade>
      <br>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  newusers.MoveNext()
Wend
%>
      <!-- InstanceEndEditable --></td>
  </tr>
  <tr> 
    <td width="166" height="31"><a href="../FMIS1.ica" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('enter','','../images/enter_f2.jpg',1)"><img name="enter" src="../images/enter.jpg" width="166" height="31" border="0" alt="Click here to Launch the FMIS Application"></a></td>
  </tr>
  <tr> 
    <td width="166" height="31"><a href="../about.htm" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('about','','../images/about_f2.jpg',1)"><img name="about" src="../images/about.jpg" width="166" height="31" border="0" alt="Click for information About FMIS"></a></td>
  </tr>
  <tr> 
    <td width="166" height="31"><a href="../manual.htm" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('training','','../images/training_f2.jpg',1)"><img name="training" src="../images/training.jpg" width="166" height="31" border="0" alt="Click to view the FMIS Training Manual"></a></td>
  </tr>
  <tr> 
    <td width="166" height="26"><a href="../faq.htm" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('faq','','../images/faq_f2.jpg',1)"><img name="faq" src="../images/faq.jpg" width="166" height="26" border="0" alt="Click for FMIS Frequently Asked Questions"></a></td>
  </tr>
  <tr> 
    <td width="166" height="31"><a href="../support.htm" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('support','','../images/support_f2.jpg',1)"><img name="support" src="../images/support.jpg" width="166" height="31" border="0" alt="Click for FMIS Technical Support"></a></td>
  </tr>
  <tr> 
    <td width="166" height="31"><a href="../contacts.htm" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('contacts','','../images/contacts_f2.jpg',1)"><img name="contacts" src="../images/contacts.jpg" width="166" height="31" border="0" alt="Click to view a listing of FMIS Contacts"></a></td>
  </tr>
  <tr> 
    <td width="166" height="31"><a href="#" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('email','','../images/email_f2.jpg',1)"><img src="../images/email.jpg" alt="Click here to email the FMIS Staff" name="email" width="166" height="31" border="0" onClick="MM_openBrWindow('email.htm','','scrollbars=yes,width=550,height=350')"></a></td>
  </tr>
  <tr> 
    <td width="166" height="31"><a href="http://insite.cfo.gsa.gov" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('cfo_insite','','../images/cfo_insite_f2.jpg',1)"><img name="cfo_insite" src="../images/cfo_insite.jpg" width="166" height="31" border="0" alt="Click here to go to CFO Insite"></a></td>
  </tr>
  <tr> 
    <td width="166" height="71"><img name="gsa" src="../images/gsa.jpg" width="166" height="71" border="0" alt="GSA Logo"></td>
  </tr>
  <tr> 
    <td width="166" valign="top"><!--DWLayoutEmptyCell-->&nbsp;</td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
<%
newusers.Close()
Set newusers = Nothing
%>
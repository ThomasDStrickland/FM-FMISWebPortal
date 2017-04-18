<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<script type="text/javascript" language="javascript"> 

function enableDisable() { 
var bWhich = arguments[0]; 
for (var el,a=1; a<arguments.length; ++a) { 
el = arguments[a]; 
el.disabled = !bWhich; 
if (el.style) el.style.background = bWhich ? '#ffffff' : '#c0c0c0'; //field background 
if (bWhich && a == 1) { 
el.focus(); 
el.select(); 
} 
} 
} 

</script>



<html><!-- InstanceBegin template="/Templates/fmis_template.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" --><title>FMIS</title>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') { num = parseFloat(val);
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
<!-- InstanceEndEditable -->
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
    <td rowspan="11" colspan="3" valign="top" bgcolor="#ffffff"><!-- InstanceBeginEditable name="Body" --><br>
      <form action="process.php" name="form1">
        <table width="80%" border="0" align="center">
          <tr> 
            <td> <input name="date_requested" type="hidden" value="<%=Date()%>"></td>
          </tr>
          <tr> 
            <td><span class="headtext">System ID Requested for</span> : 
              <select name="system">
                <option value="FMIS" selected>FMIS</option>
                <option value="PEGASYS">PEGASYS</option>
                <option value="NEAR">NEAR</option>
              </select></td>
          </tr>
          <tr> 
            <td> <br> <table width="100%"  border="0">
                <caption align="top" class="headtext">
                User Information 
                </caption>
                <tr> 
                  <th width="22%" scope="row"> <div align="right">Name : </div></th>
                  <td> <input type="text" name="user" value="" size="32"></td>
                </tr>
                <tr> 
                  <th width="22%" scope="row"> <div align="right">Phone # : </div></th>
                  <td> <input type="text" name="userphone" value="" size="32"></td>
                </tr>
                <tr> 
                  <th width="22%" scope="row"> <div align="right">Reason ID Needed 
                      :</div></th>
                  <td> <input type="text" name="purpose" value="" size="32"></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td><table width="100%"  border="0">
                <caption align="top" class="headtext">
                <br>
                Confirmation Information 
                </caption>
                <tr> 
                  <th width="22%" scope="row"> <div align="right">Question :</div></th>
                  <td><input type="text" name="confques" value="" size="32"> <strong><font size="-2">(Example 
                    : What high school did I attend?)</font></strong></td>
                </tr>
                <tr> 
                  <th width="22%" scope="row"> <div align="right">Answer :</div></th>
                  <td><input type="text" name="confansw" value="" size="32"> <font size="-2"><strong>(Example 
                    : Douglass High School)</strong></font></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td><table width="100%"  border="0">
                <caption align="top" class="headtext">
                <br>
                Approving Director / Manager Information 
                </caption>
                <tr> 
                  <th width="22%" scope="row"> <div align="right">Manager's Name 
                      : </div></th>
                  <td> <input type="text" name="mgrname" value="" size="32"> </td>
                </tr>
                <tr> 
                  <th width="22%" scope="row"> <div align="right">Manager's Phone 
                      : </div></th>
                  <td> <input type="text" name="mgrphone" value="" size="32"> 
                  </td>
                </tr>
                <tr> 
                  <th width="22%" scope="row"> <div align="right">Manager's Title 
                      : </div></th>
                  <td> <input type="text" name="mgrtitle" value="" size="32"> 
                  </td>
                </tr>
                <tr> 
                  <th width="22%" scope="row"> <div align="right"> Manager's Email 
                      :</div></th>
                  <td><input name="mgr_email" type="text" id="mgr_email" onBlur="MM_validateForm('mgr_email','','RisEmail');return document.MM_returnValue" size="32"> 
                    <strong><font color="#FF0000" size="-2">**Required Field</font></strong></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td><hr size="1" noshade></td>
          </tr>
          <tr> 
            <td><table width="100%"  border="0">
                <caption align="top" class="headtext">
                User Agreement 
                </caption>
                <tr> 
                  <th width="3%" scope="row"> <p>&nbsp;</p></th>
                  <td width="97%"><p>The following guidelines are established 
                      by the Office of Finance for all FMIS System Users.<br>
                      As a system user, it is your responsibility to ensure :</p>
                    <ol>
                      <li>The confidentiality of you password.</li>
                      <li>That the userid will be used for OFFICIAL BUISNESS ONLY.</li>
                      <li>That proper care will be exercized to protect all assets 
                        while performing your duties.<br>
                      </li>
                    </ol></td>
                </tr>
                <tr> 
                  <th width="3%" scope="row">&nbsp;</th>
                  <td>I accept the responsibilities described above : 
                    <select name="useragree" id="useragree">
                      <option value="Yes" selected>Yes</option>
                      <option value="No">No</option>
                    </select></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td> <p> <span class="headtext"><br>
                Is this request for an external user?</span> 
                <label></label>
                <label> 
                <input type="radio" name="external" onclick="enableDisable(this.checked,ex_agency,ex_lan,ex_ip,ex_email)" value="Yes">
                Yes</label>
                <label> 
                <input name="external" type="radio" value="No" checked>
                No</label>
                <br>
              </p></td>
          </tr>
          <tr> 
            <td> <table width="100%"  border="0">
                <tr> 
                  <th width="28%" scope="row"><div align="right">Agency : </div></th>
                  <td width="72%"> <input name="ex_agency" type="text" id="ex_agency" disabled="disabled" style="background:#c0c0c0;"></td>
                </tr>
                <tr> 
                  <th width="28%" scope="row"><div align="right">LAN Contact &amp; 
                      Phone : </div></th>
                  <td> <input name="ex_lan" type="text" id="ex_lan" disabled="disabled" style="background:#c0c0c0;"></td>
                </tr>
                <tr> 
                  <th width="28%" scope="row"><div align="right">IP Address : 
                    </div></th>
                  <td> <input name="ex_ip" type="text" id="ex_ip" disabled="disabled" style="background:#c0c0c0;"></td>
                </tr>
                <tr> 
                  <th width="28%" scope="row"><div align="right">Email Address 
                      : </div></th>
                  <td> <input name="ex_email" type="text" id="ex_email" disabled="disabled" style="background:#c0c0c0;"></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td><input name="submit" type="submit" value="Send Request"></td>
          </tr>
        </table>
      </form>
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

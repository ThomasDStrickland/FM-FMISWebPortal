<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>Request a System UserID</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="fmis.css" rel="stylesheet" type="text/css">
</head>

<body>
<form action="process.php" name="form1">
  <table width="650" border="0" cellspacing="0" cellpadding="0">
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
      <td> <br>
        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
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
            <th width="22%" scope="row"> <div align="right">Reason ID Needed :</div></th>
            <td> <input type="text" name="purpose" value="" size="32"></td>
          </tr>
        </table> </td>
    </tr>
    <tr> 
      <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
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
      <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <caption align="top" class="headtext">
          <br>
          Approving Director / Manager Information 
          </caption>
          <tr> 
            <th width="22%" scope="row"> <div align="right">Manager's Name : </div></th>
            <td> <input type="text" name="mgrname" value="" size="32"> </td>
          </tr>
          <tr> 
            <th width="22%" scope="row"> <div align="right">Manager's Phone : 
              </div></th>
            <td> <input type="text" name="mgrphone" value="" size="32"> </td>
          </tr>
          <tr> 
            <th width="22%" scope="row"> <div align="right">Manager's Title : 
              </div></th>
            <td> <input type="text" name="mgrtitle" value="" size="32"> </td>
          </tr>
          <tr> 
            <th width="22%" scope="row"> <div align="right"> Manager's Email :</div></th>
            <td><input name="mgr_email" type="text" id="mgr_email3" size="32"></td>
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
            <td width="97%"><p>The following guidelines are established by the 
                Office of Finance for all FMIS System Users.<br>
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
            <td>I accept the responsibilities described above : 
              <select name="useragree" id="useragree">
                <option value="Yes" selected>Yes</option>
                <option value="No">No</option>
              </select></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td><input name="submit" type="submit" value="Send Request"></td>
    </tr>
  </table>
</form>
</body>
</html>

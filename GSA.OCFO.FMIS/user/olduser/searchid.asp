<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%> 
<html>
<head>
<title>Search UserIDs</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="fmis.css" rel="stylesheet" type="text/css">
</head>

<body>
<form name="form1" method="get" action="completeid.asp">
  <table width="650" border="0" cellspacing="0" cellpadding="0">
    <caption align="left">
    Choose your search criteria below<br>
    <br>
    </caption>
    <tr> 
      <th width="240" scope="row">
<div align="left">View All UserIDs</div></th>
      <td width="410"><a href="completeid.asp">Click here to view all System UserIDs</a><br>
        <br>
      </td>
    </tr>
    <tr> 
      <th scope="row">
<div align="left">Filter by System</div></th>
      <td>
<select name="system" id="system">
          <option value="" selected>All</option>
          <option value="FMIS">FMIS</option>
          <option value="NEAR">Near</option>
          <option value="PEGASYS">Pegasys</option>
        </select></td>
    </tr>
    <tr> 
      <th scope="row">
<div align="left">Filter by Internal/External Users</div></th>
      <td>
<select name="external" id="external">
          <option selected value="">All</option>
          <option value="No">Internal</option>
          <option value="Yes">External</option>
        </select></td>
    </tr>
    <tr> 
      <th scope="row">
<div align="left"></div></th>
      <td><input type="submit" name="Submit" value="Search IDs"></td>
    </tr>
  </table>
</form>
</body>
</html>

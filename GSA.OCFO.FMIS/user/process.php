<html>
<head>
<title>Request Sent!</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<?php
$aDBLink = @odbc_connect( "fmisuser", "", "" );
if ( $external =="No" )
{
$aSQL  = "INSERT INTO users (user, userphone, date_requested, mgrname, mgrphone, mgrtitle, mgremail, purpose, confques, confansw, useragree, system, external) ";
$aSQL .= "VALUES ('$user', '$userphone', '$date_requested', '$mgrname', '$mgrphone', '$mgrtitle', '$mgr_email', '$purpose', '$confques', '$confansw', '$useragree', '$system', 'No')"; 
}
elseif ( $external == "Yes" )
{
$aSQL  = "INSERT INTO users (user, userphone, date_requested, mgrname, mgrphone, mgrtitle, mgremail, purpose, confques, confansw, useragree, system, external, ex_agency, ex_lan, ex_ip, ex_email) ";
$aSQL .= "VALUES ('$user', '$userphone', '$date_requested', '$mgrname', '$mgrphone', '$mgrtitle', '$mgr_email', '$purpose', '$confques', '$confansw', '$useragree', '$system', '$external', '$ex_agency', '$ex_lan', '$ex_ip', '$ex_email')"; 
}

$aQResult = @odbc_exec( $aDBLink, $aSQL );

$msg  = "The following user has requested a $system UserID and listed you as the approving manager.\n";
$msg .= "User :\t $user \n";
$msg .= "Phone :\t $userphone \n";
$msg .= "Reason ID Requested :\t $purpose \n\n";
$msg .= "If you believe this request to be invalid, please send an email stating so to Central.FMIS@gsa.gov\n\n";
$msg .= "Thank you";

$recipient = "$mgr_email";
$subject = "$system UserID Request for $user";

$mailheaders = "From: central.fmis@gsa.gov \n";

mail($recipient, $subject, $msg, $mailheaders);

echo "<HTML><HEAD><TITLE>Request Sent!</TITLE></HEAD><BODY>";
echo "<P align=center><font face='Verdana, Arial, Helvetica, sans-serif' size='-1'>Thank you for submitting your request for a $system UserID.<BR>";
echo "You will be contacted when your userid is ready.<BR>";
echo "<a href='http://cfo.fmis.gsa.gov'>Click here to return to the main FMIS website.</a></font></P>";
echo "</BODY></HTML>";

?>
</body>
</html>

<?php 

$msg = "FMIS Printer Driver Request\n";
$msg .= "Name:\t$_POST[name]\n";
$msg .= "Login ID:\t$_POST[loginid]\n";
$msg .= "Phone:\t$_POST[phone]\n";
$msg .= "PC Name:\t$_POST[pcname]\n";
$msg .= "Operating System:\t$_POST[os]\n";
$msg .= "Printer Type:\t$_POST[printer]\n\n";

$recipient = "central.fmis@gsa.gov";
$subject = "FMIS Print Driver Request from $_POST[name]";

$mailheaders = "From: FMIS Print Driver Request <> \n";

mail($recipient, $subject, $msg, $mailheaders);

echo "<HTML><HEAD><TITLE>Email Sent!</TITLE></HEAD><BODY>";
echo "<P align=center><font face='Verdana, Arial, Helvetica, sans-serif' size='-1'>Thank You, $_POST[name]<BR>";
echo "Your print driver request has been sent to the FMIS Staff.<BR>";
echo "A member of our staff will be contacting shortly reguarding this issue.<BR>";
echo "<a href='javascript:window.close();'>Click here to close this window</a></font></P>";
echo "</BODY></HTML>";

?>


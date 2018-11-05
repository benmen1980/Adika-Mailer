# UTF-8

# ENVIRONMENT - TABULASYS, TABULAENV
# SQL -  HTML headers, body ,  Subject
# Parameters - IV, PLOG

# - get priority data
# - 1 read IV, IVNUM, BRANCHNAME,
# - 2  check that PDF exists using IVNUM
# - 3 - get HTML & SUBJECT & smtp details
# - 4 - prepare HTML email
# - 5 - email
# - 6- save log in table

[CmdletBinding()]
Param(
  [Parameter(Mandatory=$True,Position=1)]
   [string]$IV,
	
	
   [Parameter(Mandatory=$False,Position=2)]
   [string]$plog,
   [Parameter(Mandatory=$False,Position=3)]
   [string]$dbname

)


Write-Host("Roy debug here")
Write-Host($IV)
Write-Host($dbname)

$Version      = "1.3a";
$Logfile      = "emailer" + "_" + "$env:computername" + "_" + (Get-Date -format yyyy-MM-dd-HH-mm-ss) + ".log";
$TabulaIni    = "$env:SystemRoot\tabula.ini";     
$PriorityDir  = "";        # path to Priority\system\prep
$PriorityTmp  = "";
$PriorityBin  = "";
$PriorityMail = "";
$SQLServer    = "";        # SQL Server Instance.            
$SQLTimeout   = 60;
$Puser	      = "tabula"; 
$Ppassword    = "Adika`$tyle"
$PDFdelay     = 10;      
$mailSentLog  = "";
$mailErrLog   = "";

# Mail settings
$SMTPServer   = "smtp.mandrillapp.com" ;
$SMTPport     = 587;
$SMTPusername = "livne@adikastyle.com";
$SMTPpassword = "Nq0ku-W2tLDpXd-xrsFb4Q";
$SMTPSSL      = $false;
$errMail 	  = "tech@adkiastyle.com; gilad@infobase.co.il";

$body = "";

$error.clear();
if ([string]::IsNullOrEmpty($plog))
{
	$sw = new-object system.IO.StreamWriter("$env:Temp\$Logfile"); # Init logfile stream
}
else
{
	$sw = new-object system.IO.StreamWriter("$plog",$true); # Init logfile from parameter
}

Function LogWrite
{
   Param ([string]$logstring)

   $sw.writeline($logstring);
   $sw.Flush();
   Write-Output ($logstring);
}

Function Mailer 
{ 
	Param ([string]$emailTo, [string]$attachment, [string]$body, [string]$subject, [string]$emailFrom);
	
	if 	(-not $emailTo) 
	{
		throw "Mailer: No recipient specified";
		return $false;	
	}
	if 	(-not $attachment) 
	{
		throw "Mailer: No attachment specified";
		return $false;	
	}
	if 	(-not $subject) 
	{
		throw "Mailer: No subject specified";
		return $false;	
	}
	if 	(-not $emailFrom) 
	{
		throw "Mailer: No sender email (emailFrom) specified";
		return $false;	
	}
	if (-not (test-path $attachment))
	{
		throw "Mailer: $attachment, does not exist";
		return $false;
	}
				
	$att = new-object Net.Mail.Attachment($attachment);
	$msg = new-object Net.Mail.MailMessage;		
	$msg.To.Add($emailTo); # production
	#$msg.Bcc.Add("gilad@infobase.co.il");	
	#$msg.Bcc.Add("livne@adikastyle.com");
	
	#$msg.SubjectEncoding=[System.Text.Encoding]::GetEncoding("windows-1255");
	$msg.SubjectEncoding=[System.Text.Encoding]::GetEncoding("UTF-16");
	$msg.Subject=$subject;
		
	$msg.BodyEncoding=[System.Text.Encoding]::UTF8;
	$msg.Body = $body;
	$msg.IsBodyHtml = $true;
	$msg.From = $emailFrom;			
	#$msg.Headers.Add("Disposition-Notification-To", "$emailFrom"); ## read reciept request
	$msg.Attachments.Add($att);	
	
	$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, $SMTPport) 			
	
	if ([string]::IsNullOrEmpty($SMTPUsername))
	{
		# send without authentication
		$SMTPClient.Send($msg); 
	}
	else
	{	# Send with SMTP Authentication
		$SMTPClient.EnableSsl = $SMTPSSL;
		$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($SMTPusername, $SMTPpassword); 
		$SMTPClient.Send($msg)
	}

	return $true;
}

Write-Host("Logging to file $env:Temp\$logfile");
LogWrite("+ $Version, Started at " + (Get-Date -format yyyy-MM-dd-HH:mm:ss) + " Host: $env:Computername");       

try 
{

	# Parse INI
	LogWrite("> Parsing tabula.ini file ({0})" -f $TabulaIni);
	$ini = Get-Content($TabulaIni);
	foreach ($line in $ini)
	{ 
		$line = $line.split('=');
		$line[0] = $line[0].ToUpper().trim(); 
		if ($line[1] -ne $null) 
		{
			$line[1] = $line[1].trim();
			if ($line[0] -eq "TABULA TMP" -and $PriorityTmp -eq "") { $PriorityTmp = $line[1]; }
			if ($line[0] -eq "TABULA HOST" -and $SQLServer -eq "") { $SQLServer = $line[1]; }
			if ($line[0] -eq "TABULA PATH" -and $PriorityBin -eq "") { $PriorityBin = $line[1]; }
			if ($line[0] -eq "PRIORITY DIRECTORY" -and $PriorityDir -eq "") { $PriorityDir = $line[1]; }
		}
			
	}

	if (-not $PriorityDir.EndsWith("\\")) { $PriorityDir += '\'; }
	if (-not $PriorityTmp.EndsWith("\\")) { $PriorityTmp += '\'; }
	if (-not $PriorityBin.EndsWith("\\")) { $PriorityBin += '\'; }
	$PriorityMail = "$PriorityDir..\mail\";
	$PriorityLoad = "$PriorityDir..\load\";
	New-Item -ItemType Directory -Force -Path $PriorityMail | out-null # try creating the mail\aging folder
	if (-not (test-path -PathType container $PriorityMail)) {throw "$PriorityMail, Path does not exist"}
	if (-not (test-path -PathType container $PriorityTmp))  {throw "$PriorityTmp, Path does not exist"}
	if (-not (test-path -PathType container $PriorityDir))  {throw "$PriorityDir, Path does not exist"}
	if (-not (test-path -PathType container $PriorityBin))  {throw "$PriorityBin, Path does not exist"}
	 
	LogWrite(">> Found Priority in $PriorityDir");
	LogWrite(">> SQL Server: $SQLServer");

	## Import SQL Server modules
	[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.ConnectionInfo');            
	[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.Management.Sdk.Sfc');            
	[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO');            
	[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMOExtended');            
	$srv = New-Object Microsoft.SqlServer.Management.Smo.Server $SQLServer;                     

	# Get DBS
	LogWrite ("> Getting list of databases")
	$QueryDBs = "select top 1 DNAME, HRFLAG, POS, TITLE
				from system.dbo.ENVIRONMENT 
				where DNAME <> ''
				order by HRFLAG desc, POS";
	$QueryCompany = "select top 1 company from system.dbo.t`$license";

	$conn=new-object System.Data.SqlClient.SQLConnection
	#$ConnectionString = "Server={0};Database={1};Integrated Security=True;Connect Timeout={2}" -f $SQLServer,"system",$SQLTimeout
	$ConnectionString = "Server={0};Database={1};Connect Timeout={2};User ID={3};Password={4}" -f $SQLServer,"system",$SQLTimeout,$PUser,$Ppassword
	$conn.ConnectionString=$ConnectionString
	$conn.Open()

	#Get company and database name ($Company, $Db);
	$cmd=new-object system.Data.SqlClient.SqlCommand($QueryCompany,$conn)
	$cmd.CommandTimeout=$SQLTimeout
	$ds=New-Object system.Data.DataSet
	$da=New-Object system.Data.SqlClient.SqlDataAdapter($cmd)
	[void]$da.fill($ds)
	$Company = $ds.Tables[0].Rows[0][0];
	$cmd=new-object system.Data.SqlClient.SqlCommand($QueryDBs,$conn)
	$cmd.CommandTimeout=$SQLTimeout
	$ds=New-Object system.Data.DataSet
	$da=New-Object system.Data.SqlClient.SqlDataAdapter($cmd)
	[void]$da.fill($ds)


	#$Db= $env:TABULAENV
	#$db = "a080413";
    $db = $dbname;

	if (-not $db) 
	{
		throw "`$db not defined";
	}

	LogWrite("> License: $Company");
	LogWrite("`r`n> Starting ... `r`n");

	#Get invoice details
	$QueryInvoices = " 
	use $db; `r`n" +
	"select top 1 I.IVNUM, B.BRANCHNAME, CA.EMAIL, I.ORD, B.EMAIL [EMAILFROM]
	from INVOICES I inner join 
		BRANCHES B on I.BRANCH = B.BRANCH inner join
		CUSTOMERSA CA on CA.CUST = I.CUST
	where I.IV = @iv";

	$QueryBody = "
	use $db;
	declare @text as nvarchar(max);
	set @text = ' ';
	select @text = @text + coalesce(reverse(TEXT),'') + ' '
	from INFO_SHIPTYPESTEXT
	where SHIPTYPE = @shiptype
	order by TEXTORD;
	select @text;
	";

	$QueryEmail = "
	SELECT C.CUSTNAME [CUSTNAME] ,
		O.BOOKNUM , ZADK_DELIVERYCODE [DCODE],
		S.STCODE , S.SHIPTYPE, 
		convert(nvarchar(100),system.dbo.tabula_hebconvert(coalesce(S.INFO_MAILHEADER,''))) [SUBJECT]
	FROM ORDERS O inner join 
		SHIPTYPES S on O.SHIPTYPE = S.SHIPTYPE inner join
		CUSTOMERS C on C.CUST = O.CUST
	WHERE O.ORD = @ord
	";

	$cmdInvoice=new-object system.Data.SqlClient.SqlCommand($QueryInvoices,$conn);
	$cmdInvoice.CommandTimeout=$SQLTimeout;
	$cmdInvoice.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@iv",[Data.SQLDBType]::BigInt))) | out-null;

	$cmdBody=new-object system.Data.SqlClient.SqlCommand($QueryBody,$conn);
	$cmdBody.CommandTimeout=$SQLTimeout;
	$cmdBody.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@shiptype",[Data.SQLDBType]::VarChar, 66))) | out-null;

	$cmdEmail=new-object system.Data.SqlClient.SqlCommand($QueryEmail,$conn);
	$cmdEmail.CommandTimeout=$SQLTimeout;
	$cmdEmail.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@ord",[Data.SQLDBType]::BigInt))) | out-null;

	$cmdInvoice.Parameters[0].Value = $iv;
	$ds=New-Object system.Data.DataSet
	$da=New-Object system.Data.SqlClient.SqlDataAdapter($cmdInvoice);
	[void]$da.fill($ds);
	#$ds.Tables[0];
	#return

	if ($ds.Tables -ne $null)
	{
		$ivnum=$ds.Tables[0].Rows[0]['IVNUM'];
		$branchname=$ds.Tables[0].Rows[0]['BRANCHNAME'];
		$branchname=$branchname.toLower();
		$email=$ds.Tables[0].Rows[0]['EMAIL'];
		$ord=$ds.Tables[0].Rows[0]['ORD'];
		$emailFrom=$ds.Tables[0].Rows[0]['EMAILFROM'];
	}

	if ([string]::IsNullOrEmpty($branchname))
	{
		$branchname = 'adika';
		$emailFrom = "Adika Style <info@adikastyle.com>"
	}

	if ([string]::IsNullOrEmpty($emailFrom))
	{
		throw "Emailfrom (sender's mail) null or empty";
	}

	if ([string]::IsNullOrEmpty($email))
	{
		throw "Email null or empty";
	}

	if ([string]::IsNullOrEmpty($ord))
	{
		throw "`$ord null or empty";
	}

	LogWrite (">> Ivnum: $ivnum($iv), Branchname: '$branchname', Email: $email, Ord: $ord");

	$cmdEmail.Parameters[0].Value = $ord;
	$ds=New-Object system.Data.DataSet
	$da=New-Object system.Data.SqlClient.SqlDataAdapter($cmdEmail);
	[void]$da.fill($ds)
	if ($ds.Tables -ne $null)
	{
		$custname=$ds.Tables[0].Rows[0]['CUSTNAME'];
		$subject=$ds.Tables[0].Rows[0]['SUBJECT'];
		$shiptype=$ds.Tables[0].Rows[0]['SHIPTYPE'];
		$dcode=$ds.Tables[0].Rows[0]['DCODE'];
		$booknum=$ds.Tables[0].Rows[0]['BOOKNUM'];
	}

	LogWrite (">> Custname: '$custname', ($subject), Shiptype: $shiptype, Dcode:$dcode, Booknum:$booknum");

	$cmdBody.Parameters[0].Value = $shiptype;
	$body = $cmdBody.ExecuteScalar(); 	
	if ([string]::IsNullOrEmpty($body))
		{
			throw "No BODY content";
		}
			
	if (-not (test-path "$PriorityMail\templates\$branchname.html"))
		{
			throw "$PriorityMail\templates\$branchname.html Branch email template file not found ";
		}
		
	$html="";
	$html = Get-Content "$PriorityMail\templates\$branchname.html";

	$body = $body.replace("#",$dcode);
	$body = $body.replace("!",$booknum);
	$html = $html.replace("{{body}}","<!--body start-->" + $body  + "<!--body end-->");
	$html = $html.replace("{{var customer.email}}",$email);

	if ([string]::IsNullOrEmpty($subject)) 
	{
		$subject = "Empty subject";
	}

	#base64 utf8 subject encoding
	#$b=[System.Text.Encoding]::UTF8.GetBytes($subject);
	#$c=[System.Convert]::ToBase64String($b);
	#$subject = ("=?UTF-8?B?{0}?=" -f $c);

	$attachment = $priorityLoad + 'IV_' + $ivnum + '.pdf';
	if (-not (test-path $attachment))
		{
			throw "Mailer: $attachment, does not exist";
			return $false;
		}
	
	LogWrite(">> Attaching $attachment");
	LogWrite(">> Sending to $email, Subject:$subject, From:$emailFrom");
	Mailer $email $attachment $html $subject $emailFrom
	
}
catch 
{		
	LogWrite ("!> Error while emailing: " + $error[0]);	
	$sw.Dispose();
	Mailer $errMail "$env:Temp\$Logfile" ("Error while emailing: " + $error[0]) "Adika-Mailer error" "info@adikastyle.com" ;
	$error.clear();
}		
	
$sw.Dispose();

#pause

#$html | Out-File "c:\tmp\1.html";

#Remove-Item -Force "$env:Temp\$Logfile"; # fails

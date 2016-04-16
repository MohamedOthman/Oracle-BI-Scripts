param([string]$Env)
### Global Variables Section 
$MAIN_PATH=
$RPD_PATH=
$TARGET_COMMAND_FILE=
$OBIEE_CLIENT_BIN=
$sharepoint_file=
$VARIABLES_CVS=
$PASSWORD_CVS=
$Config_Sheet=
$Config_File=
$PASSWORD_SHEET=

function evaluate_env() {
if ( $Env -eq "PRD" ) 
{
	echo "the target environment is PROD"
	$global:COLUMN_ID="B"
	$global:field_id=1
}
elseif ( $Env -eq "UAT" ) 
{
	echo "the target environment is UAT"
	$global:COLUMN_ID="C"
	$global:field_id=2
}
elseif ( $Env -eq "SUP" ) 
{
	echo "the target environment is SUP"
	$global:COLUMN_ID="D"
	$global:field_id=3
}
elseif ( $Env -eq "RIN" ) 
{
	echo "the target environment is RIN"
	$global:COLUMN_ID="E"
	$global:field_id=4
}
else 
{
	echo "this is not valid option please enter one of the followings  PRD | UAT | SUP | RIN  "
	exit 
}
}


function download_password_sheet() {
echo "downloading password sheet"
$Filename = [System.IO.Path]::GetFileName($sharepoint_file)
$tofile="E:\RPD_PREP\$Filename" 
$username= Read-Host 'enter your sharepoint username'
$password= Read-Host -AsSecureString 'enter your sharepoint Password' 
$PasswordPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
# Get the plain text version of the password
$PlainTextPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto($PasswordPointer)
# Free the pointer
[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($PasswordPointer)
# Return the plain text password
$webclient = New-Object System.Net.WebClient
$webclient.Credentials = new-object System.Net.NetworkCredential($username, $PlainTextPassword)
$webclient.DownloadFile($sharepoint_file,$tofile)
}


function print_rpd_variables () {

$SheetName = "RPD Variables"

# Create an Object Excel.Application using Com interface
$objExcel = New-Object -ComObject Excel.Application

# Disable the 'visible' property so the document won't open in excel
$objExcel.Visible = $false

# Open the Excel file and save it in $WorkBook
$WorkBook = $objExcel.Workbooks.Open($PASSWORD_SHEET)

# Load the WorkSheet 'BuildSpecs'
$WorkSheet = $WorkBook.sheets.item($SheetName)
$intRowMax =  ($worksheet.UsedRange.Rows).count

for($intRow = 3 ; $intRow -le $intRowMax ; $intRow++)
{
 $VAR = $worksheet.Range("A"+$intRow).Text
 $Variable = $worksheet.Range($COLUMN_ID+$intRow).Text
 "SetProperty `"variable`" `"$VAR`" initializer `"`'$Variable`'`"" | out-file -append -encoding "UTF8"  $TARGET_COMMAND_FILE
}
$objExcel.quit()
}


### Variables of connection pools 
function password_values() {
$DACREP7964_tmp=cat $PASSWORD_CVS | select-string -pattern "DACREP7964"
$global:DACREP7964=echo "$DACREP7964_tmp" |  %{$_.split(',')[1]}
$SIEBEL_tmp=cat $PASSWORD_CVS | select-string -pattern "SIEBEL"
$global:SIEBEL=echo "$SIEBEL_tmp" |  %{$_.split(',')[1]}
$UNITY_BIPLATFORM_tmp=cat $PASSWORD_CVS | select-string -pattern "UNITY_BIPLATFORM" 
$global:UNITY_BIPLATFORM=echo "$UNITY_BIPLATFORM_tmp" |   %{$_.split(',')[1]}
$ORA_EBS_DSN_tmp=cat $VARIABLES_CVS | select-string -pattern "ORA_EBS_OLTP_DSN"
$global:ORA_EBS_DSN= echo "$ORA_EBS_DSN_tmp" |  %{$_.split(',')[$field_id]}
$ORA_EBS_USER_tmp= cat $PASSWORD_CVS | select-string -pattern "XXOC_OBIEE" | select-string -pattern $ORA_EBS_DSN
$global:ORA_EBS_USER=echo "$ORA_EBS_USER_tmp" |  %{$_.split(',')[1]}
$DWH7964_tmp=cat $PASSWORD_CVS | select-string -pattern "DWH7964"
$global:DWH7964=echo "$DWH7964_tmp" |  %{$_.split(',')[1]}

echo "password of dacrep7964 is $DACREP7964 "
echo "password of SIEBEL is $SIEBEL"
echo "password of UNITY_BIPLATFORM is $UNITY_BIPLATFORM"
echo "password of ORA_EBS_USER is $ORA_EBS_USER"
echo "passowrd of DWH7964 is $DWH7964"

}

function SaveToCSV() {
$ExcelWB = new-object -comobject excel.application
$ExcelWB.Visible = $false
$ExcelWB.DisplayAlerts = $FALSE
$Workbook = $ExcelWB.Workbooks.Open("$PASSWORD_SHEET")
$WorkSheet=$WorkBook.sheets.item("RPD Variables")
$WorkSheet.SaveAs("$VARIABLES_CVS" ,6)
$WorkSheet=$WorkBook.sheets.item("$Env")
$WorkSheet.SaveAs("$PASSWORD_CVS" ,6)
$WorkSheet=$WorkBook.sheets.item("$Config_Sheet")
$WorkSheet.SaveAs("$Config_File" ,6)
$Workbook.Close() 
$ExcelWB.quit()
} 

function connection_pool() {
## this function usage will be 
## read config file and get the following 
## connection pool name by cutting the first tow strings 
## username which will be the third one 
## password and this one will be gotten by the below check 
## get the username and check the value of DSN is it starts with VALUEOF(test) .. it will grep it from variables CSV and then grep it from password CSV 
## if it not VALUEOF .. it will be just instance name and will be gotten directly from the password sheet 
## here are the steps 
$Lines= Get-Content $Config_File
foreach ( $Line in $Lines ) {
$PhysicalConnection= echo $Line | %{$_.split(',')[0]}
$ConnectionPool=echo $Line  | %{$_.split(',')[1]}
$UserName=echo $Line | %{$_.split(',')[2]}
$DataSource=echo $Line | %{$_.split(',')[3]}

if ( $UserName.StartsWith("VALUEOF") ) 
{
		$UserName_f=echo $UserName | %{$_.split('(')[1]} | %{$_.split(')')[0]}
		echo $UserName_f
		$UserName_tmp=cat $VARIABLES_CVS | select-string -pattern "^$UserName_f,"
		$global:UserName_final=echo "$UserName_tmp" |  %{$_.split(',')[$field_id]}
}
else 
{
		$global:UserName_final=$UserName
}

if ( $DataSource.StartsWith("VALUEOF") ) 
{
		$DataSource_f=echo $DataSource | %{$_.split('(')[1]} | %{$_.split(')')[0]}
		$DSN_tmp=cat $VARIABLES_CVS | select-string -pattern "^$DataSource_f,"
		$global:DSN=echo "$DSN_tmp" |  %{$_.split(',')[$field_id]} 
}
else 
{
		$global:DSN=$DataSource
}

## the below lines write the code to the command line of obiee 
$PASSWORD_tmp=cat $PASSWORD_CVS | select-string -pattern "^$UserName_final," | select-string -pattern "$DSN" 
$global:PASSWORD=echo "$PASSWORD_tmp" | %{$_.split(',')[1]}
echo "the password of $UserName_final is $PASSWORD" 

"SetProperty `"Connection Pool`"  `"$PhysicalConnection`".`"$ConnectionPool`" `"Password`" `"$PASSWORD`"" | out-file -append -encoding "UTF8"  $TARGET_COMMAND_FILE
"SetProperty `"Connection Pool`"  `"$PhysicalConnection`".`"$ConnectionPool`" `"User`" `"$UserName`"" | out-file -append -encoding "UTF8"  $TARGET_COMMAND_FILE
"SetProperty `"Connection Pool`"  `"$PhysicalConnection`".`"$ConnectionPool`" `"DSN`" `"$DataSource`"" | out-file -append -encoding "UTF8"  $TARGET_COMMAND_FILE

}
}


##################################
#### Main Script starts here #####
##################################
if ( Test-Path $RPD_PATH )
{
	echo "RPD file exists the script will proceed " 
}
else 
{
	echo "RPD file doesn't exist . please check the path of file and make sure that the RPD file path is "
	echo "$RPD_PATH"
	echo "the script will abort" 
	exit 
}
evaluate_env
download_password_sheet
SaveToCSV
"`' this is command file to change the rpd from command line"  | out-file -filepath  $TARGET_COMMAND_FILE  -encoding "UTF8" 
"OpenOffline `"$RPD_PATH`" Canon123" | out-file -append -encoding "UTF8" $TARGET_COMMAND_FILE
"`'Variables Section" | out-file -append -encoding "UTF8" $TARGET_COMMAND_FILE
print_rpd_variables
"`'connection pools section " | out-file -append -encoding "UTF8" $TARGET_COMMAND_FILE

connection_pool
#rm -r $MAIN_PATH\*.cvs 
#rm $PASSWORD_SHEET
"`'end of connection pool section" | out-file -append -encoding "UTF8"  $TARGET_COMMAND_FILE 
"SaveAs `"E:\RPD_PREP\OracleBIAnalyticsApps.rpd`"" | out-file -append -encoding "UTF8"  $TARGET_COMMAND_FILE 
"Close" | out-file -append -encoding "UTF8"  $TARGET_COMMAND_FILE 
"Exit" | out-file -append -encoding "UTF8"  $TARGET_COMMAND_FILE 

cd $OBIEE_CLIENT_BIN 
.\admintool.exe /command $TARGET_COMMAND_FILE
cd $MAIN_PATH

##################################
#### 	  End of Script        #####
##################################

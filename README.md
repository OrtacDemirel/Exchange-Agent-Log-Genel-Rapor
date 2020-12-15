# Exchange-Agent-Log-Genel-Rapor
Tüm Exchange AgentLog kayıtlarının HTML olarak raporlanmasını ve email atılmasını sağlar
# Ortaç Demirel   - www.ortacdemirel.com 
# Exchange Server Agent Log Report  
# Agent Log un tuttuğu tüm kayıtları getirir. 
# Rapor klasörünün altına rapor.htm dosyasını yaratır. 
# Raporun email atılması için ayarlar yapılmalıdır. 
 
 
# $MailTo = "user@domain.com" 
# $MailServer = "10.0.0.100" 
# $MailFrom = "agentreport@domain.com" 
 
 
$rapor = @"  
<Title> Exchange Agent Log Raporu </Title> 
 
<mce:style><!-- 
BODY 
{background-color:peachpuff;} 
TABLE 
{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;} 
TH 
{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:thistle} 
TD 
{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:palegoldenrod;padding:10px} 
 
th:nth-child(1) {min-width:180px !important} 
th:nth-child(2) {min-width:150px !important} 
th:nth-child(3) {min-width:100px !important} 
--></mce:style><style _mce_bogus="1"><!-- 
BODY 
{background-color:peachpuff;} 
TABLE 
{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;} 
TH 
{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:thistle} 
TD 
{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:palegoldenrod;padding:10px} 
 
th:nth-child(1) {min-width:180px !important} 
th:nth-child(2) {min-width:150px !important} 
th:nth-child(3) {min-width:100px !important} 
--></style> 
<H1>Exchange Agent Log Raporu</H1> 
"@ 
 
Get-agentlog | Select-Object timestamp,IPAddress,P1fromaddress,@{n='Recipients';e={$_.recipients}},Action,SmtpResponse,Reason,ReasonDAta,Directionality | ConvertTo-HTML -Head $rapor | Out-File C:\rapor\rapor.htm 
 
# $output = "c:\rapor\rapor.htm" 
# Send-MailMessage -Attachments $output -To $MailTo -From $MailFrom -Subject "Exchange Antispam Report" -BodyAsHtml $Output -SmtpServer $MailServer

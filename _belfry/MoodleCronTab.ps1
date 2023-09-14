$url = "http://127.0.0.1:8080/moodle/admin/cron.php"
$myPath = "C:\Bitnami\cron-logs\log$(get-date -f yyyyMMddhhss).htm"
$client = new-object System.Net.WebClient
$client.Proxy = $null
$client.DownloadFile( $url, $myPath)
#[io.file]::WriteAllBytes((Invoke-WebRequest -URI $url).content,$myPath)
#Invoke-WebRequest -URI $url -OutFile $myPath -TimeoutSec 60 -UseBasicParsing
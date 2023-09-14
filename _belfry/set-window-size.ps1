# set windows width and height as well as buffer width and height
# https://blogs.technet.microsoft.com/heyscriptingguy/2006/12/04/how-can-i-expand-the-width-of-the-windows-powershell-console/
# retrieved: 2017-03-20
#
$pshost = Get-Host #get-host

$pswindow = $pshost.UI.RawUI
$newsize = $pswindow.buffersize
$newsize.height = 3000
$newsize.width = 150
$pswindow.buffersize = $newsize

$newsize = $pswindow.WindowSize
$newsize.height = 50
$newsize.width = 150
$pswindow.windowsize = $newsize

.\_pssetup.ps1

[string]$corrid = Read-Host 'What is the Corelation ID to search for?'

get-splogevent | ?{$_.Correlation -eq $corrid} | select Area, Category, Level, EventID, Message |Format-List

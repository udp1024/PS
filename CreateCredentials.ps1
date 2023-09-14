$password = ConvertTo-SecureString “P@ssw0rd” -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential (“duffney”, $password)


#how to handle empty credentials
if($Credential -ne [System.Management.Automation.PSCredential]::Empty) {
    Invoke-Command -ComputerName:$ComputerName -Credential:$Credential  {
        Set-ItemProperty -Path $using:Path -Name $using:Name -Value $using:Value
    }
} else {
    Invoke-Command -ComputerName:$ComputerName {
        Set-ItemProperty -Path $using:Path -Name $using:Name -Value $using:Value
    }
}


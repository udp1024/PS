$ma = Get-SPManagedAccount GOA\CHR.Applications.S
$maType = $ma.GetType()

$bindingFlags = [Reflection.BindingFlags]::NonPublic -bor [Reflection.BindingFlags]::Instance

$m_Password = $maType.GetField("m_Password", $bindingFlags) 
$pwdEnc = $m_Password.GetValue($ma) 

$ssv = $pwdEnc.SecureStringValue 
$ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($ssv) 
[System.Runtime.InteropServices.Marshal]::PtrToStringUni($ptr)

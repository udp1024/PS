function Bindings()
{
return [System.Reflection.BindingFlags]::CreateInstance -bor
[System.Reflection.BindingFlags]::GetField -bor
[System.Reflection.BindingFlags]::Instance -bor
[System.Reflection.BindingFlags]::NonPublic
}
function GetFieldValue([object]$o, [string]$fieldName)
{
$bindings = Bindings
return $o.GetType().GetField($fieldName, $bindings).GetValue($o);
}
function ConvertTo-UnsecureString([System.Security.SecureString]$string)
{
$intptr = [System.IntPtr]::Zero
$unmanagedString = [System.Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($string)
$unsecureString = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($unmanagedString)
[System.Runtime.InteropServices.Marshal]::ZeroFreeGlobalAllocUnicode($unmanagedString)
return $unsecureString
}
Get-SPManagedAccount | select UserName, @{Name="Password"; Expression={ConvertTo-UnsecureString (GetFieldValue $_ "m_Password").SecureStringValue}}

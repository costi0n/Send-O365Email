####
##
#
#
##
####  



function Get-StringHash {
    param
    (
        [String] $String,
        $HashName = "MD5"
    )
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($String)
    $algorithm = [System.Security.Cryptography.HashAlgorithm]::Create('MD5')
    $StringBuilder = New-Object System.Text.StringBuilder 
  
    $algorithm.ComputeHash($bytes) | 
    ForEach-Object { 
        $null = $StringBuilder.Append($_.ToString("x2")) 
    } 
  
    $StringBuilder.ToString() 
}

function Get-Unique-Id ($recipient) {

    $hwSerial = (Get-WmiObject win32_bios).SerialNumber
    $uh = $env:username + $hwSerial + $recipient
    Return Get-StringHash $uh
}


#Get-Unique-Id("c.ghita@netcare.it")

Clear-Host

$emailMitente = "medicinadibasecovid@asl.rieti.it"
$nomeVisualizzato = $null

$zz = "$($nomeVisualizzato) <$($emailMitente)>"



$re="[a-z0-9!#\$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#\$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?"
$mail = [regex]::MAtch($zz, $re, "IgnoreCase ")

$mail.value
<#
 .Synopsis
  Send email via Office365 / ExchageOnline V0.1

 .Description
  Send email via Office365 / ExchageOnline Account. 

 .Parameter Recipient
  Mandatory, the recipient email address 

 .Parameter Sender
  Mandatory, the sender email address "SENDER <sender@example.com>"

 .Parameter crdpath
  Optional, path where will be stored the O365 credentials, if this is empty will be stored on temporary user folder
  ( at the firs run the user will be ask to insert username and password of sender email account )

 .Parameter bbc
  Only for debugging purpose Blind carbon copy recipient email address

 .Parameter subject
  Recommended, there is the subject of sended email

 .Parameter HtmlBody
  Optional, only if we want to add some extra html, for experts only

 .Parameter template
  Mandatory, full path to the template html file with placeholders

.Parameter template
  Mandatory, this object will contain all user data to be used on html template, see example below

.Example

    $utente_nuovo = @{
        nome          = "Firstname"
        cognome       = "Surname"
        nomeutente    = "f.surname" 
        cellulare     = "3297660848"
    }

    $recipient = "recipient@email.com"
    $template = "c:\folder_wehere_is_html\EmailTemplateNew.html"
    $crdpath = "c:\folder_wehre_is_xml\"

    $return = Send-O365Email -recipient $recipient -sender "SENDER <sender@example.com>" -crdpath $crdpath -template $template -datiUtente $utente_nuovo
    $return
#>


# Genera hash partendo da una stringa
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


# Genera id unico basato su nome utente e hardware serial number
function Get-Unique-Id {
    $hwSerial = (Get-WmiObject win32_bios).SerialNumber
    $uh = $env:username + $hwSerial
    Return Get-StringHash $uh
}
# maschera una stringa per privacy
function MaskSring {
    param (
        [string] $var = $null
    )
    $length = $var.length

    $begin = $var.substring(0, 3)
    $end = $var.substring($length - 3, 3)
    $rst = $length - ($begin.Length) - ($end.Length)
    $m = $null
    For ($i = 1; $i -le $rst; $i++) {
        $m = $m + "*"
    }
    Return($begin + $m + $end) 
}

# invia il messaggio di posta via Office365 in formato Html
function Send-O365Email {
    param (
        [string] $recipient = $null,
        [string] $sender = $null,
        [string] $crdpath = $null,
        [string] $bcc = $null,
        [string] $subject = "Credenziali Account Aziendale",
        [array] $HtmlBody = @(),
        [string] $template = $null,
        [Object] $datiUtente = @{}
    )
 
    $xml = Get-Unique-Id
    $crdXML = $crdpath + $xml + ".xml"

    if ( Test-Path $crdXML ) {
        $cred = Import-Clixml $crdXML
    }
    else {
        Get-Credential $sender | Export-Clixml  $crdXML #Store Credentials
        if ( Test-Path $crdXML ) {
            $cred = Import-Clixml $crdXML
        }
        else {
            Return "Qualcosa non va con le tue credenziali !"
            Exit
        }
    }

    $HtmlContent = Get-Content -Path $template

    foreach ($ContentLine in $HtmlContent) {
        $cellulare = MaskSring -var $datiUtente.cellulare
        # If more variables are added to the message, just copy and modIfy the lines below.
        $ContentLine = $ContentLine `
            -replace '{nome}'     , $datiUtente.nome`
            -replace '{cognome}'  , $datiUtente.cognome`
            -replace '{nomeutente}' , $datiUtente.nomeutente`
            -replace '{cellulare}' , $cellulare`

        $HtmlBody += $ContentLine
    }

    $EmailParams = @{
        SmtpServer  = 'smtp.office365.com'
        From        = $sender
        To          = $recipient
        Bcc         = $bcc
        Subject     = $subject
        BodyAsHtml  = $true
        Body        = ($HtmlBody | Out-String)
        ErrorAction = 'Stop'
        Encoding    = 'UTF8'
        Port        = '587'
    }


    try {
        Send-MailMessage @EmailParams -usessl -Credential $cred -Priority High
        Return "inviato a: " + $recipient
    }
    Catch { 
        return "Si è verificato un errore ! Il messaggio non è stato inviato !"
        Exit 1
    }

}

Export-ModuleMember -Function Send-O365Email
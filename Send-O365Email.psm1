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
        cellulare     = "3291234848"
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
    param (
        [String] $thissender = $null
    )
    $hwSerial = (Get-WmiObject win32_bios).SerialNumber
    $uh = $env:username + $hwSerial + $thissender
    Return Get-StringHash $uh
}
# maschera una stringa per privacy
function MaskSring {
    param (
        [string] $var = $null
    )

    if ($var) {
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
    else {
        Return $null
    }

}

# invia il messaggio di posta via Office365 in formato Html
function Send-O365Email {
    param (
        [string] $recipient = $null,
        [string] $sender = $null,
        [string] $crdpath = $null,
        [string] $cc  = $null,
        [string] $bcc = $null,
        [string] $subject = "Credenziali Account Aziendale",
        [array] $HtmlBody = @(),
        [string] $template = $null,
        [Object] $datiUtente = @{ },
        [string] $Attachments = $null,
        [string] $Masked = $false,
        [string] $verbose = $false
    )
 
    
    $re="[a-z0-9!#\$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#\$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?"
    $mailSender = [regex]::MAtch($sender, $re, "IgnoreCase ")


    $xml = Get-Unique-Id $mailSender.Value
    $crdXML = $crdpath + "\" + $xml + ".xml"

    if ( Test-Path $crdXML ) {
        $cred = Import-Clixml $crdXML
    }
    else {
        Get-Credential $mailSender.Value | Export-Clixml  $crdXML #Store Credentials
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
        if ( $Masked ) { 
                #$cellulare = MaskSring -var $datiUtente.cellulare
                $cellulare = $datiUtente.cellulare
            } else {  
                $cellulare = $datiUtente.cellulare 
            } 
        # If more variables are added to the message, just copy and modIfy the lines below.
        $ContentLine = $ContentLine `
            -replace '{nome}'     , $datiUtente.nome`
            -replace '{cognome}'  , $datiUtente.cognome`
            -replace '{nomeutente}' , $datiUtente.nomeutente`
            -replace '{cellulare}' , $cellulare`
            -replace '{errata}', $datiUtente.errata`

        $HtmlBody += $ContentLine
    }

    $EmailParams = @{
        SmtpServer  = 'smtp.office365.com'
        From        = $sender
        To          = $recipient
        Subject     = $subject
        BodyAsHtml  = $true
        Body        = ($HtmlBody | Out-String)
        ErrorAction = 'Stop'
        Encoding    = 'UTF8'
        Port        = '587'
    }

    # add on to hash table $EmailParams the $bcc if is not null
    if ( $bcc ) { $EmailParams.Add( "Bcc", $bcc) }
    #add on cc hash table $EmailParams the $cc if is not null
    if ( $cc ) { $EmailParams.Add( "Cc", $cc) }
    # add on to hash table $EmailParams the $Attachments if is not null
    if ( $Attachments ) { $EmailParams.Add( "Attachments", $Attachments) }

    try {
        Send-MailMessage @EmailParams -usessl -Credential $cred -Priority High
        if ($verbose -eq $true) {
            $ret = "To:$($recipient) "
            if ( $cc ) {
                $ret += "Cc:$($cc) "
            }
            if ( $bcc ) {
                $ret += "Bcc:$($bcc)"
            }
        } else {
            $ret = $recipient
        }

        Return $ret
    }
    Catch { 
        return "Si è verificato un errore ! Il messaggio non è stato inviato !"
        Exit 1
    }

}

Export-ModuleMember -Function Send-O365Email
<#
.Synopsis
   Disable AD User Object and hide from Global Address List.
.DESCRIPTION
   Disable AD User Object, move to NoLongerEmployed OU,
   set new Description and hide from Global Address List,
   all in one go.
.EXAMPLE
   Unemploy-User ZPamfil -Seriously YES -SerenaTicketID 123456 -RequestInitiator 'Jane Doe'
.EXAMPLE
   Unemploy-User ZPamfil -Seriously YES -Verbose
   This will print out action details and prompt for SerenaTicketID and RequestInitiator.
.EXAMPLE
   Unemploy-User ZPamfil -WhatIf
   For a Dry-run without touching the user object.
#>
function Unemploy-User
{
    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([string])]
    Param
    (
        # AD username to be remove from all groups.
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0)][string]$username,
        [Parameter(Mandatory=$true)][string]$SerenaTicketID,
        [Parameter(Mandatory=$true)][string]$RequestInitiator,
        [Parameter()][string]$Seriously
    )


    Begin 
    {
        if (!($Seriously -or ($WhatIfPreference -eq $True)))
        { 
            Write-Host -ForegroundColor Yellow "
            You need to specify -Seriously:YES as parameter to really do it.
            Example: Unemploy-User ZPamfil -Seriously YES
            For help, do: Get-Help Unemploy-User
            "
            Throw ('Not serious.')
        }
        $filter = 'OSWOriflameAllEmployees', 'Domain Users', '*AllUserObjects'
    }


    Process
    {
        Try
        {
            Get-ADUser $username -property DistinguishedName, DisplayName, Title | Select-Object DistinguishedName, DisplayName, Title | Format-List
        }
        Catch
        {
            Throw "No such user object ($username)."
        }

        Write-Verbose "Disabling AD user object $username"
        Disable-ADAccount $username
        
        Write-Verbose "Set new Description"
        Set-ADUser $username -Description "Disabled $(Get-Date -Format 'dd.MM.yyyy') / Ticket#: $SerenaTicketID / Requested by: $RequestInitiator"

        if (TryGoodbyeOU($username))
        {
            Write-Verbose 'Moving freshly disabled user to NoLongerEmployed OU'
            Get-ADUser $username | Move-ADObject -TargetPath (GetGoodbyeOU($username))
        }
        else
        {
            Write-Warning "No OU named NoLongerEmployed. User stays in current OU."
        }
        
        Write-Verbose "Running Get-ADPrincipalGroupMembership on $username"
        $grplist = (Get-ADPrincipalGroupMembership $username)
        ForEach ($grp in $grplist)
        {
            if ($grp.name -notin ($filter))
            {
                Write-Verbose "Removing $username from group: $grp.name"
                Try
                {
                    Remove-ADGroupMember -Identity $grp -Members $username -Confirm:$false -ErrorAction SilentlyContinue
                }
                Catch
                {
                    Write-Warning "While trying to remove $username from group $grp.name, this happened:"
                    Write-Warning $_.Exception.Message
                }
            }
        }

        Write-Verbose "Hiding user from Global Address List"
        Hide-UserFromGAL $username
    }
    
    End
    {
        Write-Output 'All done.'
    }
}


function GetGoodbyeOU($username)
{
    $dn = (Get-ADUser $username).DistinguishedName
    $parts = $dn -split 'OU=Users'
    return 'OU=NoLongerEmployed,OU=Users' + $parts[1..$($parts.Count-1)] -join ','
}

function TryGoodbyeOU($username)
{
    [string]$GoodbyeOU = GetGoodbyeOU($username)
    return [adsi]::Exists("LDAP://$GoodbyeOU")
}

<#
.Synopsis
   Hide user from Global Address List.
.EXAMPLE
   Hide-UserFromGAL ZPamfil
#>
function Hide-UserFromGAL
{
    [CmdletBinding(SupportsShouldProcess=$True)]
    [OutputType([string])]
    Param
    (
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)][string]$username
    )
     
        if ($PSCmdlet.ShouldProcess("$Username", "Set HiddenFromAddressListsEnabled bit to True"))
        {
            Set-Mailbox -Identity $username -HiddenFromAddressListsEnabled $True
        }
       
        if (! $?)
        {
            Write-Warning -Message ("Set-Mailbox failed." +
            ' No mailbox associated with username' +
            ' or you are not running this inside the' +
            ' Exchange Management Shell.')
            Throw "Terminating Error: Set-Mailbox failed."
        }
 }

<#
.Synopsis
   Displays AD thumbnailPhoto for given username 
   in a nice Windows Forms box.
.EXAMPLE
   Get-Photo XKumar
#>
function Get-Photo
{
    [CmdletBinding()]
    [OutputType([byte[]])]
    Param
    (
        # Target AD Username
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Username
    )

  
    Process
    {
        Try
        {
            [byte[]]$thumb = (Get-ADUser $Username -property thumbnailPhoto -EA Stop | Select -ExpandProperty thumbnailPhoto)
        }
        Catch
        {
            Throw "$_.Exception.Message"
        }
        
        Add-Type -AssemblyName System.Windows.Forms

        $img = [System.Drawing.Image]::FromStream([System.IO.MemoryStream]$thumb)
        $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
        
        [System.Windows.Forms.Application]::EnableVisualStyles()
        $pictureBox = new-object Windows.Forms.PictureBox
        $pictureBox.Width =  $img.Width
        $pictureBox.Height =  $img.Height
        $pictureBox.Image = $img

        $form = new-object Windows.Forms.Form
        $form.Text = "thumb"
        $form.Width = $img.Width
        $form.Height =  $img.Height
        $form.AutoSize = $True
        $form.AutoSizeMode = "GrowAndShrink"
        $form.Icon = $icon
        $form.controls.add($pictureBox)
        $form.Add_Shown( { $form.Activate() } )
        $form.ShowDialog()
    }

    End
    {
        $form.Dispose()
        $pictureBox.Dispose()
        $icon.Dispose()
        $img.Dispose()
        Remove-Variable thumb
    }

}

<#
.Synopsis
   Sets thumbnailPhoto AD Property
.DESCRIPTION
   Sets thumbnailPhoto AD Property from file,
   while resizing on the fly to 96x96 pixels.
   Input file can be PNG/JPG/BMP or any other
   image MIME type that Windows can read natively.
.EXAMPLE
   Set-Photo XKumar C:\TMP\XKumar_400px.jpg
#>
function Set-Photo
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        # Target AD Username
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Username,

        # Source image filename with full path
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        $Filename
    )

    Begin
    {
        Try { Get-ADUser $Username -ErrorAction Stop | Out-Null }
        Catch { Throw "No such Username ($Username)." }
        
        (Get-Item $Filename -ErrorAction Stop | Out-Null).Exists
        $Filename_w_path = (Get-Item $Filename).Fullname
        
        $tmpFile = "$env:TEMP" + "\$Username"
    }

    Process
    {
        Add-Type -AssemblyName System.Drawing

        $image = [System.Drawing.Image]::FromFile($Filename_w_path)
        $newWidth = 96
        $newHeight = 96

        #Encoder parameters for image quality
        $JpegQuality = 80
        $myEncoder = [System.Drawing.Imaging.Encoder]::Quality
        $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1) 
        $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter($myEncoder, $JpegQuality)
        $encoderCodec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | ? {$_.MimeType -eq 'image/jpeg'}

        [double]$aspectRatio = $image.Width / $image.Height
        $newBitmap = New-Object System.Drawing.Bitmap $newWidth, $newHeight
        $canvas = [System.Drawing.Graphics]::FromImage($newBitmap)      
        $canvas.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $canvas.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        #Fill canvas
        $canvas.Clear([System.Drawing.Color]::White)
        $offset = [Math]::Round($newWidth - ($newWidth * $aspectRatio))/2
        $canvas.DrawImage($image, $offset, 0, $newWidth * $aspectRatio, $newHeight)
        $newBitmap.Save($tmpFile, $encoderCodec, $($encoderParams))
        
        #set thumbnailPhoto property with byte array
        [byte[]]$thumb = Get-Content $tmpFile -Encoding byte
        Set-ADUser $Username -Replace @{thumbnailPhoto=$thumb}
    } #Process section ends

    End
    {
        $canvas.Dispose()
        $newBitmap.Dispose()
        $image.Dispose()
        $encoderParams.Dispose()
        Remove-Item $tmpFile
        Remove-Variable thumb
    }

}

Export-ModuleMember -Function "*-*"
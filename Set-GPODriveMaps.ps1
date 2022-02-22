<#
.SYNOPSIS
This script is intended to create or modify a GPO with drive maps configured in a matching Excel workbook.
.DESCRIPTION
This script creates or modifies a GPO with drive maps configured in a matching Excel workbook.
The Excel workbook needs to be configured as follows:
- The worksheet holding the configuration needs to be named 'DriveMaps'
- The first row is ignored and may therefore contain headings
- Starting from the second row the columns need to contain the following information:
    1. UNC path pointing to the share to map
    2. Drive letter
    3. Drive label (optional)
    4. Filter (optional, can either be a group name or a distinguished name pointing at an OU)
.PARAMETER GPName
The value provided for this parameter will be used as the GPOs name.
.PARAMETER Domain
The value provided for this parameter will be used as the target domain for the GPO.
If no value is provided the current domain will be used as the target domain.
.PARAMETER Replace
This switch defines which action is used for drive maps, Update or Replace.
.NOTES
Version:    1.8.2
Author:     Mönks, Dominik
.LINK
https://msdn.microsoft.com/en-us/library/cc232619.aspx
.LINK
https://msdn.microsoft.com/en-us/library/cc232618.aspx
#>

param([ValidateNotNullOrEmpty()]
        [string]$GPName,
        [ValidateNotNullOrEmpty()]
        [string]$Domain = (Get-ADDomain).DistinguishedName.Replace(',DC=','.').TrimStart('DC='),
        [bool]$Replace)

#region:Custom functions for XML
function docStart()
{
    $xml.WriteStartDocument()
}

function docEnd()
{
    $xml.WriteEndDocument()
}

function eleStart([string]$name)
{
    $xml.WriteStartElement($name)
}

function eleEnd()
{
    $xml.WriteEndElement()
}

function att([string]$name, [string]$value)
{
    $xml.WriteAttributeString($name, $value)
}
#endregion

$encoding = [Text.UTF8Encoding]::new($false)
$groupRegex = '\w[-\w]*\\[-\w]+'
$ouRegex = '(OU=[-\p{L}\p{N}\s]+)(,OU=[-\p{L}\p{N}\s]+)*(,DC=\w[-\w]*)+'
$outputWidth = 125

# Check for configuration file
Write-Host "Checking if administration file exists:".PadRight($outputWidth) -NoNewline
if (Test-Path "$PSScriptRoot\DriveMaps.xlsx")
{
    Write-Host 'Succeeded' -ForegroundColor Green
    # Check for existing GPO
    Write-Host 'Checking for existing GPO:'.PadRight($outputWidth) -NoNewline
    if (($gpo = Get-GPO $GPName -Server $Domain -ErrorAction SilentlyContinue) -eq $null)
    {
        Write-Host 'Failed' -ForegroundColor Yellow -NoNewline
        Write-Host ', creating GPO'
        $gpo = New-GPO $GPName -Server $Domain
    }
	else
    {
        Write-Host 'Succeeded' -ForegroundColor Green -NoNewline
        Write-Host ', backing up GPO'
        Backup-GPO $GPName -Server $Domain -Path $PSScriptRoot -Comment $([datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')) | Out-Null
    }
    # Check for configuration subfolder in GPO folder
    $GPO_FOLDER = "\\$Domain\SYSVOL\$Domain\Policies\{$($gpo.Id)}"
    $DRIVES_FILE = "$GPO_FOLDER\User\Preferences\Drives\Drives.xml"
    Write-Host 'Checking for existing configuration:'.PadRight($outputWidth) -NoNewline
    if (-not (Test-Path (Split-Path $DRIVES_FILE)))
    {
        Write-Host 'Failed' -ForegroundColor Yellow -NoNewline
        Write-Host ', creating needed subfolders'
        New-Item (Split-Path $DRIVES_FILE) -ItemType Directory -Force | Out-Null
    }
    elseif (Test-Path $DRIVES_FILE)
    {
        Write-Host 'Succeeded' -ForegroundColor Green -NoNewline
        Write-Host ', deleting configuration'
        Remove-Item $DRIVES_FILE -Force
    }
    # Adjust version information
    $INI_FILE = "$GPO_FOLDER\GPT.INI"
    Write-Host 'Checking for existing version information:'.PadRight($outputWidth) -NoNewline
    if (($match = [regex]::Match((Get-Content $INI_FILE), 'version=(\d+)', [Text.RegularExpressions.RegexOptions]::IgnoreCase)).Success)
    {
        Write-Host 'Succeeded' -ForegroundColor Green -NoNewline
        Write-Host ', increasing user version'
        # Convert current version number to HEX
        $HEXver = [Convert]::ToString([int]($match.Groups[1]).Value, 16)
        $HEXverUSR = $HEXver.PadLeft(8, '0').Substring(0, 4)
        $HEXverCMP = $HEXver.PadLeft(8, '0').Substring(4, 4)
        # Convert user part of version number to DEC and increase
        $DECverUSR = [int][Convert]::ToString("0x$HEXverUSR", 10) + 1
        # Convert user part of version number to HEX and combine both parts
        $HEXverUSR = [Convert]::ToString($DECverUSR, 16)
        $HEXver = "$HEXverUSR$HEXverCMP"
        # Convert to DEC and assign
        $version = [Convert]::ToString("0x$HEXver", 10)
    }
    else
    {
        Write-Host 'Failed' -ForegroundColor Yellow -NoNewline
        Write-Host ', setting versions to default value'
        $version = 65537
    }
    Write-Host 'Updating GPO with new version information:'.PadRight($outputWidth) -NoNewline
    $INI_CONTENT = [Collections.Generic.List[string]]::new()
    $INI_CONTENT.Add('[General]')
    $INI_CONTENT.Add("Version=$version")
    $INI_CONTENT.Add("displayName=$GPName")
    [IO.File]::WriteAllLines($INI_FILE, $INI_CONTENT, $encoding)
    Get-ADObject -LDAPFilter "(&(objectClass=groupPolicyContainer)(name={$($gpo.Id)}))" -Server $Domain | Set-ADObject -Replace @{gPCUserExtensionNames='[{00000000-0000-0000-0000-000000000000}{2EA1A81B-48E5-45E9-8BB7-A6E3AC170006}][{5794DAFD-BE60-433F-88A2-1A31939AC01F}{2EA1A81B-48E5-45E9-8BB7-A6E3AC170006}]';versionNumber="$version"}
    Write-Host 'Succeeded' -ForegroundColor Green
    # Create new configuration
    Write-Host 'Starting to write drive maps to configuration file...'
    $excel = New-Object -ComObject Excel.Application
    $excel.Workbooks.Open("$PSScriptRoot\DriveMaps.xlsx") | Out-Null
    $rows = 2
    $xmlcfg = [Xml.XmlWriterSettings]::new()
    $xmlcfg.CloseOutput = $true
    $xmlcfg.Encoding = $encoding
    $xmlcfg.Indent = $true
    $xml = [Xml.XmlWriter]::Create($DRIVES_FILE, $xmlcfg)
    docStart
        eleStart 'Drives'
            att 'clsid' '{8FDDCC1A-0C3C-43cd-A6B4-71A6DF20DA8C}'
            while ($excel.Workbooks['DriveMaps.xlsx'].Worksheets['DriveMaps'].Cells($rows, 1).FormulaR1C1Local -ne '')
            {
                Write-Host ''
                $path = $excel.Workbooks['DriveMaps.xlsx'].Worksheets['DriveMaps'].Cells($rows, 1).FormulaR1C1Local.ToLower()
                $letter = $excel.Workbooks['DriveMaps.xlsx'].Worksheets['DriveMaps'].Cells($rows, 2).FormulaR1C1Local.ToUpper()
                $label = $excel.Workbooks['DriveMaps.xlsx'].Worksheets['DriveMaps'].Cells($rows, 3).FormulaR1C1Local
                $filter = $excel.Workbooks['DriveMaps.xlsx'].Worksheets['DriveMaps'].Cells($rows, 4).FormulaR1C1Local
                $groups = [regex]::Matches($filter, $groupRegex).Value
                $ous = [regex]::Matches($filter, $ouRegex).Value
                Write-Host "Working on drive map '$path':".PadRight($outputWidth) -NoNewline
                if ($path -ne '' -and $letter -ne '')
                {
                    Write-Host 'Succeeded' -ForegroundColor Green -NoNewline
                    Write-Host ', used the following settings:'
                    Write-Host "$(''.PadLeft($outputWidth))Drive letter: $letter"
                    Write-Host "$(''.PadLeft($outputWidth))Drive label: $label"
                    Write-Host "$(''.PadLeft($outputWidth))Groups in filter: $($groups -join ', ')"
                    Write-Host "$(''.PadLeft($outputWidth))OUs in filter: $($ous -join ', ')"
                    eleStart 'Drive'
                        att 'clsid' '{935D1B74-9CB8-4e3c-9914-7DD559B7A417}'
                        att 'name' "${letter}:"
                        att 'status' "${letter}:"
                        if ($Replace)
                        {
                            att 'image' '1'
                        }
                        else
                        {
                            att 'image' '2'
                        }
                        att 'changed' (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
                        att 'uid' "{$((New-Guid).ToString().ToUpper())}"
                        if ($Replace)
                        {
                            att 'removePolicy' '1'
                        }                        
                        att 'userContext' '1'
                        att 'bypassErrors' '1'
                        eleStart 'Properties'
                            if ($Replace)
                            {
                                att 'action' 'R'
                            }
                            else
                            {
                                att 'action' 'U'
                            }
                            att 'thisDrive' 'SHOW'
                            att 'allDrives' 'NOCHANGE'
                            att 'path' $path
                            if ($label -ne '')
                            {
                                att 'label' $label
                            }
                            if ($Replace)
                            {
                                att 'persistent' '1'
                            }
                            else
                            {
                                att 'persistent' '0'
                            }
                            att 'useLetter' '1'
                            att 'letter' $letter
                        eleEnd
                        if ($groups.Count -gt 0 -or $ous.Count -gt 0)
                        {
                            eleStart 'Filters'
                            foreach ($group in $groups)
                            {
                                if (($adobject = Get-ADGroup $group.Split('\')[1] -Server $group.Split('\')[0]) -ne $null)
                                {
                                    eleStart 'FilterGroup'
                                        att 'bool' 'OR'
                                        att 'not' '0'
                                        att 'name' $group.Trim()
                                        att 'sid' $adobject.SID.Value
                                        att 'userContext' '1'
                                        att 'primaryGroup' '0'
                                        att 'localGroup' '0'
                                    eleEnd
                                }
                            }
                            foreach ($ou in $ous)
                            {
                                if (($adobject = Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$ou'") -ne $null)
                                {
                                    eleStart 'FilterOrgUnit'
                                        att 'bool' 'OR'
                                        att 'not' '0'
                                        att 'name' $ou.Trim()
                                        att 'userContext' '1'
                                        att 'directMember' '0'
                                    eleEnd
                                }
                            }
                            eleEnd
                        }
                    eleEnd
                }
                else
                {
                    Write-Host 'Failed' -ForegroundColor Red -NoNewline
                    Write-Host ', missing either path or driver letter'
                }
                $rows++
            }
        eleEnd
    docEnd
    $xml.Close()
    $excel.DisplayAlerts = $false    
    $excel.Workbooks.Close()
    $excel.Quit()
    # Activate drive mappings in administrator context
    if (($gpo | Get-GPPrefRegistryValue -Context Computer -Key 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -ValueName 'EnableLinkedConnections') -eq $null)
    {
        $gpo | Set-GPPrefRegistryValue -Context Computer -Key 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -ValueName 'EnableLinkedConnections' -Value 1 -Type DWord -Action Update | Out-Null
    }
    $gpo.Description = "Configuration file was last edited on $((Get-Item "$PSScriptRoot\DriveMaps.xlsx").LastWriteTime.ToString('yyyy-MM-dd'))."
}
else
{
    Write-Host 'Failed' -ForegroundColor Red
}

<#
    .TITLE
    Connect-MSRemote-PSv5.0.ps1
    
    .SYNOPSIS
    Provides options to connect to 7 Powershell Sessions in Separate Console Windows to MS Online Services
    
    Brad C. Stevens
    brad@bradcstevens.com

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 5.0, June 27th, 2017

    .DESCRIPTION
    This script will connect seven PowerShell sessions in separate console windows 
    to MS Online services passing a single set of credentials to all seven services.
    
    Services connected to are:  
    - MSOnline
    - Exchange Online
    - Exchange Online MFA
    - Skype Online
    - Azure RM
    - SharePoint Online
    - Security & Compliance Center

    .LINK
    - http://bradcstevens.com

    .GITHUB
    https://github.com/bradcstevens

    .NOTES
    Requirements:
    - MSonline PowerShell Modules - Install-Module MsOnline -Force
    - Skype Online PowerShell Modules - Install-Module SkypeOnlineConnector -Force
    - Azure PowerShell Modules - Install-Module Azure -Force
    - SharePoint Online PowerShell Modules 

    Revision History
    ------------------------------------------------------------------------------------
    1.0 Initial community release
    1.1 Added SharePoint Online
    1.2 Added File Tab to Window. 
    1.3 Changed Window Sizing
    1.4 Fixed Bugs/Issues
    1.5 Fixed Bugs/Issues
    1.6 Fixed Bugs/Issues
    1.7 Fixed Bugs/Issues
    1.8 Fixed Bugs/Issues
    1.9 Fixed Bugs/Issues
    2.0 Debugged console window output
    3.0 Added brief introduction informaiton and debugged consle window output
    4.0 Created cleanup function and enhanced logic speed and order of operations with script load and decompile
    4.1 Added Security & Compliance Center
    5.0 Major revision to utilize SyncHash/XML WPF form to present layout and inclusion of Exchange Online MFA
  

#>


$inputXML = @"
<Window x:Class="Connect_MSRemote_PSv5._0.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Connect_MSRemote_PSv5._0"
        mc:Ignorable="d"
        Title="Connect to Microsoft Online Services" Height="400" Width="395">
    <Grid>
        <DockPanel>
            <Menu DockPanel.Dock="Top">
                <MenuItem Header="_File">
                    <MenuItem x:Name="menu_exit" Header="_Exit" />
                </MenuItem>
            </Menu>
        </DockPanel>
        <CheckBox Content="Skype for Business Online" HorizontalAlignment="Left" Margin="10,30,0,0" VerticalAlignment="Top"/>
            <CheckBox x:Name="EOCheck" Content="Exchange Online" HorizontalAlignment="Left" Margin="10,60,0,0" VerticalAlignment="Top"/>
            <CheckBox x:Name="AzureRMCheck" Content="Azure Resource Management" HorizontalAlignment="Left" Margin="10,90,0,0" VerticalAlignment="Top"/>
            <CheckBox x:Name="MSOLCheck" Content="Microsoft Online" HorizontalAlignment="Left" Margin="10,120,0,0" VerticalAlignment="Top"/>
            <CheckBox x:Name="SPOCheck" Content="SharePoint Online" HorizontalAlignment="Left" Margin="10,150,0,0" VerticalAlignment="Top"/>
            <CheckBox x:Name="SCCCheck" Content="Security and Compliance Center" HorizontalAlignment="Left" Margin="10,180,0,0" VerticalAlignment="Top"/>
            <CheckBox x:Name="EOMFACheck" Content="Exchange Online MFA" HorizontalAlignment="Left" Margin="10,210,0,0" VerticalAlignment="Top"/>
            <TextBox x:Name="username" HorizontalAlignment="Left" Margin="118,247,0,0" Width="250" Height="20" TextWrapping="Wrap" Text="global.admin@company.onmicrosoft.com" VerticalAlignment="Top" />
            <TextBlock  HorizontalAlignment="Left" Margin="10,277,0,0" TextWrapping="Wrap" Text="Password" VerticalAlignment="Top"/>
            <PasswordBox x:Name="passwordBox" HorizontalAlignment="Left" Margin="118,277,0,0" Width="250" Height="20" VerticalAlignment="Top"/>
            
            <Button x:Name="connect_button" Content="Connect" HorizontalAlignment="Left"  Margin="118,310,0,0" VerticalAlignment="Top" Width="250"/>
        <TextBlock HorizontalAlignment="Left" Margin="10,247,0,0" TextWrapping="Wrap" Text="Username" VerticalAlignment="Top"/>
        <DockPanel x:Name="StatusBar">
            <StatusBar Margin="0,353,0,0">
                <StatusBarItem Margin="0,-20,0,0">
                    <TextBlock Name="StatusText" Text="Ready"/>
                </StatusBarItem>
            </StatusBar>
        </DockPanel>
    </Grid>
</Window>


  
"@ 
  
    $inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [xml]$XAML = $inputXML
    $reader = (New-Object System.Xml.XmlNodeReader $xaml) 
    $Form=[Windows.Markup.XamlReader]::Load( $reader )
    $xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}
  
    Function MSRemote-Cleanup {
        Start-Sleep -s 10
        $MSOLfile = Test-Path "$env:temp\Connect-MSOL.ps1"
        $EOFile = Test-Path "$env:temp\Connect-EO.ps1"
        $EOMFAFile = Test-Path "$env:temp\Connect-EOMFA.ps1"
        $SPOfile = Test-Path "$env:temp\Connect-SPO.ps1"
        $SCCfile = Test-Path "$env:temp\Connect-SCC.ps1"
        $SfBOfile = Test-Path "$env:temp\Connect-SfBO.ps1"
        $AzureRMfile = Test-Path "$env:temp\Connect-AzureRM.ps1"
        $EOMFAUPNfile = Test-Path "$env:temp\UserPrincipalName.txt"
        $CredsFile = Test-Path "$env:temp\SecureCredentials.xml"
        If($MSOLfile) {
            Remove-Item "$env:temp\Connect-MSOL.ps1"
        }
        If($EOFile) {
            Remove-Item "$env:temp\Connect-EO.ps1"
        }
        If($EOMFAFile) {
            Remove-Item "$env:temp\Connect-EOMFA.ps1"
        }
        If($SPOfile) {
            Remove-Item "$env:temp\Connect-SPO.ps1"
        }
        If($SCCfile) {
            Remove-Item "$env:temp\Connect-SCC.ps1"
        }
        If($SfBOfile) {
            Remove-Item "$env:temp\Connect-SfBO.ps1"
        }
        If($AzureRMfile) {
            Remove-Item "$env:temp\Connect-AzureRM.ps1"
        }
        If ($CredsFile) {
            Remove-Item "$env:temp\SecureCredentials.xml"
        }
        If ($EOMFAUPNfile) {
            Remove-Item "$env:temp\UserPrincipalName.txt"
        }
    
   } 
    Function Connect-EO {
        $Path = "$env:temp" + "\Connect-EO.ps1"
        Set-Content $Path {
        $Credentials = Import-clixml $env:temp\SecureCredentials.xml
        $host.ui.RawUI.WindowTitle = 'Exchange Online'
        $EOConnectionURI = "https://outlook.office365.com/powershell-liveid/"
        $ExchangeOnlineSession = New-PSSession `
            -ConfigurationName Microsoft.Exchange `
            -ConnectionUri $EOConnectionURI `
            -Credential $Credentials `
            -Authentication "Basic" `
            -AllowRedirection
        Import-PSSession $ExchangeOnlineSession
        Clear-Host
        Write-Output `n
        Write-Output "Environment : Exchange Online"
        Write-Output `n
        Write-Output "AcceptedDomains:"
        Get-AcceptedDomain | Format-Table
        
        }
    }

    Function Connect-EOMFA {
        $Path = "$env:temp" + "\Connect-EOMFA.ps1"
        Set-Content $Path {
            $Credentials = Import-clixml $env:temp\SecureCredentials.xml
            $host.ui.RawUI.WindowTitle = 'Exchange Online'
            Function Install-ClickOnce {
            [CmdletBinding()] 
            Param(
                $Manifest = "https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application",
                $ElevatePermissions = $true
            )
                Try { 
                    Add-Type `
                        -AssemblyName System.Deployment
                    Write-Verbose "Start installation of ClockOnce Application $Manifest "
                    $RemoteURI = [URI]::New( $Manifest , [UriKind]::Absolute)
                    If (-not  $Manifest) {
                        throw "Invalid ConnectionUri parameter '$ConnectionUri'"
                    }
                    $HostingManager = New-Object System.Deployment.Application.InPlaceHostingManager `
                        -ArgumentList $RemoteURI , $False
                    Register-ObjectEvent `
                        -InputObject $HostingManager `
                        -EventName GetManifestCompleted `
                        -Action { 
                            new-event -SourceIdentifier "ManifestDownloadComplete"
                        } | 
                        Out-Null
                    Register-ObjectEvent `
                        -InputObject $HostingManager `
                        -EventName DownloadApplicationCompleted `
                        -Action { 
                            New-Event `
                                -SourceIdentifier "DownloadApplicationCompleted"
                         } | 
                         Out-Null
                    $HostingManager.GetManifestAsync()
                    $event = Wait-Event `
                        -SourceIdentifier "ManifestDownloadComplete" `
                        -Timeout 5
                    If ($event ) {
                        $event | 
                        Remove-Event
                        Write-Verbose "ClickOnce Manifest Download Completed"
                        $HostingManager.AssertApplicationRequirements($ElevatePermissions)
                        $HostingManager.DownloadApplicationAsync()
                        $event = Wait-Event `
                            -SourceIdentifier "DownloadApplicationCompleted" `
                            -Timeout 25
                        If ($event ) {
                            $event | 
                            Remove-Event
                            Write-Verbose "ClickOnce Application Download Completed"
                        } 
                        Else {
                            Write-error "ClickOnce Application Download did not complete in time (15s)"
                        }
                    } 
                    Else {
                       Write-error "ClickOnce Manifest Download did not complete in time (5s)"
                    }
                } 
                Finally {
                    Get-EventSubscriber| 
                    ? {$_.SourceObject.ToString() -eq 'System.Deployment.Application.InPlaceHostingManager'} | 
                    Unregister-Event
                }
            }

        Function Get-ClickOnce {
        [CmdletBinding()]  
        Param(
            $ApplicationName = "Microsoft Exchange Online Powershell Module"
        )
            $InstalledApplicationNotMSI = Get-ChildItem HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall | 
            foreach-object {Get-ItemProperty $_.PsPath}
            return $InstalledApplicationNotMSI | 
            ? { $_.displayname -match $ApplicationName } | 
            Select-Object `
                -First 1
        }

        Function Test-ClickOnce {
        [CmdletBinding()] 
        Param(
            $ApplicationName = "Microsoft Exchange Online Powershell Module"
        )
            return ( (Get-ClickOnce -ApplicationName $ApplicationName) -ne $null) 
        }

        Function Uninstall-ClickOnce {
        [CmdletBinding()] 
        Param(
            $ApplicationName = "Microsoft Exchange Online Powershell Module"
        )
            $App = Get-ClickOnce -ApplicationName $ApplicationName
            If ($App) { 
                $selectedUninstallString = $App.UninstallString 
                $parts = $selectedUninstallString.Split(' ', 2)
                Start-Process -FilePath $parts[0] -ArgumentList $parts[1] -Wait 
                $App=Get-ClickOnce -ApplicationName $ApplicationName
                if ($App) {
                    Write-verbose 'De-installation aborted'
                } 
                else {
                    Write-verbose 'De-installation completed'
                } 
            } 
            else {
            }
        }

        Function Load-ExchangeMFAModule { 
        [CmdletBinding()] 
        Param ()
            $Modules = @(Get-ChildItem -Path "$($env:LOCALAPPDATA)\Apps\2.0" -Filter "Microsoft.Exchange.Management.ExoPowershellModule.manifest" -Recurse )
            If ($Modules.Count -ne 1 ) {
                throw "No or Multiple Modules found : Count = $($Modules.Count )"  
            }  
            else {
                $ModuleName =  Join-path $Modules[0].Directory.FullName "Microsoft.Exchange.Management.ExoPowershellModule.dll"
                if ($PSVersionTable.PSVersion -ge "5.0")  { 
                    Import-Module `
                        -FullyQualifiedName $ModuleName `
                        -Force 
                } 
                else { 
                    Import-Module $ModuleName `
                        -Force 
                }

                $ScriptName =  Join-path $Modules[0].Directory.FullName "CreateExoPSSession.ps1"
                if (Test-Path $ScriptName) {
                    $textToAdd = "Clear-Host"
                    $fileContent = Get-Content $ScriptName
                    $fileContent | Where-Object {$_ -notmatch "Clear-Host"} | Set-Content $ScriptName
                    $fileContent = Get-Content $ScriptName
                    $fileContent[$lineNumber-15] += $textToAdd
                    $fileContent | Set-Content $ScriptName
                    return $ScriptName

                } 
                else {
                    throw "Script not found"
                    return $null
                }
            }
        }
        If ((Test-ClickOnce -ApplicationName "Microsoft Exchange Online Powershell Module" ) -eq $false)  {
                Install-ClickOnce -Manifest "https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application"
            }
            
            $Script = Load-ExchangeMFAModule -Verbose
            . $Script
            $UPN = Get-Content $env:temp\UserPrincipalName.txt
            Connect-EXOPSSession -UserPrincipalName $UPN
            Clear-Host
        }
    }


    Function Connect-MSOL {
        $Path = "$env:temp" + "\Connect-MSOL.ps1"
        Set-Content $Path {
            $Credentials = Import-clixml $env:temp\SecureCredentials.xml
            $host.ui.RawUI.WindowTitle = 'MSOL Online'
            Import-Module MsOnline
            Connect-MsolService `
                -Credential $Credentials
            Clear-Host
            $MSOLAccountSKU = Get-MsolAccountSku
            $TenantId = $MSOLAccountSKU.AccountObjectId[0]
            $AccountName = $MSOLAccountSKU.AccountName[0]
            Write-Output `n
            Write-Output  "Environment        : MSOL Online"
            Write-Output `n
            Write-Output  "TenantID           : $TenantId"
            Write-Output  "AccountName        : $AccountName"
            Write-Output `n
            Write-Output  "Domains            : "
            Get-MsolDomain

        }
    }
      
    Function Connect-SfBO{
        $Path = "$env:temp" + "\Connect-SfBO.ps1"
        Set-Content $Path {
        $Credentials = Import-clixml $env:temp\SecureCredentials.xml
        $host.ui.RawUI.WindowTitle = 'Skype for Business Online'
        Import-Module SkypeOnlineConnector
        $SfBOSession = New-CsOnlineSession `
            -Credential $Credentials
        Import-PSSession $SfBOSession
        $CSTenant = Get-CsTenant | Select-Object DisplayName
        Clear-Host
        Write-Output `n 
        Write-Output "Environment     : Skype for Business Online" -Fore Yellow
        Write-Output "CsTenant        : $CSTenant" -Fore Yellow
        }
    }

    Function Connect-AzureRM {
        $Path = "$env:temp" + "\Connect-AzureRM.ps1"
        Set-Content $Path {
        Import-Module Azure
        Login-AzureRmAccount
        }
    }

    Function Connect-SPO {
        $Path = "$env:temp" + "\Connect-SPO.ps1"
        Set-Content $Path {
        $Credentials = Import-clixml $env:temp\SecureCredentials.xml
        $host.ui.RawUI.WindowTitle = 'Sharepoint Online'
        $SPOModule = 'Microsoft.Online.SharePoint.PowerShell'
        $orgName = Read-Host "Please enter the name of your Office 365 organization, example: Contoso"
        $SPOConnectionURL = "https://$orgName-admin.sharepoint.com"
        Import-Module $SPOModule `
            -DisableNameChecking
        Connect-SPOService `
            -Url $SPOConnectionURL `
            -Credential $Credentials
        Clear-Host
        Write-Output `n 
        Write-Output "Environment     : SharePoint Online"
        Write-Output `n
        Get-SPOTenant

        }
    }

    Function Connect-SCC {
        $Path = "$env:temp" + "\Connect-SCC.ps1"
        Set-Content $Path {
        $Credentials = Import-clixml $env:temp\SecureCredentials.xml
        $host.ui.RawUI.WindowTitle = 'Security & Compliance Center'
        $SCCConnectionURI = "https://nam02b.ps.compliance.protection.outlook.com/powershell-liveid?PSVersion=5.1.14393.1358"
        $SCCSession = New-PSSession `
            -ConfigurationName Microsoft.Exchange `
            -ConnectionUri $SCCConnectionURI `
            -Credential $Credentials `
            -Authentication "Basic" `
            -AllowRedirection
        Import-PSSession $SCCSession `
            -AllowClobber `
            -DisableNameChecking
        Clear-Host
        Write-Output `n
        Write-Output "Environment : Security & Compliance Center"       
        }
    }

    Function Connect-MSRemotePS {

        $AccountPassword = $WPFpasswordBox.Password | 
        ConvertTo-SecureString -AsPlainText -Force
        $Username = $WPFusername.text 
        New-object System.Management.Automation.PSCredential ($Username, $AccountPassword) | 
        Export-clixml $env:temp\SecureCredentials.xml | 
        Out-Null

        If ($WPFEOCheck.IsChecked -eq $true) {
            Connect-EO
            Start-Process powershell.exe `
                    -ArgumentList @("-NoExit", "$env:temp\Connect-EO.ps1")
        }
        If ($WPFEOMFACheck.IsChecked -eq $true) {
            set-content $env:temp\UserPrincipalName.txt $Username
            Connect-EOMFA
            Start-Process powershell.exe `
                    -ArgumentList @("-NoExit", "$env:temp\Connect-EOMFA.ps1")
        }
        If ($WPFSfBCheck.IsChecked -eq $true) {
            Connect-SfBO
            Start-Process powershell.exe `
                -ArgumentList @("-NoExit", "$env:temp\Connect-SfBO.ps1")
        }
        If ($WPFMSOLCheck.IsChecked -eq $true) {
            Connect-MSOL
            Start-Process powershell.exe `
                -ArgumentList @("-NoExit", "$env:temp\Connect-MSOL.ps1")
        }
        If ($WPFAzureRMCheck.IsChecked -eq $true) {
            Connect-AzureRM
            Start-Process powershell.exe `
                -ArgumentList @("-NoExit", "$env:temp\Connect-AzureRM.ps1")
        }
        If ($WPFSPOCheck.IsChecked -eq $true) {
            Connect-SPO
            Start-Process powershell.exe `
                -ArgumentList @("-NoExit", "$env:temp\Connect-SPO.ps1")
        }
        If ($WPFSCCCheck.IsChecked -eq $true) {
            Connect-SCC
            Start-Process powershell.exe `
                -ArgumentList @("-NoExit", "$env:temp\Connect-SCC.ps1")
        }
        
        MSRemote-Cleanup
    }

	$WPFconnect_button.Add_Click({
        Connect-MSRemotePS
	})

	$WPFmenu_exit.Add_Click({
		$Form.Close()
	})


$Form.ShowDialog() | out-null

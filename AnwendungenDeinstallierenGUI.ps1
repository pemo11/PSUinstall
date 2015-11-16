<#
 .Synopsis
 Application uinstalling with a friendly GUI
 .Description
 The script uses the information in the registry to show all available applications for uninstallation
 .Notes
 Author: P. Monadjemi, pm@activetraining.de, Blog: posh-admin.de
 Run this script with no parameters
 Last Update: 14/15/2015
#>

Set-StrictMode -Version 2.0

$Script:Meldungen = @()

Import-LocalizedData -BindingVariable Localized -UICulture "de-DE" 

Add-Type –AssemblyName PresentationFramework
Add-Type –AssemblyName PresentationCore
Add-Type –AssemblyName WindowsBase

<#
 .Synopsis
    Registry-Cleaner
 .Description
    Removes Uninstall sub keys that points to not existing exe' oder msi's
#>
function Clear-UnInstallKeys
{
        [CmdletBinding()]
        param()

        $UninstallKey1 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
        $UninstallKey2 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        $UninstallKey3 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
        $UninstallKey4 = "HKCU:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"

        $UnInstallKeys = $UnInstallKey1, $UnInstallKey2, $UnInstallKey3, $UnInstallKey4

        $UnInstallKeys = $UnInstallKeys | Where { Test-Path -Path $_.Substring(0, $_.Length -2 )}

        foreach($Key in $UnInstallKeys)
        {
            Get-ItemProperty -Path $Key | ForEach-Object {
            try
            {
                $SubKey = $_
                $SubKey.UninstallString
                if ($SubKey.UninstallString -notmatch "msiexec")
                {
                    if ($SubKey.UninstallString -match "`"(.+\.exe)")
                    {
                        $Path = $Matches[1]
                        Write-Verbose "Teste $Path..." -Verbose
                        # Gibt es die Datei?
                        if (!(Test-Path -Path $Path))
                        {
                            Write-Warning "$Path existiert nicht."
                            # Schlüssel löschen
                            Remove-Item -Path $SubKey.PSPath  -Verbose
                        }
                    }
                }
            }
            catch
            {
                Write-Warning "$($SubKey.PsPath) besitzt keinen Uinstall-Eintrag"
                # Schlüssel löschen
                Remove-Item -Path $SubKey.PSPath  -Verbose
               }
            }
        }
}

<#
 .Synopsis
    Application uninstall
 .Description
    Uninstalls all selected applications
#>
function UnInstall-Apps
{
    [CmdletBinding()]
    param()

    $AppTitle = $Localized.AppTitle
    $InstalledAppLabel = $Localized.InstalledAppLabel
    $AppListButtonContent = $Localized.AppListButtonContent
    $UninstallAppButtonContent = $Localized.UninstallAppButtonContent
    $DeleteLogFileMessage = $Localized.DeleteLogFileMessage
    $FixRegistryMessage = $Localized.FixRegistryMessage
    $LogfileDeleteConfirmation = $Localized.LogfileDeleteConfirmation
    $RegistryKeyDeleteErrorMessage = $Localized.RegistryKeyDeleteErrorMessage

    $LogPfad = Join-Path -Path $PSScriptRoot -ChildPath "SoftwareUninstaller.log"

    $YesReponseKey = "J"
    if ($PSUICulture -ne "de-De")
    {
        $YesReponseKey = "Y"
    }

    # XAML Window definition
    $XAMLCode= @"
         <Window
            xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            Title="$AppTitle"
            Width="800"
            Height="740"
        >
        <StackPanel>
            <Label
               Margin="4"
               Height="Auto"
               Width="Auto"
               HorizontalAlignment="Center"
               FontSize="20"
               FontWeight="SemiBold"
               Content="$InstalledAppLabel"
             />
            <Button
                x:Name="AnwendungenAuflistenButton"
                Content="$AppListButtonContent"
                Margin="4"
                Width="320"
                Height="32"
             />
             <ListView
               x:Name="AnwendungenListView"
               Margin="4"
               Height="280"
               Width="Auto"
               ItemsSource="{Binding Path=.}"
               IsSynchronizedWithCurrentItem="True"
              >
              <ListView.View>
               <GridView>
                    <GridView.Columns>
                        <GridViewColumn>
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <CheckBox IsChecked="{Binding Path=IsSelected, Mode=TwoWay}" />
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="DisplayName" Width="240" DisplayMemberBinding="{Binding Path=DisplayName}" />
                        <GridViewColumn Header="DisplayVersion" Width="160" DisplayMemberBinding="{Binding Path=DisplayVersion}" />
                    </GridView.Columns>
               </GridView>
              </ListView.View>
             </ListView>
             <Label
                x:Name="AppCountLabel"
                Width="Auto"
                Height="30"
                Margin="4"
                HorizontalAlignment = "Stretch"
                Background = "LightYellow"
                />
            <Button
                x:Name="AnwendungenDeinstallierenButton"
                Content="$UninstallAppButtonContent"
                Margin="4"
                Width="320"
                Height="32"
                />
            <ProgressBar
                x:Name="MainProgressBar"
                Margin="4"
                Height="24"
                Width="Auto"
                />
             <ListView
               x:Name="MeldungenListView"
               Margin="4"
               Height="200"
               Background="LightYellow"
               Width="Auto"
               ItemsSource="{Binding Path=.}"
              >
              <ListView.View>
               <GridView>
                    <GridViewColumn Header="Id" Width="40" DisplayMemberBinding="{Binding Path=MessageId}" />
                    <GridViewColumn Header="Typ" Width="80" DisplayMemberBinding="{Binding Path=Typ}" />
                    <GridViewColumn Header="Message" Width="640" DisplayMemberBinding="{Binding Path=Message}" />
               </GridView>
              </ListView.View>
             </ListView>
        </StackPanel>
    </Window>
"@

     $AnwendungenAuflistenSB = {
        $UninstallKey1 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
        $UninstallKey2 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        $UninstallKey3 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
        $UninstallKey4 = "HKCU:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"

        $UnInstallKeys = $UnInstallKey1, $UnInstallKey2, $UnInstallKey3, $UnInstallKey4

        $UnInstallKeys = $UnInstallKeys | Where { Test-Path -Path $_.Substring(0, $_.Length -2 )}

        $AppListe = @()

        Get-ItemProperty -Path $UnInstallKeys -ErrorAction Ignore  |  ForEach-Object {
            $KeyEntry = $_
            if ($KeyEntry.psobject.Properties -match "DisplayName" -and `                $KeyEntry.psobject.Properties -match "DisplayVersion" -and `
                $KeyEntry.psobject.Properties -match "UninstallString")
            {
                $AppListe += New-Object -TypeName PsObject -Property @{ IsSelected = $false;
                                                                        DisplayName = $KeyEntry.DisplayName;
                                                                        DisplayVersion = $KeyEntry.DisplayVersion;
                                                                        UninstallString = $KeyEntry.UninstallString;
                                                                        KeyPath = $KeyEntry.PSPath
                                                                      }
            }
        }

        $AnwendungenListView.DataContext = $AppListe | Sort-Object -Property DisplayName -Descending:$false

        $Script:Meldungen += New-Object -TypeName PSObject -Property @{ MessageId = $Meldungen.Count + 1;
                                                                        Typ = "Info";
                                                                        Message = ("{0} Anwendungen vorhanden." -f $AppListe.Count)
                                                                      }
        $MeldungenListView.DataContext = $Meldungen
        $MeldungenListView.Dispatcher.Invoke([Action]{}, "Background")

        $AppCountLabel.Content = ("{0} installierte Anwendungen gefunden." -f $AppListe.Count)
    }

    $AnwendungenDeinstallierenSB = {
    
        # Variablen initialisieren
        $AppCount = 0
        $ErrorCount = 0
        $StartZeit = Get-Date

        $AppListe = @($AnwendungenListView.DataContext | Where IsSelected)
        $Anwendungen = @($AppListe | Where IsSelected | Select DisplayName)
        $AnwendungsNamen = $Anwendungen.DisplayName -join ","
        $MainProgressbar.Maximum = $AppListe.Count
        $LogMessage = "$($Anwendungen.Count) Anwendungen werden deinstalliert - ($AnwendungsNamen)."
        [System.Windows.MessageBox]::Show($LogMessage, $AppTitle)

        $LogMessage = "INFO: $LogMessage"
        Add-Content -Path $LogPfad -Value $LogMessage
        foreach($App in $AppListe)
        {
            $DisplayName = $App.DisplayName
            $DisplayVersion = $App.DisplayVersion
            $UnInstallpfad = $App.UninstallString
            $KeyPfad = $App.KeyPath

            $DeleteUninstallKey = $false

            # Eventuell laufende Msiexec-Prozesse beenden
            $MsiexecTerminated = @(Get-Process -Name Msiexec -ErrorAction Ignore  | Stop-Process -PassThru -Force -ErrorAction Ignore)
            if ($?)
            {
                $LogMessage = "$($MsiexecTerminated.Count) Msiexec-Prozesse wurden beendet."
                $Script:Meldungen += New-Object -TypeName PSObject -Property @{ MessageId = $Meldungen.Count + 1;
                                                                                Typ = "Info";
                                                                                Message = $LogMessage
                                                                              }
            }
            else
            {
                $LogMessage = "$((Get-Process -Name Msiexec).Count) Msiexec-Prozesse konnten nicht beendet werden."
                $Script:Meldungen += New-Object -TypeName PSObject -Property @{ MessageId = $Meldungen.Count + 1;
                                                                                Typ = "ERROR";
                                                                                Message = $LogMessage
                                                                              }

            }
            $MeldungenListView.DataContext = $Meldungen
            $MeldungenListView.Dispatcher.Invoke([Action]{}, "Background")

            Write-Verbose -Message $LogMessage -Verbose
            $LogMessage = "INFO: $LogMessage"
            Add-Content -Path $LogPfad -Value $LogMessage

            # Jetzt wird es ernst
            try
            {
                $LogMessage = "$DisplayName wird in der Version $DisplayVersion deinstalliert."
                $Script:Meldungen += New-Object -TypeName PSObject -Property @{ MessageId = $Meldungen.Count + 1;
                                                                                Typ = "Info";
                                                                                Message = $LogMessage
                                                                              }
                $MeldungenListView.SelectedIndex = $MeldungenListView.Items.Count - 1
                $MeldungenListView.ScrollIntoView($MeldungenListView.SelectedItem)
                $MeldungenListView.Dispatcher.Invoke([Action]{}, "Background")
    
                Write-Verbose -Message $LogMessage -Verbose
                $LogMessage = "INFO: $LogMessage"
                Add-Content -Path $LogPfad -Value $LogMessage
    
               # Ist es Msiexec?
                if ($UnInstallpfad -match "(msiexec.exe)(.*)")
                {
                        $MsiArgs = $Matches[2].Trim()
                        # Option /qn anhängen
                        if ($MsiArgs -notlike "*/qn*")
                        {
                            $MsiArgs += " /qn"
                        }
                        $P = Start-Process -FilePath MsiExec.exe -ArgumentList $MsiArgs -Wait -PassThru
                        # Wichtig: Abfrage des Msi-Returncodes
                        switch ($P.ExitCode)
                        {
                            0 {
                                $LogMessage = "$DisplayName wurde in der Version $DisplayVersion deinstalliert."
                                $AppCount++
                                $DeleteUninstallKey = $true
                            }
                            1602 {
                                $LogMessage = "$DisplayName wurde nicht deinstalliert - Abbruch durch Benutzer"
                                $ErrorCount ++
                            }
                            1604 {
                                $LogMessage = "$DisplayName wurde nicht deinstalliert - vorzeitiger Abbruch"
                                $ErrorCount ++
                            }
                            1605 {
                                $LogMessage = "$DisplayName wurde nicht deinstalliert - vorzeitiger Abbruch"
                                $ErrorCount ++
                            }
                            1608 {
                                $LogMessage = "$DisplayName wurde nicht deinstalliert - es läuft bereits eine Deinstallation"
                                $ErrorCount ++
                            }
                            1609 {
                                $LogMessage = "$DisplayName wurde nicht deinstalliert - Fehler bei der Deinstallation"
                                $ErrorCount ++
                                }
                           default {
                                $LogMessage = "$DisplayName - Spezieller Fehler bei der Deinstallation (" + $P.ExitCode + ")"
                                $ErrorCount ++
                            }
                        }
                    }
                    else
                    {
                        $P = Start-Process -FilePath $UnInstallPfad -Wait -PassThru
                        # Auch bei diesem Aufruf wird der Exit-Code über das Process-Objekt abgefragt
                        if ($P.ExitCode -eq 0)
                        {
                            $LogMessage = "$DisplayName wurde in der Version $DisplayVersion deinstalliert."
                            $Script:Meldungen += New-Object -TypeName PSObject -Property @{ MessageId = $Meldungen.Count + 1;
                                                                                            Typ = "Info";
                                                                                            Message = $LogMessage
                                                                                   }
                            Write-Verbose -Message $LogMessage -Verbose
                            $LogMessage = "INFO: $LogMessage"
                            Add-Content -Path $LogPfad -Value $LogMessage
                            $AppCount++
                            $DeleteUninstallKey = $true
                        }
                        else
                        {
                            $LogMessage = "Fehler beim Deinstallieren von $DisplayName  - Exitcode: $($P.ExitCode)"
                            Write-Warning -Message $LogMessage
                            $Script:Meldungen += New-Object -TypeName PSObject -Property @{ MessageId = $Meldungen.Count + 1;
                                                                                            Typ = "Error";
                                                                                            Message = $LogMessage
                                                                                          }
                            $LogMessage = "ERROR: $LogMessage"
                            Add-Content -Path $LogPfad -Value $LogMessage
                        }
                    }
                    if ($DeleteUninstallKey)
                    {
    
                        $Script:Meldungen += New-Object -TypeName PSObject -Property @{ MessageId = $Meldungen.Count + 1;
                                                                                        Typ = "Info";
                                                                                        Message = $LogMessage
                                                                                      }
                        Write-Verbose -Message $LogMessage -Verbose
                        $LogMessage = "INFO: $LogMessage"
                        Add-Content -Path $LogPfad -Value $LogMessage
                    }
                    else
                    {
                        $Script:Meldungen += New-Object -TypeName PSObject -Property @{ MessageId = $Meldungen.Count + 1;
                                                                                        Typ = "Error";
                                                                                        Message = $LogMessage
                                                                                      }
                        Write-Verbose -Message $LogMessage -Verbose
                        $LogMessage = "ERROR: $LogMessage"
                        Add-Content -Path $LogPfad -Value $LogMessage
                    }
            }
            catch
            {
                $LogMessage = "Fehler beim Deinstallieren von $DisplayName ($_)"
                Write-Warning -Message $LogMessage
                $LogMessage = "ERROR: $LogMessage"
                Add-Content -Path $LogPfad -Value $LogMessage
                $Script:Meldungen += New-Object -TypeName PSObject -Property @{ MessageId = $Meldungen.Count + 1;
                                                                                Typ = "Error";
                                                                                Message = $LogMessage
                                                                              }
            }
   
            # delete Uninstall key?
            if ($DeleteUninstallKey)
            {
                $DeleteUninstallKey = $false
                try
                {
                    Remove-Item -Path "$KeyPfad" -ErrorAction Stop
                    $LogMessage = "$KeyPfad $($Localized.LogfileDeleteConfirmation)"
                    $Script:Meldungen += New-Object -TypeName PSObject -Property @{ MessageId = $Meldungen.Count + 1;
                                                                                    Typ = "Info";
                                                                                    Message = $LogMessage
                                                                                  } 
                    Write-Verbose -Message $LogMessage -Verbose
                    $LogMessage = "INFO: $LogMessage"
                    Add-Content -Path $LogPfad -Value $LogMessage
                }
                catch
                {
                    $LogMessage = "$($Localized.RegistryKeyDeleteErrorMessage) $KeyPfad ($_)"
                    $Script:Meldungen += New-Object -TypeName PSObject -Property @{ MessageId = $Meldungen.Count + 1;
                                                                                    Typ = "Error";
                                                                                    Message = $LogMessage
                                                                                  } 
                    Write-Verbose -Message $LogMessage -Verbose
                    $LogMessage = "ERROR: $LogMessage"
                    Add-Content -Path $LogPfad -Value $LogMessage
                }

            }
            $MainProgressbar.Value++
            $MeldungenListView.DataContext = $Meldungen
            $MeldungenListView.Dispatcher.Invoke([Action]{}, "Background")
            $MeldungenListView.SelectedIndex = $MeldungenListView.Items.Count - 1
            $MeldungenListView.ScrollIntoView($MeldungenListView.SelectedItem)
        }

        $MainProgressbar.Value = 0

        $Dauer = (Get-Date) - $StartZeit
        $LogMessage = "{0} von {1} Anwendung(en) wurden in {2} Stunden, {3} Minuten und {4} Sekunden deinstalliert - es gab {5} Fehler." `         -f $AppCount, $AppListe.Count, $Dauer.Hours, $Dauer.Minutes, $Dauer.Seconds, $ErrorCount
        $Script:Meldungen += New-Object -TypeName PSObject -Property @{ MessageId = $Meldungen.Count + 1;
                                                                        Typ = "Info";
                                                                        Message = $LogMessage
                                                                      }
        Write-Verbose -Message $LogMessage -Verbose
        $LogMessage = "INFO: $LogMessage"
        Add-Content -Path $LogPfad -Value $LogMessage
        $MeldungenListView.DataContext = $Meldungen
        $MeldungenListView.Dispatcher.Invoke([Action]{}, "Background")
        $MeldungenListView.SelectedIndex = $MeldungenListView.Items.Count - 1
        $MeldungenListView.ScrollIntoView($MeldungenListView.SelectedItem)
    }

    # ***** Execution starts here *****

    # Abfrage, ob Uninstall-Einträge entfernt werden, die nicht verwendbar sind
    $Response = Read-Host -Prompt "$FixRegistryMessage ?"
    if ($Response -like "$YesReponseKey*")
    {
        Clear-UnInstallKeys
    }

    # Abfrage, ob Log-Datei gelöscht oder beibehalten wird

    if (Test-Path -Path $LogPfad)
    {
        $Response = Read-Host -Prompt "$DeleteLogFileMessage ?"
        if ($Response -like "$YesReponseKey*")
        {
            del -Path $LogPfad
            Write-Verbose -Message "$LogPfad $LogfileDeleteConfirmation."
        }
    }

    # Display the WPF window

    $MainWin = [System.Windows.Markup.XamlReader]::Parse($XamlCode)
    $AnwendungenListView = $MainWin.FindName("AnwendungenListView")
    $MeldungenListView = $MainWin.FindName("MeldungenListView")

    $AnwendungenAuflistenButton = $MainWin.FindName("AnwendungenAuflistenButton")
    $AnwendungenAuflistenButton.add_Click($AnwendungenAuflistenSB)

    $AnwendungenDeinstallierenButton = $MainWin.FindName("AnwendungenDeinstallierenButton")
    $AnwendungenDeinstallierenButton.add_Click($AnwendungenDeinstallierenSB)
    $MainProgressbar = $MainWin.FindName("MainProgressBar")

    $AppCountLabel = $MainWin.FindName("AppCountLabel")

    $MainWin.ShowDialog() | Out-Null

    # Finally terminate all running Msiexec processes if any
    Get-Process -Name Msciexec -ErrorAction Ignore | Stop-Process -Force -ErrorAction Ignore
}
#REQUIRES -Version 5 -Modules @{ModuleName="ExchangeOnlineManagement"; ModuleVersion="0.4368.1"}
<#
.SYNOPSIS
Performs a message trace for a spam message and searches for and purges it from recipient mailboxes

.DESCRIPTION
Connects to O365 - assuming you have Exchange Admin permissions - and performs a Get-MessageTrace.
From there, it will take all of the recipients, split them up separate Content Searches with the
action of purging all items.

.PARAMETER NoGUI
OPTIONAL This script detects the Windows Presentation Framework and attempts to make a GUI from it, use this
switch to force command line interactive prompts instead
.PARAMETER EmailDomain
REQUIRED Domain of the mailboxes affected by spam campaign (used for message trace search/filter). Defaults to
contoso.com and will prompt if you don't supply yours.
.PARAMETER NoMFA
OPTIONAL If the account you wish to use is enabled for basic auth and doesn't have expiring tokens, use this
switch to operate without MFA; otherwise, it will expect to use MFA and modern exchange.
.PARAMETER AccountLockdownScriptName
OPTIONAL If you wish to run a script to lockdown sender(s) that are part of the EmailDomain as a part of this
process (and your script has a $Users string parameter that can split based on commas), supply the full name of the script (e.g. Secure-Account-Manually.ps1) and add the script to the same folder as this one
.PARAMETER Mailboxes2Exclude
OPTIONAL If there are certain mailbox address to exclude (e.g. an on-prem mailbox that cannot be managed by the
O365 Compliance and Security Center), supply them to this switch to ignore them in the attempts to fix things.

.NOTES
    Created by: Brendan Horner (www.hornerit.com)
    Notes: MUST BE RUN AS SCRIPT FILE, do NOT copy-paste into PS to run
    Version History:
    --2020-07-16-Added timestamp to content search name so that there wouldn't be duplicates
    --2020-06-22-New version to use modern, public exchange v2 cmdlets and use Content Search
    --2019-08-06-Bug fix for time buttons and added credential to window title for easier identification
    --2019-07-16-Added 2 new features - buttons for recent times and filters for email status
    --2019-07-15-Fixed bug for child windows again for MFA parameter
    --2019-06-27-Bugfix for MFA Module loading to use LastWriteTime instead of Modified
    --2019-06-19-Altered MFA parameter to be NoMFA so someone can force basic auth by setting that switch and
        adjusted MFA module to pull the latest version of the module on your machine. Bug fix from anon commenter.
    --2019-05-28-Bug fixes for MFA and mailbox errors
    --2019-05-21-Added Exchange Admin prompt to gui
    --2019-05-16-Added GUI, fixed a few minor display bugs and performance bugs, rewrote sections for dynamic
        window generation based on params, allows as many Exch Admin accts to assist as you can try...
        watch out for RAM usage
    --2019-05-02-Added better logic for throttling
    --2019-04-15-Initial public version
.EXAMPLE
.\O365-SPAM-REMOVER.ps1 -NoMFA
.\O365-SPAM-REMOVER.ps1 -NoMFA -AccountLockdownScriptName "Secure-Account.ps1"
#>
[CmdletBinding()]
param(
    [switch]$NoGUI,
    [string]$EmailDomain = "contoso.com",
    [switch]$NoMFA,
    [string]$AccountLockdownScriptName,
    [string[]]$Mailboxes2Exclude
)

#Try to get the Exchange Online Powershell module that supports MFA
Import-Module ExchangeOnlineManagement
do {
    $Good = 0
    if ($NoMFA) {
        try {
            $Cred = Get-Credential -Message "Please enter exchange admin EMAIL ADDRESS...EMAIL" -ErrorAction Stop
            if ($Cred.Password.Length -eq 0) {
                throw
            }
        } catch {
            Write-Host "Error with your credential input, please try again or re-run without the -NoMFA switch"
            Continue
        }
        try {
            Connect-ExchangeOnline -Credential $Cred -ShowBanner:$False -ErrorAction Stop
        } catch {
            Read-Host "Error connecting to Exchange Online - $_. Press any key to exit."
            Exit
        }
        try {
            Get-OrganizationConfig | Select-Object Name
            $Good = 1
            $CredUPN = $Cred.UserName
        } catch {
            Write-Host "Supplied credential is not an Exchange Admin."
            Disconnect-ExchangeOnline -confirm:$false
        }
    } else {
        $CredUPN = Read-Host "Please enter an Exchange Admin email address"
        if ($CredUPN -match "^.+@.+\..+$") {
            try {
                Connect-ExchangeOnline -UserPrincipalName $CredUPN -ShowBanner:$False -ErrorAction Stop
            } catch {
                Read-Host "Error connecting to Exchange Online - $_. Press any key to exit."
                Exit
            }
            try {
                Get-OrganizationConfig | Select-Object Name
                $Good = 1
            } catch {
                Write-Host "Supplied credential is not an Exchange Admin."
                Disconnect-ExchangeOnline -confirm:$false
            }
        } else {
            Write-Host "Did not supply an email address, try again."
        }
    }
} until ($Good -eq 1)

###BELOW IF THE CODE NECESSARY FOR THE GUI...I KNOW IT LOOKS CRAZY...KINDA IGNORE IT MAYBE
try { 
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
    $GUI = $true
} catch {
    $GUI = $false
}
if ($GUI -and !($NoGUI)) {
    function Get-GUIData{
        [CmdletBinding()]
        param(
            [switch]
            $NoMFA
        )
        $TodayMinus11 = (Get-Date).AddDays(-10).ToShortDateString()
        $TodayPlus2 = (Get-Date).AddDays(2).ToShortDateString()
        $TodayPlus1=(Get-Date).AddDays(1).ToShortDateString()
        $StartDateTime = $null
        $EndDateTime = $null
        [xml]$Xaml=@"
            <Window
                xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                x:Name="Window"
                Title="SPAM Removal Script GUI"
                WindowStartupLocation="CenterScreen"
                SizeToContent="WidthAndHeight"
                ShowInTaskbar="True"
                ScrollViewer.VerticalScrollBarVisibility="Auto">
                <Window.Resources>
                    <Style x:Key="ButtonRoundedCorners" TargetType="Button">
                        <Setter Property="Template">
                            <Setter.Value>
                                <ControlTemplate TargetType="Button">
                                    <Grid>
                                        <Border x:Name="border" CornerRadius="5" BorderBrush="#707070"
                                            BorderThickness="1" Background="LightGray">
                                            <ContentPresenter HorizontalAlignment="Center"
                                                    VerticalAlignment="Center"
                                                    TextElement.FontWeight="Normal">
                                            </ContentPresenter>
                                        </Border>
                                    </Grid>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsMouseOver" Value="True">
                                            <Setter Property="Background" TargetName="border" Value="#BEE6FD"/>
                                            <Setter Property="BorderBrush" TargetName="border" Value="#3C7FB1"/>
                                        </Trigger>
                                        <Trigger Property="IsPressed" Value="True">
                                            <Setter Property="BorderBrush" TargetName="border" Value="#2C628B"/>
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                    </Style>
                    <Style TargetType="Button" BasedOn="{StaticResource ButtonRoundedCorners}"></Style>
                </Window.Resources>
                <StackPanel Orientation="Vertical" Height="Auto" VerticalAlignment="top" Margin="10">
                    <StackPanel Orientation="Horizontal">
                        <StackPanel Orientation="Vertical" VerticalAlignment="Center" Margin="0,0,10,0">
                            <Label Margin="0,0,0,-10">[Required]</Label>
                            <Label>Enter spam email address, click "--&#62;"</Label>
                            <TextBox x:Name="inputEvilSenderTxt" Height="25" Width="200"/>
                        </StackPanel>
                        <StackPanel Orientation="Vertical" VerticalAlignment="Center">
                            <Button x:Name="btnAddEvilSender" Content="--&#62;" Margin="0,0,10,0" Height="25"
                                Width="40"/>
                            <Button x:Name="btnRemoveEvilSender" Content="&#60;--" Margin="0,10,10,0" Height="25"
                                Width="40"/>
                        </StackPanel>
                        <StackPanel Orientation="Vertical">
                            <Label>$(if ($AccountLockdownScriptName.length -gt 0 -and
                                (test-path $PSScriptRoot\$AccountLockdownScriptName)) {
                                    'SPAM Senders - click &#9745; to hammer acct too'
                                } else {
                                    'SPAM Senders'
                                })</Label>
                            <ListBox x:Name="listboxEvilSenders" Height ="150" Width="250"
                                SelectionMode="Extended" ScrollViewer.VerticalScrollBarVisibility="Auto"/>
                        </StackPanel>
                    </StackPanel>
                    <Separator Margin="0,10"/>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Left" VerticalAlignment="Bottom"
                        Margin="10,0,0,0">
                        <Button x:Name="btn24h" Content="Last 24h" Margin="0,0,20,0"/>
                        <Button x:Name="btn48h" Content="Last 48h" Margin="0,0,20,0"/>
                        <Button x:Name="btn72h" Content="Last 72h" Margin="0,0,20,0"/>
                        <StackPanel x:Name="stackPanelRadioEmailStatus" Orientation="Horizontal">
                            <RadioButton GroupName="radioEmailStatus" Content="Delivered" x:Name="Delivered"
                                Margin="20,0,0,0"/>
                            <RadioButton GroupName="radioEmailStatus" Content="FilteredAsSpam"
                                x:Name="FilteredAsSpam" Margin="20,0,0,0"/>
                            <RadioButton GroupName="radioEmailStatus" Content="Both" x:Name="Both"
                                Margin="20,0,0,0" IsChecked="True"/>
                            <TextBox x:Name="textRadioEmailStatus" Visibility="Hidden" Height="1" Width="1"
                                Margin="20,0,0,0" Text="Both"/>
                        </StackPanel>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal">
                        <StackPanel Orientation="Vertical" HorizontalAlignment="Left" VerticalAlignment="Bottom">
                            <Label Margin="0,0,0,-10">[Required]</Label>
                            <Label>Local START time of spam campaign</Label>
                            <DatePicker x:Name="InputStartDate" Margin="5,0,0,0">
                                <DatePicker.BlackoutDates>
                                    <CalendarDateRange Start="01/01/1200" End="$($TodayMinus11)" />
                                    <CalendarDateRange Start="$($TodayPlus1)" End="01/01/2099" />
                                </DatePicker.BlackoutDates>
                            </DatePicker>
                        </StackPanel>
                        <StackPanel Orientation="Vertical" Margin="10,0,1,0" VerticalAlignment="Bottom">
                            <Label HorizontalAlignment="Center">Hour</Label>
                            <ComboBox x:Name="HourDropdownStart" Height="25" Width="50"/>
                        </StackPanel>
                        <StackPanel Orientation="Vertical" HorizontalAlignment="Center"
                            VerticalAlignment="Bottom">
                            <Label HorizontalAlignment="Center">AM/PM</Label>
                            <ComboBox x:Name="AmPmDropdownStart" Height="25" Width="50"/>
                        </StackPanel>
                        <Label x:Name="lblStartDateWarnings" VerticalAlignment="Bottom" FontSize="25"
                            Foreground="red" Margin="0,0,0,-8"/>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                        <StackPanel Orientation="Vertical" HorizontalAlignment="Left" VerticalAlignment="Bottom">
                            <Label Margin="0,0,0,-10">[Required]</Label>
                            <Label>Local END time of spam campaign</Label>
                            <DatePicker x:Name="InputEndDate" Margin="5,0,0,0">
                                <DatePicker.BlackoutDates>
                                    <CalendarDateRange Start="01/01/1200" End="$($TodayMinus11)" />
                                    <CalendarDateRange Start="$($TodayPlus2)" End="01/01/2099" />
                                </DatePicker.BlackoutDates>
                            </DatePicker>
                        </StackPanel>
                        <StackPanel Orientation="Vertical" Margin="10,0,1,0" VerticalAlignment="Bottom">
                            <Label HorizontalAlignment="Center">Hour</Label>
                            <ComboBox x:Name="HourDropdownEnd" Height="25" Width="50"/>
                        </StackPanel>
                        <StackPanel Orientation="Vertical" VerticalAlignment="Bottom">
                            <Label HorizontalAlignment="Center">AM/PM</Label>
                            <ComboBox x:Name="AmPmDropdownEnd" Height="25" Width="50"/>
                        </StackPanel>
                        <Label x:Name="lblEndDateWarnings" VerticalAlignment="Bottom" FontSize="25"
                            Foreground="red" Margin="0,0,0,-8"/>
                    </StackPanel>
                    <Separator Margin="0,10"/>
                    <StackPanel Orientation="Horizontal">
                        <StackPanel Orientation="Vertical" VerticalAlignment="Center">
                            <Label Margin="0,0,0,-10">[Optional]</Label>
                            <Label>Enter subject lines, click "--&#62;"</Label>
                            <TextBox x:Name="inputSubjectLinesTxt" Width="200" Height="25" Margin="0,0,10,0"/>
                        </StackPanel>
                        <StackPanel Orientation="Vertical" VerticalAlignment="Center">
                            <Button x:Name="btnAddSubject" Content="--&#62;" Margin="0,0,10,0" Height="25"
                                Width="40"/>
                            <Button x:Name="btnRemoveSubject" Content="&#60;--" Margin="0,10,10,0" Height="25"
                                Width="40"/>
                        </StackPanel>
                        <StackPanel Orientation="Vertical">
                            <Label>SPAM Email Subject Lines</Label>
                            <ListBox x:Name="listboxEvilSubjects" Height = "150" Width="250"
                                SelectionMode="Extended" ScrollViewer.VerticalScrollBarVisibility="Auto"
                                DisplayMemberPath="DisplayName"/>
                        </StackPanel>
                    </StackPanel>
                    <Button x:Name="btnSubmit" Content="[Disabled]" Margin="20" FontSize="30" IsEnabled="false"/>
                </StackPanel>
            </Window>
"@ -replace 'x:N','N'
        $Reader = (New-Object System.Xml.XmlNodeReader $Xaml)
        $Window = [Windows.Markup.XamlReader]::Load($Reader)

        <#Powershell variables for the controls in the GUI, this first part loads all objects from the above xaml
            and creates powershell variables for each one with a name.#>
        $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
            try {
                Set-Variable -Name $_.Name -Value $window.FindName($_.Name) -ErrorAction Stop
            } catch {
                throw
            }
        }

        #Event handler for the radio buttons, gets added to StackPanel that holds the radio buttons as an event
        [System.Windows.RoutedEventHandler]$Script:CheckedEventHandler = {
            $textRadioEmailStatus.Text = $_.source.name
        }
        $stackPanelRadioEmailStatus.AddHandler(
            [System.Windows.Controls.RadioButton]::CheckedEvent,
            $CheckedEventHandler
        )

        #Functions attached to different controls in the GUI or run from using different controls
        function Resolve-DateInputs {
            $InvalidStart=0
            $InvalidEnd=0
            $Valid = $false
            $StartDateTime = if ($null -ne $InputStartDate.SelectedDate) {
                Get-Date $InputStartDate.SelectedDate
            } else {
                $null
            }
            $EndDateTime = if ($null -ne $InputEndDate.SelectedDate) {
                Get-Date $InputEndDate.SelectedDate
            } else {
                $null
            }
            if ($HourDropdownStart.SelectedIndex -eq -1 -and
                $AmPmDropdownStart.SelectedIndex -eq -1 -and
                $null -ne $InputStartDate.SelectedDate) {
                    $HourDropdownStart.SelectedIndex= 11
                    $AmPmDropdownStart.SelectedIndex= 0
            }
            if ($HourDropdownEnd.SelectedIndex -eq -1 -and
                $AmPmDropdownEnd.SelectedIndex -eq -1 -and
                $null -ne $InputEndDate.SelectedDate) {
                    $HourDropdownEnd.SelectedIndex= 11
                    $AmPmDropdownEnd.SelectedIndex= 0
            }
            if ($null -ne $StartDateTime -and
                $HourDropdownStart.SelectedIndex -ne -1 -and
                !($HourDropdownStart.SelectedIndex -eq 11 -and $AmPmDropdownStart -eq 0) -and
                $AmPmDropdownStart.SelectedIndex -ne -1) {
                    $Hours2Add = if ($AmPmDropdownStart.SelectedValue -eq "PM" -and
                        $HourDropdownStart.SelectedIndex -ne 11) {
                            [int]$HourDropdownStart.SelectedValue + 12
                        } elseif ($AmPmDropdownStart.SelectedIndex -eq 0 -and
                            $HourDropdownStart.SelectedValue -eq 12) {
                                0
                        } else {
                            [int]$HourDropdownStart.SelectedValue
                        }
                    $StartDateTime = (Get-Date $InputStartDate.SelectedDate).AddHours($Hours2Add)
            }
            if ($null -ne $EndDateTime -and
                $HourDropdownEnd.SelectedIndex -ne -1 -and
                !($HourDropdownEnd.SelectedIndex -eq 11 -and $AmPmDropdownEnd -eq 0) -and
                $AmPmDropdownEnd.SelectedIndex -ne -1) {
                    $Hours2Add = if ($AmPmDropdownEnd.SelectedValue -eq "PM" -and
                        $HourDropdownEnd.SelectedIndex -ne 11) {
                            [int]$HourDropdownEnd.SelectedValue + 12 
                        } elseif ($AmPmDropdownEnd.SelectedIndex -eq 0 -and
                            $HourDropdownEnd.SelectedValue -eq 12) {
                                0
                        } else {
                            [int]$HourDropdownEnd.SelectedValue
                        }
                $EndDateTime = (Get-Date $InputEndDate.SelectedDate).AddHours($Hours2Add)
            }
            if ($null -ne $StartDateTime -and $StartDateTime -gt (Get-Date)) {
                $InvalidStart=1
            }
            if ($null -ne $EndDateTime -and $EndDateTime -gt (Get-Date).AddDays(1)) {
                $InvalidEnd=1
            }
            if ($null -ne $StartDateTime -and $null -ne $EndDateTime -and $StartDateTime -ge $EndDateTime) {
                $InvalidStart=1
                $InvalidEnd=1
            }
            if ($InvalidStart -eq 1) {
                $LblStartDateWarnings.Content="INVALID DATE"
            } else {
                $LblStartDateWarnings.Content=""
            }
            if ($InvalidEnd -eq 1) {
                $LblEndDateWarnings.Content="INVALID DATE"
            } else {
                $LblEndDateWarnings.Content=""
            }
            if ($InvalidStart -eq 0 -and
                $InvalidEnd -eq 0 -and
                $null -ne $StartDateTime -and
                $null -ne $EndDateTime) {
                    $Valid = $true
            }
            return [PSCustomObject]@{
                Valid = $Valid
                StartDateTime = $StartDateTime
                EndDateTime = $EndDateTime
            }
        }
        function Resolve-RequiredInputs {
            if ($ListboxEvilSenders.HasItems -eq $true -and ((Resolve-DateInputs).Valid)) {
                $BtnSubmit.Content="Begin SPAM Cleanup"
                $BtnSubmit.IsEnabled = $true
            } else {
                $BtnSubmit.Content="[Disabled]"
                $BtnSubmit.IsEnabled = $false
            }
        }
        function Add-EvilSenders {
            if ((-NOT [string]::IsNullOrEmpty($InputEvilSenderTxt.text))) {
                if ($InputEvilSenderTxt.text -match "^.+@.+\..+$") {
                    $EvilSendersCount = ($ListboxEvilSenders.Items |
                        Where-Object { $_ -like ($InputEvilSenderTxt.text).Trim() }).count
                    $EvilSendersCount2 = ($ListboxEvilSenders.Items |
                        Where-Object { $_.content -like ($InputEvilSenderTxt.text).Trim()}).count
                    if ($EvilSendersCount -eq 0 -and $EvilSendersCount2 -eq 0) {
                        if ($InputEvilSenderTxt.text -match "^.+@$($EmailDomain.replace(".","\."))$" -and
                            $AccountLockdownScriptName.Length -gt 0 -and
                            (Test-Path "$PSScriptRoot\$AccountLockdownScriptName")) {
                                $myChk = New-Object -TypeName System.Windows.Controls.Checkbox
                                $myChk.Content=($InputEvilSenderTxt.text).Trim()
                                $ListboxEvilSenders.Items.Add($myChk)
                        } else {
                            $ListboxEvilSenders.Items.Add(($InputEvilSenderTxt.text).Trim())
                        }
                        $InputEvilSenderTxt.Clear()
                    }
                } else {
                    $InputEvilSenderTxt.Text="NOSOUP4U"
                }
                $InputEvilSenderTxt.Focus()
                Resolve-RequiredInputs
            }
        }
        function Add-EvilSubjects {
            if ((-NOT [string]::IsNullOrEmpty($InputSubjectLinesTxt.text))) {
                $SubjectsObject = [pscustomobject]@{
                    DisplayName='"'+$InputSubjectLinesTxt.text+'"'
                    ValueSubject=$InputSubjectLinesTxt.text
                }
                $ListboxEvilSubjects.Items.Add($SubjectsObject)
                $InputSubjectLinesTxt.Clear()
                $InputSubjectLinesTxt.Focus()
            }
        }
        $BtnAddEvilSender.Add_Click({
            Add-EvilSenders
        })
        $BtnRemoveEvilSender.Add_Click({
            while($ListboxEvilSenders.SelectedItems.count -gt 0) {
                $ListboxEvilSenders.Items.RemoveAt($ListboxEvilSenders.SelectedIndex)
            }
            Resolve-RequiredInputs
        })
        $BtnAddSubject.Add_Click({
            Add-EvilSubjects
        })
        $BtnRemoveSubject.Add_Click({
            while($ListboxEvilSubjects.SelectedItems.count -gt 0) {
                $ListboxEvilSubjects.Items.RemoveAt($ListboxEvilSubjects.SelectedIndex)
            }
        })
        $InputEvilSenderTxt.Add_KeyDown({
            param ( 
                [Parameter(Mandatory)][Object]$sender,
                [Parameter(Mandatory)][Windows.Input.KeyEventArgs]$e
            )
            if ($e.Key -eq 'Enter') {
                Add-EvilSenders
            }
        })
        $InputSubjectLinesTxt.Add_KeyDown({
            param ( 
                [Parameter(Mandatory)][Object]$sender,
                [Parameter(Mandatory)][Windows.Input.KeyEventArgs]$e
            )
            if ($e.Key -eq 'Enter') {
                Add-EvilSubjects
            }
        })
        $InputStartDate.Add_SelectedDateChanged({
            Resolve-DateInputs
            Resolve-RequiredInputs
        })
        $InputEndDate.Add_SelectedDateChanged({
            Resolve-DateInputs
            Resolve-RequiredInputs
        })
        $HourDropdownStart.Add_SelectionChanged({
            Resolve-DateInputs
            Resolve-RequiredInputs
        })
        $HourDropdownEnd.Add_SelectionChanged({
            Resolve-DateInputs
            Resolve-RequiredInputs
        })
        $AmPmDropdownStart.Add_SelectionChanged({
            Resolve-DateInputs
            Resolve-RequiredInputs
        })
        $AmPmDropdownEnd.Add_SelectionChanged({
            Resolve-DateInputs
            Resolve-RequiredInputs
        })
        $Btn24h.Add_Click({
            $StartDateTime = (get-date (get-date -format "yyyy-MM-ddTHH:00:00")).AddHours(-24)
            $EndDateTime = (get-date (get-date -format "yyyy-MM-ddTHH:00:00")).AddHours(1)
            if ($StartDateTime.Hour -gt 12) {
                $HourDropdownStart.SelectedIndex = $StartDateTime.Hour - 13
                $AmPmDropdownStart.SelectedIndex = 1
            } else {
                $HourDropdownStart.SelectedIndex = $StartDateTime.Hour - 1
                if ($StartDateTime.Hour -ne 12) {
                    $AmPmDropdownStart.SelectedIndex = 0
                } else {
                    $AmPmDropdownStart.SelectedIndex = 1
                }
            }
            if ($EndDateTime.Hour -gt 12) {
                $HourDropdownEnd.SelectedIndex = $EndDateTime.Hour - 13
                $AmPmDropdownEnd.SelectedIndex = 1
            } else {
                $HourDropdownEnd.SelectedIndex = $EndDateTime.Hour - 1
                if ($EndDateTime.Hour -ne 12) {
                    $AmPmDropdownEnd.SelectedIndex = 0
                } else {
                    $AmPmDropdownEnd.SelectedIndex = 1
                }
            }
            $InputStartDate.SelectedDate = $StartDateTime.Date
            $InputEndDate.SelectedDate = $EndDateTime.Date
            Resolve-RequiredInputs
        })
        $Btn48h.Add_Click({
            $StartDateTime = (get-date (get-date -format "yyyy-MM-ddTHH:00:00")).AddHours(-48)
            $EndDateTime = (get-date (get-date -format "yyyy-MM-ddTHH:00:00")).AddHours(1)
            if ($StartDateTime.Hour -gt 12) {
                $HourDropdownStart.SelectedIndex = $StartDateTime.Hour - 13
                $AmPmDropdownStart.SelectedIndex = 1
            } else {
                $HourDropdownStart.SelectedIndex = $StartDateTime.Hour - 1
                if ($StartDateTime.Hour -ne 12) {
                    $AmPmDropdownStart.SelectedIndex = 0
                } else {
                    $AmPmDropdownStart.SelectedIndex = 1
                }
            }
            if ($EndDateTime.Hour -gt 12) {
                $HourDropdownEnd.SelectedIndex = $EndDateTime.Hour - 13
                $AmPmDropdownEnd.SelectedIndex = 1
            } else {
                $HourDropdownEnd.SelectedIndex = $EndDateTime.Hour - 1
                if ($EndDateTime.Hour -ne 12) {
                    $AmPmDropdownEnd.SelectedIndex = 0
                } else {
                    $AmPmDropdownEnd.SelectedIndex = 1
                }
            }
            $InputStartDate.SelectedDate = $StartDateTime.Date
            $InputEndDate.SelectedDate = $EndDateTime.Date
            Resolve-RequiredInputs
        })
        $Btn72h.Add_Click({
            $StartDateTime = (get-date (get-date -format "yyyy-MM-ddTHH:00:00")).AddHours(-72)
            $EndDateTime = (get-date (get-date -format "yyyy-MM-ddTHH:00:00")).AddHours(1)
            if ($StartDateTime.Hour -gt 12) {
                $HourDropdownStart.SelectedIndex = $StartDateTime.Hour - 13
                $AmPmDropdownStart.SelectedIndex = 1
            } else {
                $HourDropdownStart.SelectedIndex = $StartDateTime.Hour - 1
                if ($StartDateTime.Hour -ne 12) {
                    $AmPmDropdownStart.SelectedIndex = 0
                } else {
                    $AmPmDropdownStart.SelectedIndex = 1
                }
            }
            if ($EndDateTime.Hour -gt 12) {
                $HourDropdownEnd.SelectedIndex = $EndDateTime.Hour - 13
                $AmPmDropdownEnd.SelectedIndex = 1
            } else {
                $HourDropdownEnd.SelectedIndex = $EndDateTime.Hour - 1
                if ($EndDateTime.Hour -ne 12) {
                    $AmPmDropdownEnd.SelectedIndex = 0
                } else {
                    $AmPmDropdownEnd.SelectedIndex = 1
                }
            }
            $InputStartDate.SelectedDate = $StartDateTime.Date
            $InputEndDate.SelectedDate = $EndDateTime.Date
            Resolve-RequiredInputs
        })
        $Window.Add_ContentRendered({
            1..12 | foreach-object { $HourDropdownStart.AddChild($_);$HourDropdownEnd.AddChild($_) }
            @("AM","PM") | foreach-object { $AmPmDropdownStart.AddChild($_);$AmPmDropdownEnd.AddChild($_) }
            $Window.Activate()
        })
        $BtnSubmit.Add_Click({
            $Window.DialogResult = $true
            $Window.Close()
        })
        $GUIGood = $Window.ShowDialog()
        $Senders = $(foreach($Item in $ListboxEvilSenders.Items) {
            if ($Item.Content.length -gt 0) {
                $Item.Content
            } else {
                $Item
            }
        }) -join ","
        $Users2Lockdown = if ($AccountLockdownScriptName.Length -gt 0 -and
            (Test-Path $PSScriptRoot\$AccountLockdownScriptName)) {
                ($ListboxEvilSenders.Items |
                    Where-Object { $_.Content.Length -gt 0 -and $_.IsChecked -eq $true } |
                    Select-Object -ExpandProperty Content) -join ","
        } else {
            $null
        }
        $SubjectLineStr = $null
        if ($ListboxEvilSubjects.Items.Count -gt 0) { 
            $SubjectLineStr = '(Subject:'
            foreach($subject in $ListboxEvilSubjects.Items.ValueSubject) {
                $SubjectLineStr += '"'+$subject.Replace('"','""')+'" OR Subject:'
            }    
            $SubjectLineStr = $SubjectLineStr.TrimEnd(" OR Subject:")
            $SubjectLineStr += ')'
        }
        $DatesTimes = Resolve-DateInputs
        $StartDateTime = $DatesTimes.StartDateTime
        $EndDateTime = $DatesTimes.EndDateTime
        $EmailStatus = if ($textRadioEmailStatus.text -eq "Both") {
            "Pending","Delivered","FilteredAsSpam"
        } elseif ($textRadioEmailStatus.text -eq "Delivered") {
            "Pending","Delivered"
        } else {
            "Pending","FilteredAsSpam"
        }
        return [PSCustomObject]@{
            GUIGood = $GUIGood
            StartDateTime = $StartDateTime
            EndDateTime = $EndDateTime
            Senders = $Senders
            SubjectLineStr = $SubjectLineStr
            EmailStatus = $EmailStatus
            Users2Lockdown = $Users2Lockdown
        }
    }
    $GUIData = Get-GUIData -NoMFA:$NoMFA
    if ($GUIData.GUIGood) {
        #SET SENDERS TO AN ARRAY JOINED BY COMMAS
        $Senders = $GUIData.Senders

        #Set list of users to lockdown where applicable
        $Users2Lockdown = $GUIData.Users2Lockdown

        #SET Email status string for the message trace for Delivered, FilteredAsSpam, or both
        $EmailStatusFilter = $GUIData.EmailStatus

        #Set the search start date and end date based on the selected information
        $SearchStartDate = $GUIData.StartDateTime.ToUniversalTime()
        $SearchEndDate = $GUIData.EndDateTime.ToUniversalTime()
        
        #Set the subject lines string for the search query
        $SubjectLineStr = $GUIData.SubjectLineStr
        #Build Search Query from the SearchStartDate, SearchEndDate, subject line(s), and evil sender(s)

        $SearchQuery = "(Received:`""+$SearchStartDate.toString()+".."+$SearchEndDate.toString()+"`") AND "
        $SenderList = New-Object System.Collections.ArrayList
        #If more than one sender, create filter string with a bunch of ORs for senders; otherwise, just set one
        if ($Senders.IndexOf(",") -gt 0) {
            $SearchQuery += "("
            foreach($EvilSender in $Senders.Split(",")) {
                $addy = $EvilSender.Trim()
                $SearchQuery += "From:$addy OR "
                $SenderList.add($addy) | Out-Null
            }
            $SearchQuery = $SearchQuery.TrimEnd(" OR ")
            $SearchQuery += ")"
        } else {
            $SearchQuery += "From:$Senders"
            $SenderList.add($Senders) | Out-Null
        }
        #If a subject line is specified, add it to the search query
        if ($null -ne $SubjectLineStr) {
            $SearchQuery += " AND "+$SubjectLineStr
        }
    } else {
        Read-Host "GUI returned an error, press any key to stop"
        Disconnect-ExchangeOnline -confirm:$false
        Exit
    }
} else {
    $SendersHash = @{}
    do {
        $Prompt = "[Required]Email address of evil sender. This prompt will repeat until you press enter with " +
            "no information. Do not enter quotes or empty, extra spaces"
        $BadSender = Read-Host $Prompt
        do {
            #Validate that each email address at least matches the typical email pattern
            $SendersGood = 1
            if ($BadSender.length -gt 0) {
                if (!($BadSender -match ".*@.*\..*")) {
                    $SendersGood = 0
                    $BadSender = Read-Host "Bad email address input, try again"
                } else {
                    $SendersHash.add($BadSender,$null)
                }
            }
        } until ($SendersGood -eq 1)
    } until ($BadSender.length -eq 0)
    $Senders = $SendersHash.keys -join ","

    #Create list of users to lock down if script is supplied and found
    if ($AccountLockdownScriptName.Length -gt 0 -and
        (Test-Path $PSScriptRoot\$AccountLockdownScriptName) -and
        $Senders.IndexOf("@$EmailDomain") -gt 0) {
            $LockdownAddresses = (($SendersHash.Keys |
                Where-Object { $_ -match "^.+@$($EmailDomain.replace(".","\."))"}) -join ",")
            $Prompt = "Would you also like to run the lockdown script on these users in the " +
                $EmailDomain + ' domain?' + "`n$LockdownAddresses"
            if ((Read-Host $Prompt) -eq "y") {
                $Users2Lockdown = $LockdownAddresses
            }
    }

    #Get the start date and time of the range when that users should have received the spam
    $Prompt = "[Required]Start Date AND TIME for range when users received message (e.g. 7/18/2018 12:20 AM)"
    $StartDate = Read-Host $Prompt
    do {
        $good = 0
        try {
            $StartDate = Get-Date($StartDate)
            $good = 1
        }
        catch {
            $StartDate = Read-Host "Start date invalid, try again"
        }
    } until ($good -eq 1)
    $SearchStartDate = (get-date $StartDate).ToUniversalTime()

    #Get the end date and time of the range that users should have received the spam
    $Prompt = "[Required]End Date AND TIME for range when users received message (e.g. 7/18/2018 3:59 PM)"
    $EndDate = Read-Host $Prompt
    do {
        $good = 0
        try {
            $EndDate = Get-Date($EndDate)
            if ($EndDate -gt $StartDate) { $good = 1 } else { throw }
        }
        catch {
            $EndDate = Read-Host "End date invalid or before start date, try again"
        }
    } until ($good -eq 1)
    $SearchEndDate = (get-date $EndDate).ToUniversalTime()

    #Get the subject line(s) of the evil messages to filter
    $SubjectLineFilter = @{}
    $Prompt = "[Optional]Subject line to filter. This prompt will repeat until you press enter with no " +
        "information. Do not enter any quotes or backticks unless actually in the subject"
    do {
        $SubjectLine = Read-Host $Prompt
        if ($SubjectLine.length -gt 0) {
            $SubjectLineFilter.add($SubjectLine,$null)
        }
    } until ($SubjectLine.length -eq 0)
    $SubjectLineStr = $null
    if ($SubjectLineFilter.count -gt 0) {
        $SubjectLineStr = '(Subject:'
        foreach($subject in $SubjectLineFilter.keys) {
            $SubjectLineStr += '"'+$subject.Replace('"','""')+'" OR Subject:'
        }
        $SubjectLineStr = $SubjectLineStr.TrimEnd(" OR Subject:")
        $SubjectLineStr += ")"
    }

    #Get the desired Email status - Delivered, FilteredAsSpam, or both
    $EmailStatusFilter = @("Pending")
    $Prompt = "[Required]Status of emails being processed: `n[1]Delivered, [2]FilteredAsSpam, [3]Both`n" +
        "(Default is 3):"
    do {
        try {
            $StatusInput = [int](Read-Host $Prompt -ErrorAction Stop)
            switch ($StatusInput) {
                1 { $EmailStatusFilter = "Pending","Delivered" }
                2 { $EmailStatusFilter = "Pending","FilteredAsSpam" }
                Default { $EmailStatusFilter = "Pending","FilteredAsSpam","Delivered" }
            }
            $Good = 1
        } catch {Write-Host "Invalid Response"}
    } until ($Good -eq 1)

    #Build Search Query from the SearchStartDate, SearchEndDate, subject line(s), and evil sender(s)

    $SearchQuery = "(Received:`""+$SearchStartDate.toString()+".."+$SearchEndDate.toString()+"`") AND "
    $SenderList = New-Object System.Collections.ArrayList
    #If more than one sender, create filter string with a bunch of ORs for senders; otherwise, just set one
    if ($Senders.IndexOf(",") -gt 0) {
        $SearchQuery += "("
        foreach($EvilSender in $Senders.Split(",")) {
            $addy = $EvilSender.Trim()
            $SearchQuery += "From:$addy OR "
            $SenderList.add($addy) | Out-Null
        }
        $SearchQuery = $SearchQuery.TrimEnd(" OR ")
        $SearchQuery += ")"
    } else {
        $SearchQuery += "From:$Senders"
        $SenderList.add($Senders) | Out-Null
    }
    #If a subject line is specified, add it to the search query
    if ($null -ne $SubjectLineStr) {
        $SearchQuery += " AND "+$SubjectLineStr
    }
}

if ($Users2Lockdown.length -gt 0) {
    $Msg = "Starting script located at $PSScriptRoot\$AccountLockdownScriptName with the users $Users2Lockdown " +
        "and waiting until it is complete to continue."
    Write-Host $Msg
    $ArgString = "-file `"$("$PSScriptRoot\$AccountLockdownScriptName")`" -Users `"$Users2Lockdown`""
    Start-Process powershell -Passthru -Wait -ArgumentList $ArgString
}
$MyRecipients = @{}
$Page = 0
$SearchStartDateStr = $SearchStartDate.ToString()
$SearchEndDateStr = $SearchEndDate.ToString()
Write-Host "Getting recipients of evil message... from $SearchStartDateStr (UTC) to $SearchEndDateStr (UTC) from $SenderList"
#Since Message Traces cut off after 5k results, we use PageSize to limit it to 5k users and try another page of results till we run out
$MessageTraceArgs = @{
    SenderAddress = $SenderList
    StartDate = $SearchStartDateStr
    EndDate = $SearchEndDateStr
    Pagesize = 5000
    Status = $EmailStatusFilter
    ErrorAction = "Stop"
}
do {
    $Page++
    Write-Host "  Getting Page $Page of results, can take up to 5 minutes..."
    $a = (Get-MessageTrace @MessageTraceArgs -Page $Page | select-object recipientaddress).recipientaddress
    Write-Host "  Done."
    if ($null -ne $a) {
        <#For every person found in the trace, we look to make sure it is not already in the list and that it is
            an address in the supplied domain to which we can actually do something#>
        foreach($Recipient in $a) {
            if (!($MyRecipients.ContainsKey($Recipient)) -and
                $Recipient.IndexOf("@$EmailDomain") -gt 0 -and
                $Mailboxes2Exclude -notcontains $Recipient) {
                    $MyRecipients.Add($Recipient,$null)
            }
        }
    }
    #Just because these searches can be resource-intensive and occasionally freak, wait a second
    Start-Sleep -Seconds 1
} until ($null -eq $a)
Write-Host Done

try {
    if ($NoMFA) {
        Connect-IPPSSession -Credential $Cred -ErrorAction Stop
    } else {
        Connect-IPPSSession -UserPrincipalName $CredUPN -ErrorAction Stop
    }
} catch {
    $ErrMsg = "Error connecting to O365 Compliance and Security Center to create content search - $_. " +
        "Press any key to exit."
    Read-Host $ErrMsg
    Disconnect-ExchangeOnline -confirm:$false
    Exit
}

$ContentSearchCounter=0
$SearchNames=New-Object System.Collections.ArrayList
$CompletedSearches=0
$TotalRecipients2Process = $MyRecipients.Keys.Count
for($i=0;$i -lt $TotalRecipients2Process;$i+=50000) {
    $ContentSearchCounter++
    $SearchName = "SPAM-SearchAndHardPurge-$CredUPN-$(Get-Date -Format FileDateTime)"
    if ($MyRecipients.Keys.Count -gt 49999) {
        $SearchName+="-Part$ContentSearchCounter"
    }
    $SearchNames.Add($SearchName) | Out-Null
    $ComplianceArgs = @{
        Name = $SearchName
        Description = "Incident response, purging spam"
        ExchangeLocation = @($MyRecipients.Keys)[($i)..($i+49999)]
        AllowNotFoundExchangeLocationsEnabled = $true
        ContentMatchQuery = $SearchQuery
    }
    New-ComplianceSearch @ComplianceArgs | Start-ComplianceSearch
    try {
        do {
            Start-Sleep -Seconds 5
            Write-Host "  Checking to see if the search is completed and pausing for a few seconds if not."
        } until((Get-ComplianceSearch -Identity $SearchName).Status -eq "Completed")
    } catch {
        Read-Host "Error with the compliance search status - $_"
    }
    New-ComplianceSearchAction -SearchName $SearchName -Purge -PurgeType HardDelete -Confirm:$false
}
$PauseInMinutes = 1
do {
    Write-Host "$(Get-Date -Format u) : Pausing for $PauseInMinutes minutes to allow job to progress."
    Start-Sleep -Seconds ($PauseInMinutes*60)
    foreach($SearchName in $SearchNames) {
        $ActionInfo = Get-ComplianceSearchAction -Identity ($SearchName+'_Purge') -Details
        $Progress = ($ActionInfo.Results -split ";" |
            Where-Object { $_ -like " Item count:*" -and $_ -notlike " Item count: 0"}).count
        if ($Progress -lt $TotalRecipients2Process -and $ActionInfo.Status -ne "Completed") {
            Write-Host ("$(Get-Date -Format u) : Current Progress for $SearchName - " +
                "$Progress/$TotalRecipients2Process ($([int]($Progress/$TotalRecipients2Process*100))%) Complete")
        } else {
            Write-Host "$(Get-Date -Format u) : Current Progress for $SearchName - 100% Complete"
            $CompletedSearches++
        }
    }
} until($CompletedSearches -eq $ContentSearchCounter)
#When done, we want to use Remove-PSSession to make sure that we properly close our powershell o365 sessions
Disconnect-ExchangeOnline -confirm:$false
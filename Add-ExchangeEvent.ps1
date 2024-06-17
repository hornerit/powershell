#Requires -Modules Az.Accounts,Az.KeyVault,MSAL.PS
<#
.SYNOPSIS
Insert an appointment onto the calendar in a set of O365 mailboxes defined in a CSV file

.DESCRIPTION
Using an Azure App Registration (that has Calendar.ReadWrite permissions) Client Id and Secret and
the appropriate tenant information (Azure AD TenantId and Azure Subscription Id), this script will connect to
Azure, prompt for all typical information for a calendar event and a CSV file that has an EmailAddress column with
some recipients, connects to Microsoft Graph API, and attempts to add the calendar event silently to each mailbox.

.PARAMETER TenantId
REQUIRED The Azure Active Directory tenant ID (GUID) - you can find this in portal.azure.com -> Azure Active Directory

.PARAMETER ClientId
REQUIRED The client ID or app ID for the app registration created for adding events to calendars and has been granted the
application permission of Calendar.ReadWrite. This should be a GUID

.PARAMETER SubscriptionId
REQUIRED The Azure subscription ID where the keyvault is located and the current user has access to it

.PARAMETER VaultName
REQUIRED The name of the Azure KeyVault in which the client secret for the ClientId is located.

.PARAMETER SecretName
REQUIRED The name of the Secret entry in the Azure KeyVault that holds the client secret generated for the clientID above.

.PARAMETER Subject
OPTIONAL The Title of the calendar event

.PARAMETER Body
OPTIONAL The body of the message in the calendar event. If using HTML, do not include the <body> or <html>
tags. It is best to try avoiding this particular parameter if you have complex HTML because single and double
quotes cause problems with either the GUI or with the final result unless you surround your entire text in
parentheses - e.g. -Body ('<p>this is <a href="https://www.microsoft.com">microsoft's</a> website link')

.PARAMETER Start
OPTIONAL Date and time when the event should start. If only a date is supplied, 12:00 AM is assumed. Supply in local time.

.PARAMETER End
OPTIONAL Date and time when the event should end. If only a date is supplied, 12:00 AM is assumed. Supply in local time.

.PARAMETER ReminderMinutesBeforeStart
OPTIONAL The number of minutes before the start of event at which time an Outlook Reminder is triggered.

.PARAMETER ShowAs
OPTIONAL Valid values are "Free","Busy","Tentative","OOF", and "WorkingElsewhere". OOF means Out of Office.

.PARAMETER Location
OPTIONAL Where exactly is this event taking place?

.PARAMETER CsvFile
OPTIONAL Path to a CSV that contains a column for EmailAddress and there's a non-empty address in the first 4 entries.

.PARAMETER CSVPreCheck
OPTIONAL If set to $true, will pre-validate the UPN entries against local ActiveDirectory and separate invalid entries

.PARAMETER LocalADDomain
OPTIONAL If you plan on pre-checking local AD to verify UPNs, supply the domain.

.PARAMETER Credential
OPTIONAL If you plan on pre-checking local AD and wish to use a specific credential, supply it here.
.NOTES
    Author: Brendan Horner (MIT)
    Version History:
    --2024-04-16-Added error handling for users/mailboxes not found so it doesn't kill the script
    --2022-10-04-Added params for AD domain to make portability easier
    --2022-05-02-Added feature for pre-checking CSV for invalid UPNs from local AD
    --2021-10-04-Fixed bug for the appointment body itself not appearing
    --2021-09-30-Changed timezone from UTC to America/New_York
    --2021-09-01-Shifted logic to make Body an optional element for testing
    --2021-08-25-Updated output and added more data to the log file. Fixed horizontal scroll bar for Body field
    --2021-03-16-Re-arranged functions, added function for getting Azure token
    --2020-12-14-Initial version with a GUI and uses Graph API for the process to support modern auth.

.EXAMPLE
.\Add-ExchangeAppointment.ps1
.\Add-ExchangeAppointment.ps1 -Subject "Test Subject" -Body "Test Body" -Location "Your Desk"
.\Add-ExchangeAppointment.ps1 -Subject "Test Subject" -Body "This is an important body" -Start
    "2020-08-20 4:00 PM" -End "2020-08-20 4:30 PM" -Location "Your Desk" -ReminderMinutesBeforeStart 1440 -ShowAs
    "Free" -CsvFile = "C:\temp\mylist.csv"
#>

[CmdletBinding()]
param(
    [string]$Subject,
    [string]$Body, 
    [datetime]$Start, 
    [datetime]$End,
    [int]$ReminderMinutesBeforeStart,
    [ValidateSet("Free","Busy","Tentative","OOF","WorkingElsewhere")][string]$ShowAs,
    [string]$Location,
    [string]$CsvFile,
    [Parameter(Mandatory=$true)][string]$TenantId,
    [Parameter(Mandatory=$true)][string]$ClientId,
    [Parameter(Mandatory=$true)][string]$SubscriptionId,
    [Parameter(Mandatory=$true)][string]$VaultName,
    [Parameter(Mandatory=$true)][string]$SecretName,
    [bool]$CSVPreCheck,
    [string]$LocalADDomain,
    [pscredential]$Credential
)
###BEGIN BORING FUNCTION DECLARATIONS###
function Get-AzureToken {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true)][string]$TenantId,
        [parameter(Mandatory=$true)][string]$ClientId,
        [parameter(Mandatory=$true)][securestring]$ClientSecret
    )
    try {
        $GetTokenArgs = @{
            ClientId = $ClientId
            ClientSecret = $ClientSecret
            TenantId = $TenantId
            Scopes = "https://graph.microsoft.com/.default"
        }
        Write-Verbose -Message "$(Get-Date -format "yyyy-MM-ddTHH:mm:ss") - Obtaining a token from Microsoft Graph"
        return (Get-MsalToken @GetTokenArgs -ErrorAction Stop).AccessToken
    } catch {
        $Msg = "$(Get-Date -format "yyyy-MM-ddTHH:mm:ss") - Error obtaining token from Azure - $_."
        throw $Msg
    }
}
function Resolve-DateInputs {
    $InvalidStart=0
    $InvalidEnd=0
    $Valid = $false
    $StartDateTime = if ($null -ne $ApptStartDate.SelectedDate) {
        Get-Date $ApptStartDate.SelectedDate
    } else { $null }
    $EndDateTime = if ($null -ne $ApptEndDate.SelectedDate) {
        Get-Date $ApptEndDate.SelectedDate
    } else { $null }
    if ($HourDropdownStart.SelectedIndex -eq -1 -and
    $AmPmDropdownStart.SelectedIndex -eq -1 -and
    $null -ne $ApptStartDate.SelectedDate) {
        $HourDropdownStart.SelectedIndex = 11
        $AmPmDropdownStart.SelectedIndex = 0
    }
    if ($HourDropdownEnd.SelectedIndex -eq -1 -and
    $AmPmDropdownEnd.SelectedIndex -eq -1 -and
    $null -ne $ApptEndDate.SelectedDate) {
        $HourDropdownEnd.SelectedIndex = 11
        $AmPmDropdownEnd.SelectedIndex = 0
    }
    if ($null -ne $StartDateTime -and
    $HourDropdownStart.SelectedIndex -ne -1 -and
    !($HourDropdownStart.SelectedIndex -eq 11 -and $AmPmDropdownStart -eq 0) -and
    $AmPmDropdownStart.SelectedIndex -ne -1) {
        $Hours2Add = if ($AmPmDropdownStart.SelectedValue -eq "PM" -and
        $HourDropdownStart.SelectedIndex -ne 11) {
            [int]$HourDropdownStart.SelectedValue + 12
        } elseif ($AmPmDropdownStart.SelectedIndex -eq 0 -and $HourDropdownStart.SelectedValue -eq 12) {
            0
        } else {
            [int]$HourDropdownStart.SelectedValue
        }
        $StartDateTime = (Get-Date $ApptStartDate.SelectedDate).AddHours($Hours2Add)
    }
    if ($null -ne $EndDateTime -and
    $HourDropdownEnd.SelectedIndex -ne -1 -and
    !($HourDropdownEnd.SelectedIndex -eq 11 -and $AmPmDropdownEnd -eq 0) -and
    $AmPmDropdownEnd.SelectedIndex -ne -1) {
        $Hours2Add = if ($AmPmDropdownEnd.SelectedValue -eq "PM" -and
        $HourDropdownEnd.SelectedIndex -ne 11) {
            [int]$HourDropdownEnd.SelectedValue + 12
        } elseif ($AmPmDropdownEnd.SelectedIndex -eq 0 -and $HourDropdownEnd.SelectedValue -eq 12) {
            0
        } else {
            [int]$HourDropdownEnd.SelectedValue
        }
        $EndDateTime = (Get-Date $ApptEndDate.SelectedDate).AddHours($Hours2Add)
    }
    if ($null -ne $StartDateTime -and $null -ne $EndDateTime -and $StartDateTime -gt $EndDateTime) {
        $InvalidStart=1
        $InvalidEnd=1
    }
    if ($InvalidStart -eq 1) {
        $LabelStartDateWarnings.Content="INVALID"
    } else {
        $LabelStartDateWarnings.Content=""
    }
    if ($InvalidEnd -eq 1) {
        $LabelEndDateWarnings.Content="INVALID"
    } else {
        $LabelEndDateWarnings.Content=""
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
    if ($TextCSVPath.Text.Length -gt 0 -and
    (Test-Path -Path $TextCSVPath.Text)) {
        $Entries = (ConvertFrom-Csv (get-content ($TextCSVPath.Text) -TotalCount 5)).EmailAddress |
            Where-Object { $null -ne $_ }
        if($Entries.count -eq 0){
            $LabelTextCSVPath.Content="BAD CSV"
        }
    }
    if ($Entries.count -gt 0 -and
    (Resolve-DateInputs).Valid -and
    $ShowApptAsDropdown.SelectedIndex -ne -1 -and
    $ReminderNumberOfMinutes.Text -match "\d+" -and
    $ApptSubject.Text.Length -gt 0 -and
    $ApptLocation.Text.Length -gt 0) {
        $BtnSubmit.Content="Begin Calendar Injection"
        $BtnSubmit.IsEnabled = $true
    } else {
        $BtnSubmit.Content="[Disabled]"
        $BtnSubmit.IsEnabled = $false
    }
}
function Get-GUIData {
    [CmdletBinding()]
    param(
        [string]$subject,
        [string]$body, 
        [datetime]$start, 
        [datetime]$end,
        [int]$reminderMinutesBeforeStart,
        [string]$showAs,
        [string]$location,
        [string]$csvFile,
        [bool]$csvPreCheck
    )
    $TodayPlus3 = (Get-Date).AddDays(3).ToShortDateString()
    if ($start -lt $TodayPlus3 -and $null -ne $start) { Remove-Variable -Name "start" }
    if ($null -ne $end -and ($end -lt $start -or $end -lt $TodayPlus3)) { Remove-Variable -Name "end" }
    $body = [System.Security.SecurityElement]::Escape($body)
    [xml]$Xaml=@"
        <Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="Window" Title="Calendar Adder" WindowStartupLocation="CenterScreen"
        SizeToContent="WidthAndHeight" ShowInTaskbar="True"
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
                    <Label>Path to Recipient CSV:</Label>
                    <TextBox x:Name="TextCSVPath" Text="$($CsvFile)" Width="250" AcceptsReturn="False" Height="25" FontSize="14"/>
                    <Button x:Name="BtnBrowse" Content="Browse..." Margin="5,0,10,0" FontSize="18" IsEnabled="true" Height="25"/>
                    <Label x:Name="LabelTextCSVPath" FontSize="14" Foreground="red" VerticalAlignment="Top" Height="25" Margin="0"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <Label>Appointment Start:</Label>
                    <DatePicker x:Name="ApptStartDate" Margin="5,0,0,0" VerticalAlignment="Center" $(if($null -ne $start){'SelectedDate="'+$start.ToShortDateString()+'"'})>
                        <DatePicker.BlackoutDates>
                            <CalendarDateRange Start="01/01/1200" End="$($TodayPlus3)" />
                        </DatePicker.BlackoutDates>
                    </DatePicker>
                    <Label VerticalAlignment="Center">Hour</Label>
                    <ComboBox x:Name="HourDropdownStart" Height="25" Width="50" VerticalAlignment="Center"/>
                    <Label VerticalAlignment="Center">AM/PM</Label>
                    <ComboBox x:Name="AmPmDropdownStart" Height="25" Width="50" VerticalAlignment="Center"/>
                    <Label x:Name="LabelStartDateWarnings" FontSize="25" Foreground="red" VerticalAlignment="Center"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <Label>Appointment End:</Label>
                    <DatePicker x:Name="ApptEndDate" Margin="5,0,0,0" VerticalAlignment="Center" $(if($null -ne $end){'SelectedDate="'+$end.ToShortDateString()+'"'})>
                        <DatePicker.BlackoutDates>
                            <CalendarDateRange Start="01/01/1200" End="$($TodayPlus3)" />
                        </DatePicker.BlackoutDates>
                    </DatePicker>
                    <Label VerticalAlignment="Center">Hour</Label>
                    <ComboBox x:Name="HourDropdownEnd" Height="25" Width="50" VerticalAlignment="Center"/>
                    <Label VerticalAlignment="Center">AM/PM</Label>
                    <ComboBox x:Name="AmPmDropdownEnd" Height="25" Width="50" VerticalAlignment="Center"/>
                    <Label x:Name="LabelEndDateWarnings" FontSize="25" Foreground="red" VerticalAlignment="Center"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <Label>Show Appointment As:</Label>
                    <ComboBox x:Name="ShowApptAsDropdown" Height="25" Width="150" VerticalAlignment="Center"/>
                    <StackPanel x:Name="stackPanelCSVPreCheck" Orientation="Horizontal">
                        <Label>Pre-Check CSV Addresses?</Label>
                        <RadioButton GroupName="CSVPreCheck" VerticalAlignment="Center" Content="Yes" x:Name="Yes"/>
                        <RadioButton GroupName="CSVPreCheck" Margin="10,0,0,0" VerticalAlignment="Center" Content="No" x:Name="No" IsChecked="True"/>
                        <TextBox x:Name="textCSVPreCheck" Visibility="Hidden" Height="1" Width="1" Text="No"/>
                    </StackPanel>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <Label>How many MINUTES before Appt Start should reminder appear?</Label>
                    <TextBox x:Name="ReminderNumberOfMinutes" Width="30" $(if($reminderMinutesBeforeStart -gt 0){ 'Text="'+$reminderMinutesBeforeStart+'"' })/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <Label>Appt Subject:</Label>
                    <TextBox x:Name="ApptSubject" Text="$($Subject)" MinWidth="300" Width="Auto"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <Label>Appt Location:</Label>
                    <TextBox x:Name="ApptLocation" Text="$($Location)" MinWidth="300" Width="Auto"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <Label>Appt Body (if HTML, you do not need the Body tags, just everything inside):</Label>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                <TextBox x:Name="ApptBody" Text="$($Body)" Width="400" MaxWidth="400" MinHeight="50" Height="Auto" VerticalAlignment="stretch" AcceptsReturn="True" TextWrapping="NoWrap" HorizontalScrollBarVisibility="Auto"/>
                </StackPanel>
                <Button x:Name="BtnSubmit" Content="[Disabled]" Margin="20" FontSize="30" IsEnabled="false"/>
            </StackPanel>
        </Window>
"@ -replace 'x:N','N'
    $Reader = (New-Object System.Xml.XmlNodeReader $Xaml)
    $Window = [Windows.Markup.XamlReader]::Load($Reader)

    <#Powershell variables for the controls in the GUI, this first part loads all objects from the above xaml
        and creates powershell variables for each one with a name.#>
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
        #"trying item $($_.Name)"
        try {
            Set-Variable -Name $_.Name -Value $window.FindName($_.Name) -ErrorAction Stop
        } catch {
            throw
        }
    }
    #Event handler for the radio buttons, gets added to StackPanel that holds the radio buttons as an event
    [System.Windows.RoutedEventHandler]$Script:CheckedEventHandler = {
        $textCSVPreCheck.Text = $_.source.name
    }
    $stackPanelCSVPreCheck.AddHandler(
        [System.Windows.Controls.RadioButton]::CheckedEvent,
        $CheckedEventHandler
    )
    #Here we add triggers to the dropdowns and datepickers for the date fields to check for valid inputs
    $ApptStartDate.Add_SelectedDateChanged({
        Resolve-DateInputs
        Resolve-RequiredInputs
    })
    $ApptEndDate.Add_SelectedDateChanged({
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
    $ShowApptAsDropdown.Add_SelectionChanged({
        Resolve-DateInputs
        Resolve-RequiredInputs
    })
    #Add a trigger to validate the CSV path being supplied and that it has an EmailAddress column header
    $TextCSVPath.Add_TextChanged({
        if ((Test-Path $TextCSVPath.Text) -and $TextCSVPath.Text.Length -gt 0) {
            $Entries = (ConvertFrom-Csv (get-content -Path ($TextCSVPath.Text) -TotalCount 5)).EmailAddress |
                Where-Object { $null -ne $_ }
            if($Entries.Count -gt 0){
                $LabelTextCSVPath.Content=""
                Resolve-RequiredInputs
            } else {
                $LabelTextCSVPath.Content="BAD CSV"
            }
        } elseif ($TextCSVPath.Text.Length -gt 0) {
            $LabelTextCSVPath.Content="INVALID"
        }
    })
    #Validate all required inputs any time these fields are changed
    $ApptSubject.Add_TextChanged({
        Resolve-RequiredInputs
    })
    $ApptLocation.Add_TextChanged({
        Resolve-RequiredInputs
    })
    $ApptBody.Add_TextChanged({
        Resolve-RequiredInputs
    })
    $ReminderNumberOfMinutes.Add_TextChanged({
        Resolve-RequiredInputs
    })
    #This adds a proper Browse button for the CSV path selection that works in Windows and captures the path
    $BtnBrowse.Add_Click({
        $fileDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
        $fileResult = $fileDialog.ShowDialog()
        switch($fileResult){
            "OK" {
                $TextCSVPath.Text = $fileDialog.FileName
            }
            "Cancel" {

            }
            default {
                $TextCSVPath.Text = $null
            }
        }
        Resolve-RequiredInputs
    })
    $BtnSubmit.Add_Click({
        $Window.DialogResult = $true
        $Window.Close()
    })
    #Once the main window is rendered on the screen, we need to fill the dropdowns and pre-set them if supplied
    $Window.Add_ContentRendered({
        1..12 | foreach-object { $HourDropdownStart.AddChild($_);$HourDropdownEnd.AddChild($_) }
        @("AM","PM") | foreach-object { $AmPmDropdownStart.AddChild($_);$AmPmDropdownEnd.AddChild($_) }
        if ($null -ne $start) {
            if ($start.Hour -eq 0) {
                $HourDropdownStart.SelectedIndex = 11
                $AmPmDropdownStart.SelectedIndex = 0
            } elseif ($start.Hour -gt 12) {
                $HourDropdownStart.SelectedIndex = $start.Hour - 13
                $AmPmDropdownStart.SelectedIndex = 1
            } elseif ($start.Hour -eq 12) {
                $HourDropdownStart.SelectedIndex = 11
                $AmPmDropdownStart.SelectedIndex = 1
            } else {
                $HourDropdownStart.SelectedIndex = $start.Hour - 1
                $AmPmDropdownStart.SelectedIndex = 0
            }
        }
        if ($null -ne $end) {
            if ($end.Hour -eq 0) {
                $HourDropdownEnd.SelectedIndex = 11
                $AmPmDropdownStart.SelectedIndex = 0
            } elseif ($end.Hour -gt 12) {
                $HourDropdownEnd.SelectedIndex = $end.Hour - 13
                $AmPmDropdownEnd.SelectedIndex = 1
            } elseif ($end.Hour -eq 12) {
                $HourDropdownEnd.SelectedIndex = 11
                $AmPmDropdownEnd.SelectedIndex = 1
            } else {
                $HourDropdownEnd.SelectedIndex = $end.Hour - 1
                $AmPmDropdownEnd.SelectedIndex = 0
            }
        }
        @("Free","Busy","Tentative","OOF","WorkingElsewhere") | ForEach-Object {
            $ShowApptAsDropdown.AddChild($_)
        }
        if ($showAs -match "(Free|Busy|Tentative|OOF|WorkingElsewhere)") {
            $ShowApptAsDropdown.SelectedValue = $showAs
        }
        $Window.Activate()
    })
    Resolve-RequiredInputs
    $GUIGood = $Window.ShowDialog()
    $DatesTimes = Resolve-DateInputs
    $StartDateTime = $DatesTimes.StartDateTime
    $EndDateTime = $DatesTimes.EndDateTime
    return [PSCustomObject]@{
        GUIGood = $GUIGood
        CSVPath = $TextCSVPath.Text
        CSVPreCheck = if ($textCSVPreCheck.Text -eq "Yes") { $true } else { $false }
        ApptStartDateTime = $StartDateTime
        ApptEndDateTime = $EndDateTime
        ApptSubject = $ApptSubject.Text
        ApptLocation = $ApptLocation.Text
        ApptBody = $(if ($null -ne $ApptBody.Text -and $ApptBody.Text.Length -gt 1) { $ApptBody.Text } else { $null })
        ShowApptAs = $ShowApptAsDropdown.SelectedValue
        ApptReminder = [int]$ReminderNumberOfMinutes.Text
    }
}
###END OF BORING FUNCTION DECLARATIONS###
###MAIN###
$TokenArgs = @{
    ClientId = $ClientId
    TenantId = $TenantId
}
$LogPath = "$PSScriptRoot\Add-ExchangeEventSuccessLog.txt"
try {
    Write-Verbose -Message "$(Get-Date -format "yyyy-MM-ddTHH:mm:ss") - Connecting to Azure..."
    Connect-AzAccount -Subscription $SubscriptionId -Tenant $TenantId | Out-Null
    Write-Verbose -Message "$(Get-Date -format "yyyy-MM-ddTHH:mm:ss") - Done."
} catch {
    Read-Host "$(Get-Date -format "yyyy-MM-ddTHH:mm:ss") - Unable to connect to Azure - $_. press Enter to exit"
    exit
}
try {
    Write-Verbose -Message "$(Get-Date -format "yyyy-MM-ddTHH:mm:ss") - Obtaining app secret from KeyVault..."
    $TokenArgs."ClientSecret" = (Get-AzKeyVaultSecret -VaultName $VaultName -SecretName $SecretName).SecretValue
    if ($null -eq $TokenArgs.ClientSecret) { throw "Secret was null or empty"}
    Write-Verbose -Message "$(Get-Date -format "yyyy-MM-ddTHH:mm:ss") - Done."
} catch {
    Write-Host "$(Get-Date -format "yyyy-MM-ddTHH:mm:ss") - Unable to get app secret from Key Vault - $_."
    Read-Host "press Enter to exit script"
    Disconnect-AzAccount *> $null
    exit
}
$ScriptUsername = (Get-AzContext).Account.Id
Disconnect-AzAccount *> $null

#For GUI, load the assembly framework
try {
    Write-Verbose -Message "$(Get-Date -format "yyyy-MM-ddTHH:mm:ss") - Adding GUI assembly to PowerShell..."
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
    Write-Verbose -Message "$(Get-Date -format "yyyy-MM-ddTHH:mm:ss") - Done."
} catch {
    Read-Host "GUI failed to load, unable to continue. Press enter to exit."
    exit
}
$GUIArgs = @{
    subject = $subject
    body = $body
    start = $start
    end = $end
    reminderMinutesBeforeStart = $reminderMinutesBeforeStart
    showAs = $showAs
    location = $location
    csvFile = $csvFile
    csvPreCheck = $CSVPreCheck
}
if ($null -eq $Start) { $GUIArgs.Remove("start") }
if ($null -eq $End) { $GUIArgs.Remove("end") }
try {
    Write-Verbose -Message "$(Get-Date -format "yyyy-MM-ddTHH:mm:ss") - Prompting for GUI info..."
    $GUIData = Get-GUIData @GUIArgs
    if ($GUIData.GUIGood -eq $false) {
        throw
    }
    Write-Verbose -Message "$(Get-Date -format "yyyy-MM-ddTHH:mm:ss") - Done."
} catch {
    Read-Host "Error using the GUI or GUI was canceled - $_, press Enter to exit..."
    exit
}
$ReRunMessage = ("If you would like to run this script again, this is the command " +
    "to pre-fill all of your values:`n& '$PSScriptRoot\Add-ExchangeEvent.ps1' " +
    "-Subject '$($GUIData.ApptSubject)' " +
    "-Body '$($GUIData.ApptBody)' " +
    "-ReminderMinutesBeforeStart $($GUIData.ApptReminder) -ShowAs '$($GUIData.ShowApptAs)' " +
    "-Location '$($GUIData.ApptLocation)' -CsvFile '$($GUIData.CSVPath)'" +
    "$(if($GUIData.ApptStartDateTime){" -Start '$($GUIData.ApptStartDateTime)' " +
        "-End '$($GUIData.ApptEndDateTime)"})'" +
    "$(if($GUIData.CSVPreCheck -eq $true){ " -CSVPreCheck `$true" })")
    $postBodyHash = @{
    start = @{
        dateTime=$GUIData.ApptStartDateTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss")
        timeZone="UTC"
    }
    end = @{
        dateTime = $GUIData.ApptEndDateTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss")
        timeZone="UTC"
    }
    subject = $GUIData.ApptSubject
    location = @{
        displayname=$GUIData.ApptLocation
    }
    showAs = $GUIData.ShowApptAs
    reminderMinutesBeforeStart = $GUIData.ApptReminder
    isReminderOn = "true"
}
if ($null -ne $GUIData.ApptBody) {
    $postBodyHash."body" = @{
        contentType = "HTML"
        content = $GUIData.ApptBody
    }
}
#This next line is to take any special characters entered in the textbox for the body message as they get auto
    #converted to \\uXXXX as, I believe, an xml code. Using this converts the \\uXXXX to its actual character value
$postBody = $postBodyHash | ConvertTo-Json | foreach-object {
    [Regex]::Replace($_, "\\u(?<Value>[a-zA-Z0-9]{4})", {
        param($m) ([char]([int]::Parse($m.Groups['Value'].Value,
            [System.Globalization.NumberStyles]::HexNumber))).ToString() 
        }
    )
}
$PostArgs = @{
    Method = "POST"
    Body = $PostBody
    Headers = @{
        "Authorization" = "Bearer $(Get-AzureToken @TokenArgs)"
        "Content-Type" = "application/json"
        "Accept" = "application/json, text/plain"
    }
}
$Recipients = @((import-csv $GUIData.CSVPath).EmailAddress | Sort-Object)
if ($GUIData.csvPreCheck -eq $true -or $CSVPreCheck -eq $true) {
    $adDCArgs = @{
        Discover = $true
        NextClosestSite = $true
        ErrorAction = "STOP"
    }
    if ($LocalADDomain) {
        $adDCArgs.DomainName = $LocalADDomain
    }
    do {
        try {
            $ScriptDC = [string](Get-ADDomainController @adDCArgs).Hostname
            if ($Credential) {
                $adu = Get-ADUser $ScriptUsername -ErrorAction "Stop" -Server $ScriptDC -Credential $Credential
            } else {
                $adu = Get-ADUser $ScriptUsername -ErrorAction "Stop" -Server $ScriptDC
            }
        } catch {
            if ($_.exception.message -like "*rejected the client credentials*") {
                $Credential = Get-Credential $ScriptUsername -Message "Previous cred info was invalid, try again"
            }
        }
    } until ($null -ne $ScriptDC)
    $adArgs = @{
        ErrorAction = "Stop"
        Server = $ScriptDC
    }
    if ($Credential) {
        $adArgs.Credential = $Credential
    }
    $badRecipients = New-Object -TypeName System.Collections.ArrayList
    Write-Host "$(Get-Date -Format u) - Beginning CSV Pre-Check"
    foreach ($recipient in $Recipients) {
        try {
            $adu = @(Get-ADUser -Filter "userPrincipalName -eq '$recipient'" @adArgs)
            if ($adu.Count -ne 1) {
                throw
            }
        } catch {
            $badRecipients.Add($recipient) | Out-Null
        }
    }
    Write-Host "$(Get-Date -Format u) - Completed CSV Pre-Check"
    if ($badRecipients.Count -gt 0) {
        $badPath = "$PSScriptRoot\Add-ExchangeEventBadCSVEntries.txt"
        $now = Get-Date -Format u
        "$now - Invalid UPNs supplied:`n$($badRecipients -join "`n")" | Out-File -FilePath $badPath -Append
        Write-Host "$($badRecipients.Count) bad recipients removed during CSV pre-check and output to '$badPath'."
        $Recipients = $Recipients | Where-Object { $badRecipients -notcontains $_ }
    }
}
Write-Host "Beginning to post calendar events for $($Recipients.count) and will log progress in $LogPath"
("**************$(Get-Date -format u) - Beginning to post calendar events for $($Recipients.Count) mailboxes " +
    "using the following criteria:`nSubject - '$($GUIData.ApptSubject)'`nBody - '$($GUIData.ApptBody)'" +
    "`nReminderMinutesBeforeStart - $($GUIData.ApptReminder)`nShowAs - '$($GUIData.ShowApptAs)'" +
    "`nLocation - '$($GUIData.ApptLocation)'`nCsvFile - '$($GUIData.CSVPath)'" +
    "$(if($GUIData.ApptStartDateTime){"`nStart - '$($GUIData.ApptStartDateTime)'`nEnd - '$($GUIData.ApptEndDateTime)"})'" +
    "$(if($GUIData.CSVPreCheck -eq $true){ "`nCSVPreCheck - `$true" })") |
    Out-File -FilePath $LogPath -Append
$totalErrors = 0
$RecipientsNotFound = New-Object -TypeName System.Collections.ArrayList
foreach ($recipient in $Recipients) {
    Start-Sleep -Milliseconds 250
    $errorCounter = 0
    $successful = $false
    $PostArgs.Uri = "https://graph.microsoft.com/v1.0/users/$recipient/calendar/events"
    do {
        try {
            $result = Invoke-RestMethod @PostArgs -ErrorAction Stop
            $successful = $true
            "$(Get-Date -format u) - $recipient processed successfully." | Out-File -FilePath $LogPath -Append
        } catch {
            $errorCounter++
            if ($errorCounter -lt 2) {
                if ($_ -like "*expired*" -or $_ -like "*unauthorized*") {
                    try {
                        $PostArgs.Headers."Authorization" = "Bearer $(Get-AzureToken @TokenArgs)"
                    } catch {
                        $Msg = ("$(Get-Date -format u) - Error obtaining refresh token from Azure while processing " +
                            "$Recipient - $_.")
                        Write-Host $Msg
                        $Msg | Out-File -FilePath $LogPath -Append
                        Write-Host -Object $ReRunMessage
                        Read-Host "press Enter to exit."
                        exit
                    }
                } else {
                    Start-Sleep -Milliseconds 250
                }
            }
        }
    } until ($successful -or $errorCounter -eq 2)
    if ($errorCounter -eq 2) {
        if ($Error[0].Exception.Message -like "*not found*") {
            $RecipientsNotFound.Add($recipient) | Out-Null
        } else {
            $Message = "$(Get-Date -format u) - Error adding event to calendar for $recipient - " +
                "$($Error[0].exception.message)"
            Write-Host $Message
            $Message | Out-File -FilePath $LogPath -Append
            $totalErrors++
        }
    }
    if ($totalErrors -gt 100) {
        ("********************************$(Get-Date -format u) - ERROR OUTPUT, TOO MANY ERRORS, SCRIPT HALTED " +
            "*****************************") | Out-File -FilePath $LogPath -Append
        $Error | Out-File -FilePath $LogPath -Append
        exit
    }
}
if ($RecipientsNotFound.Count -gt 0) {
    "The following users were not found in Exchange Online: $($RecipientsNotFound -join "`n")" |
        Out-File -FilePath $LogPath -Append
    Write-Warning -Message ("$() users supplied were not found in Exchange Online. Pretty please, YELL AT WHOEVER " +
        "GAVE YOU THE LIST.")
}
Write-Host -Object $ReRunMessage
Write-Host -Object "Don't forget to remove those that were successful from your CSV if you must re-run"
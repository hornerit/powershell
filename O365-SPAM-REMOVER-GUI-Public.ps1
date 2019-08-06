#REQUIRES -Version 5
<#
.SYNOPSIS
Performs a message trace for a spam message and searches for and purges it from recipient mailboxes with multiple windows

.DESCRIPTION
Connects to O365 - assuming you have Exchange Admin permissions for each credential you supply - and performs a Get-MessageTrace. From there, it will take all of the recipients, split them up into rounds of child windows, and perform Search-Mailbox -DeleteContent commands against each recipient.

.PARAMETER StartDate
REQUIRED The beginning of the search for the spam message - typically the day BEFORE users receive the message, must be within the last 7 days
.PARAMETER EndDate
REQUIRED The end of the search for the spam message - typically the day AFTER users receive the message
.PARAMETER Recipients
OPTIONAL Will be determined by message trace but can be supplied separately - expecting a comma-separated string of email addresses, max of 1000 for performance's sake
.PARAMETER CredU
OPTIONAL Username of the Exchange Admin credentials
.PARAMETER CredP
OPTIONAL Encrypted password secure string of the Exchange Admin credentials (password encrypted via convertfrom-securestring command)
.PARAMETER SearchQuery
OPTIONAL Uses the date when users are expected to have received the spam message and the evil senders to form a query for Search-Mailbox command
.PARAMETER NoGUI
OPTIONAL This script detects the Windows Presentation Framework and attempts to make a GUI from it, use this switch to force command line interactive prompts instead
.PARAMETER EmailDomain
REQUIRED Domain of the mailboxes affected by spam campaign (used for message trace search/filter). Defaults to contoso.com and will prompt if you don't supply yours.
.PARAMETER WindowsPerCred
OPTIONAL This script defaults to 3 simultaneous powershell windows per Exchange Admin credential supplied due to typical tenant limits in O365, adjust at your own peril
.PARAMETER MailboxesPerWindow
OPTIONAL Number of Mailboxes to process per powershell window generated. Defaults to 500 and auto-adjusts for smaller jobs, adjust at your own peril. The total length of all email addresses plus the search terms must be less than 80k characters.

.NOTES
  Created by: Brendan Horner (www.hornerit.com)
  Notes: MUST BE RUN AS SCRIPT FILE, do NOT copy-paste into PS to run
  Version History:
  --2019-08-06-Bugfix for hour calculation using new buttons
  --2019-07-16-Added 2 new features - buttons for recent times and filters for email status
  --2019-07-15-Fixed bug for child windows again for MFA parameter
  --2019-06-27-Bugfix for MFA Module loading to use LastWriteTime instead of Modified since Modified doesn't exist
  --2019-06-19-Altered MFA parameter to be NoMFA so someone can force basic auth by setting that switch and adjusted MFA module to pull the latest version of the module on your machine. Bug fix from anonymous commenter.
  --2019-05-28-Bug fixes for MFA and mailbox errors
  --2019-05-21-Added Exchange Admin prompt to gui
  --2019-05-16-Added GUI, fixed a few minor display bugs and performance bugs, rewrote sections for dynamic window generation based on params, allows as many Exch Admin accts to assist as you can try...watch out for RAM usage
  --2019-05-02-Added better logic for throttling
  --2019-04-15-Initial public version

.EXAMPLE
.\O365-SPAM-REMOVER.ps1 -NoMFA
.\O365-SPAM-REMOVER.ps1 -Recipients "someone@CONTOSO.COM,someoneelse@CONTOSO.COM" -SearchQuery "FROM:bob@something.com AND Received:04/19/2018"
#>
[CmdletBinding()]
param(
[string]
$StartDate,
[string]
$EndDate,
[string]
$Recipients,
[string]
$CredU,
[string]
$CredP,
[string]
$SearchQuery,
[switch]
$NoGUI,
[string]
$EmailDomain = "contoso.com",
[int]
$WindowsPerCred = 3,
[int]
$MailboxesPerWindow = 500,
[switch]
$NoMFA
)

#Try to get the Exchange Online Powershell module that supports MFA
if(!($NoMFA)){
    try{
        $getChildItemSplat = @{
            Path = "$Env:LOCALAPPDATA\Apps\2.0\*\CreateExoPSSession.ps1"
            Recurse = $true
            ErrorAction = 'Stop'
            Verbose = $false
        }
        $MFAExchangeModule = ((Get-ChildItem @getChildItemSplat | Sort-Object LastWriteTime -Descending | where-object {(Test-Path "$($_.PSParentPath)\Microsoft.Exchange.Management.ExoPowershellModule.dll") -eq $true} | Select-Object -First 1 | Select-Object -ExpandProperty fullname).Replace("\CreateExoPSSession.ps1", ""))
        . "$MFAExchangeModule\CreateExoPSSession.ps1" 3>$null
        Write-Host "MFA Module found and imported"
    } catch {
        $NoMFA = $true
        Write-Host "MFA Module not found. If legacy auth is disabled for your tenant, this script will most likely fail. To install the latest module, go to https://aka.ms/exopspreview"
    }
}

#If supplied, create the Credential object used to log into O365 session. If we are using MFA, the credential token cache should hopefully still be working so just connecting without creds will work
if($null -ne $CredU -and $CredU.Length -gt 0){
    write-host "Username supplied, attempting to connect to O365"
    if(!($NoMFA)){
        Get-PSSession | Remove-PSSession
        try {
            Connect-EXOPSSession -UserPrincipalName $CredU 3>$null
        } catch {
            Write-Host "Unable to establish a connection to O365. You will need to re-run the script using this command to simply try again or you can change something:"
            Write-Host "$($MyInvocation.MyCommand.Definition) -CredU $CredU -CredP $CredP -SearchQuery `"$SearchQuery`" -Recipients $Recipients"
            do{
                $ReadyToClose = Read-Host "Press 'Y' key to close this script (make sure you copy the above command to paste and run in another powershell window to try again without running the whole giant script again)"
            } until ($ReadyToClose -eq "y")
            exit
        }
        $Session = Get-PSSession
    } else {
        $Cred = New-Object PSCredential($CredU,(ConvertTo-SecureString $CredP))
        try{
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Authentication Basic -AllowRedirection -Credential $Cred -ErrorAction Stop
        } catch {
            Write-Host "Unable to establish a connection to O365. You will need to re-run the script using this command to simply try again or you can change something:"
            Write-Host "$($MyInvocation.MyCommand.Definition) -CredU $CredU -CredP $CredP -SearchQuery `"$SearchQuery`" -Recipients $Recipients -NoMFA"
            do{
                $ReadyToClose = Read-Host "Press 'Y' key to close this script (make sure you copy the above command to paste and run in another powershell window to try again without running the whole giant script again)"
            } until ($ReadyToClose -eq "y")
            exit
        }
    }
    try {
        Invoke-Command -Session $Session -ScriptBlock { Get-OrganizationConfig | Select-Object Name } -ErrorAction Stop
    } catch {
        Read-Host "The account supplied does not appear to be an Exchange Admin account, please try again. Press any key to close..."
        exit
    }
} else {
    $Creds = New-Object -TypeName System.Collections.ArrayList
}

#Change window appearance if this is a child window so that it is smaller
if($Recipients.Length -gt 0){
    $title = "SPAM REMOVER - Processing "+$Recipients.Substring(0,$Recipients.IndexOf("@"))+" thru "+$Recipients.Substring($Recipients.LastIndexOf(",")+1,$Recipients.LastIndexOf("@")-$Recipients.LastIndexOf(",")-1)
    $Host.ui.RawUI.WindowTitle = $title
    $newSize = $Host.UI.RawUI.WindowSize
    $newSize.Height = 30
    $newSize.Width = 75
    $Host.UI.RawUI.WindowSize = $newSize
    $newBuffer = $Host.UI.RawUI.BufferSize
    $newBuffer.Height = 3000
    $newBuffer.Width = 75
    $Host.UI.RawUI.BufferSize = $newBuffer
}

#If SearchQuery has not been supplied, get the evil sender(s) then the received date and, finally, the subject line(s) of the evil emails
if($SearchQuery.Length -eq 0){
    #In case someone forgot to change the email domain, ask for it
    if($EmailDomain -eq "contoso.com"){
        do{
            $EmailDomain = Read-Host "You have not updated the script with your email domain for your recipients. Please enter the domain (the part after the @ symbol, e.g. mycompany.com)"
            if($EmailDomain -notmatch "^[a-zA-Z0-9]+\..*$"){
                Write-Host "Invalid entry"
                $EmailDomain = $null
            }
        } until ($null -ne $EmailDomain)
    }
    
    ###BELOW IF THE CODE NECESSARY FOR THE GUI...I KNOW IT LOOKS CRAZY...KINDA IGNORE IT MAYBE
    try { 
        Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
        $GUI = $true
    } catch {
        $GUI = $false
    }
    if($GUI -and !($NoGUI)){
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
            $Creds = New-Object -TypeName System.Collections.ArrayList
            [xml]$Xaml=@"
                <Window
                xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                x:Name="Window" Title="SPAM Removal Script GUI" WindowStartupLocation="CenterScreen" SizeToContent="WidthAndHeight" ShowInTaskbar="True" ScrollViewer.VerticalScrollBarVisibility="Auto">
                    <Window.Resources>
                        <Style x:Key="ButtonRoundedCorners" TargetType="Button">
                            <Setter Property="Template">
                                <Setter.Value>
                                    <ControlTemplate TargetType="Button">
                                        <Grid>
                                            <Border x:Name="border" CornerRadius="5" BorderBrush="#707070" BorderThickness="1" Background="LightGray">
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
                            <StackPanel Orientation="Vertical" VerticalAlignment="Center" Margin="3,0,10,0">
                                <Label Margin="0,0,0,-10">[Required]</Label>
                                <Label>Enter EXO Admin Accts, click "--&#62;"</Label>
                                <TextBox x:Name="inputExchAdmin" Height="25" Width="200"/>
                            </StackPanel>
                            <StackPanel Orientation="Vertical" VerticalAlignment="Bottom">
                                <Button x:Name="btnAddExchAdmin" Margin="0,0,10,0" Height="25" Width="40">
                                    <Button.Content>
                                        <TextBlock x:Name="txtBtnAddExchAdmin" RenderTransformOrigin="0.5,0.53" Margin="0" Padding="0">
                                            <TextBlock.RenderTransform>
                                                <RotateTransform x:Name="btnAddExchAdminRotate" Angle="0" />
                                            </TextBlock.RenderTransform>
                                            <TextBlock.Resources>
                                                <Storyboard x:Key="RotateAddExchAdminButton" x:Name="RotateAddExchAdminButton">
                                                    <DoubleAnimation Storyboard.TargetName="btnAddExchAdminRotate" Storyboard.TargetProperty="(RotateTransform.Angle)" From="0.0" To="360" Duration ="0:0:1" RepeatBehavior="Forever"/>
                                                </Storyboard>
                                            </TextBlock.Resources>
                                            <TextBlock.Text>--&#62;</TextBlock.Text>
                                        </TextBlock>
                                    </Button.Content>
                                </Button>
                                <Button x:Name="btnRemoveExchAdmin" Content="&#60;--" Margin="0,10,10,0" Height="25" Width="40"/>
                            </StackPanel>
                            <StackPanel Orientation="Vertical">
                                <Label>Exchange Online Admins</Label>
                                <ListBox x:Name="listboxExchangeAdmins" Height ="50" Width="250" SelectionMode="Extended" ScrollViewer.VerticalScrollBarVisibility="Auto"/>
                            </StackPanel>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal">
                            <StackPanel Orientation="Vertical" VerticalAlignment="Center" Margin="0,0,10,0">
                                <Label Margin="0,0,0,-10">[Required]</Label>
                                <Label>Enter spam email address, click "--&#62;"</Label>
                                <TextBox x:Name="inputEvilSenderTxt" Height="25" Width="200"/>
                            </StackPanel>
                            <StackPanel Orientation="Vertical" VerticalAlignment="Center">
                                <Button x:Name="btnAddEvilSender" Content="--&#62;" Margin="0,0,10,0" Height="25" Width="40"/>
                                <Button x:Name="btnRemoveEvilSender" Content="&#60;--" Margin="0,10,10,0" Height="25" Width="40"/>
                            </StackPanel>
                            <StackPanel Orientation="Vertical">
                                <Label>SPAM Senders</Label>
                                <ListBox x:Name="listboxEvilSenders" Height ="150" Width="250" SelectionMode="Extended" ScrollViewer.VerticalScrollBarVisibility="Auto"/>
                            </StackPanel>
                        </StackPanel>
                        <Separator Margin="0,10"/>
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Left" VerticalAlignment="Bottom" Margin="10,0,0,0">
                            <Button x:Name="btn24h" Content="Last 24h" Margin="0,0,20,0"/>
                            <Button x:Name="btn48h" Content="Last 48h" Margin="0,0,20,0"/>
                            <Button x:Name="btn72h" Content="Last 72h" Margin="0,0,20,0"/>
                            <StackPanel x:Name="stackPanelRadioEmailStatus" Orientation="Horizontal">
                                <RadioButton GroupName="radioEmailStatus" Content="Delivered" x:Name="Delivered" Margin="20,0,0,0"/>
                                <RadioButton GroupName="radioEmailStatus" Content="FilteredAsSpam" x:Name="FilteredAsSpam" Margin="20,0,0,0"/>
                                <RadioButton GroupName="radioEmailStatus" Content="Both" x:Name="Both" Margin="20,0,0,0" IsChecked="True"/>
                                <TextBox x:Name="textRadioEmailStatus" Visibility="Hidden" Height="1" Width="1" Margin="20,0,0,0" Text="Both"/>
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
                                <ComboBox x:Name="cbStartHour" Height="25" Width="50"/>
                            </StackPanel>
                            <StackPanel Orientation="Vertical" HorizontalAlignment="Center" VerticalAlignment="Bottom">
                                <Label HorizontalAlignment="Center">AM/PM</Label>
                                <ComboBox x:Name="cbStartAMPM" Height="25" Width="50"/>
                            </StackPanel>
                            <Label x:Name="lblStartDateWarnings" VerticalAlignment="Bottom" FontSize="25" Foreground="red" Margin="0,0,0,-8"/>
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
                                <ComboBox x:Name="cbEndHour" Height="25" Width="50"/>
                            </StackPanel>
                            <StackPanel Orientation="Vertical" VerticalAlignment="Bottom">
                                <Label HorizontalAlignment="Center">AM/PM</Label>
                                <ComboBox x:Name="cbEndAMPM" Height="25" Width="50"/>
                            </StackPanel>
                            <Label x:Name="lblEndDateWarnings" VerticalAlignment="Bottom" FontSize="25" Foreground="red" Margin="0,0,0,-8"/>
                        </StackPanel>
                        <Separator Margin="0,10"/>
                        <StackPanel Orientation="Horizontal">
                            <StackPanel Orientation="Vertical" VerticalAlignment="Center">
                                <Label Margin="0,0,0,-10">[Optional]</Label>
                                <Label>Enter subject lines, click "--&#62;"</Label>
                                <TextBox x:Name="inputSubjectLinesTxt" Width="200" Height="25" Margin="0,0,10,0"/>
                            </StackPanel>
                            <StackPanel Orientation="Vertical" VerticalAlignment="Center">
                                <Button x:Name="btnAddSubject" Content="--&#62;" Margin="0,0,10,0" Height="25" Width="40"/>
                                <Button x:Name="btnRemoveSubject" Content="&#60;--" Margin="0,10,10,0" Height="25" Width="40"/>
                            </StackPanel>
                            <StackPanel Orientation="Vertical">
                                <Label>SPAM Email Subject Lines</Label>
                                <ListBox x:Name="listboxEvilSubjects" Height = "150" Width="250" SelectionMode="Extended" ScrollViewer.VerticalScrollBarVisibility="Auto" DisplayMemberPath="DisplayName"/>
                            </StackPanel>
                        </StackPanel>
                        <Button x:Name="btnSubmit" Content="[Disabled]" Margin="20" FontSize="30" IsEnabled="false"/>
                    </StackPanel>
                </Window>
"@
            $Reader = (New-Object System.Xml.XmlNodeReader $Xaml)
            $Window = [Windows.Markup.XamlReader]::Load($Reader)

            #Powershell variables for the controls in the GUI
            $stackPanelRadioEmailStatus = $Window.FindName('stackPanelRadioEmailStatus')
            $InputEvilSenderTxt = $Window.FindName('inputEvilSenderTxt')
            $ListboxEvilSenders = $Window.FindName('listboxEvilSenders')
            $BtnAddEvilSender = $Window.FindName('btnAddEvilSender')
            $BtnRemoveEvilSender = $Window.FindName('btnRemoveEvilSender')
            $Btn24h = $Window.FindName('btn24h')
            $Btn48h = $Window.FindName('btn48h')
            $Btn72h = $Window.FindName('btn72h')
            $InputStartDate = $Window.FindName('InputStartDate')
            $HourDropdownStart = $Window.FindName('cbStartHour')
            $AmPmDropdownStart = $Window.FindName('cbStartAMPM')
            $LblStartDateWarnings = $Window.FindName('lblStartDateWarnings')
            $InputEndDate = $Window.FindName('InputEndDate')
            $HourDropdownEnd = $Window.FindName('cbEndHour')
            $AmPmDropdownEnd = $Window.FindName('cbEndAMPM')
            $LblEndDateWarnings = $Window.FindName('lblEndDateWarnings')
            $InputSubjectLinesTxt = $Window.FindName('inputSubjectLinesTxt')
            $ListboxEvilSubjects = $Window.FindName('listboxEvilSubjects')
            $BtnAddSubject = $Window.FindName('btnAddSubject')
            $BtnRemoveSubject = $Window.FindName('btnRemoveSubject')
            $BtnSubmit = $Window.FindName('btnSubmit')
            $InputExchAdmin = $Window.FindName('inputExchAdmin')
            $BtnAddExchAdmin = $Window.FindName('btnAddExchAdmin')
            $TxtBtnAddExchAdmin = $Window.FindName('txtBtnAddExchAdmin')
            $Animation4txtBtnAddExchAdmin = $Window.FindName('RotateAddExchAdminButton')
            $BtnRemoveExchAdmin = $Window.FindName('btnRemoveExchAdmin')
            $ListboxExchangeAdmins = $Window.FindName('listboxExchangeAdmins')
            $textRadioEmailStatus = $Window.FindName('textRadioEmailStatus')

            #Event handler for the radio buttons, gets added to the StackPanel that holds the radio buttons as an event
            [System.Windows.RoutedEventHandler]$Script:CheckedEventHandler = {
                $textRadioEmailStatus.Text = $_.source.name
            }
            $stackPanelRadioEmailStatus.AddHandler([System.Windows.Controls.RadioButton]::CheckedEvent, $CheckedEventHandler)

            #Functions attached to different controls in the GUI or run from using different controls
            function Resolve-DateInputs{
                $InvalidStart=0
                $InvalidEnd=0
                $Valid = $false
                $StartDateTime = if($null -ne $InputStartDate.SelectedDate){ Get-Date $InputStartDate.SelectedDate } else { $null }
                $EndDateTime = if($null -ne $InputEndDate.SelectedDate){ Get-Date $InputEndDate.SelectedDate } else { $null }
                if($HourDropdownStart.SelectedIndex -eq -1 -and $AmPmDropdownStart.SelectedIndex -eq -1 -and $null -ne $InputStartDate.SelectedDate){
                    $HourDropdownStart.SelectedIndex= 11
                    $AmPmDropdownStart.SelectedIndex= 0
                }
                if($HourDropdownEnd.SelectedIndex -eq -1 -and $AmPmDropdownEnd.SelectedIndex -eq -1 -and $null -ne $InputEndDate.SelectedDate){
                    $HourDropdownEnd.SelectedIndex= 11
                    $AmPmDropdownEnd.SelectedIndex= 0
                }
                if($null -ne $StartDateTime -and $HourDropdownStart.SelectedIndex -ne -1 -and !($HourDropdownStart.SelectedIndex -eq 11 -and $AmPmDropdownStart -eq 0) -and $AmPmDropdownStart.SelectedIndex -ne -1){
                    $Hours2Add = if($AmPmDropdownStart.SelectedValue -eq "PM" -and $HourDropdownStart.SelectedIndex -ne 11){ [int]$HourDropdownStart.SelectedValue + 12 } elseif($AmPmDropdownStart.SelectedIndex -eq 0 -and $HourDropdownStart.SelectedValue -eq 12){ 0 } else { [int]$HourDropdownStart.SelectedValue }
                    $StartDateTime = (Get-Date $InputStartDate.SelectedDate).AddHours($Hours2Add)
                }
                if($null -ne $EndDateTime -and $HourDropdownEnd.SelectedIndex -ne -1 -and !($HourDropdownEnd.SelectedIndex -eq 11 -and $AmPmDropdownEnd -eq 0) -and $AmPmDropdownEnd.SelectedIndex -ne -1){
                    $Hours2Add = if($AmPmDropdownEnd.SelectedValue -eq "PM" -and $HourDropdownEnd.SelectedIndex -ne 11){ [int]$HourDropdownEnd.SelectedValue + 12 } elseif($AmPmDropdownEnd.SelectedIndex -eq 0 -and $HourDropdownEnd.SelectedValue -eq 12){ 0 } else { [int]$HourDropdownEnd.SelectedValue }
                    $EndDateTime = (Get-Date $InputEndDate.SelectedDate).AddHours($Hours2Add)
                }
                if($null -ne $StartDateTime -and $StartDateTime -gt (Get-Date)){
                    $InvalidStart=1
                }
                if($null -ne $EndDateTime -and $EndDateTime -gt (Get-Date).AddDays(1)){
                    $InvalidEnd=1
                }
                if($null -ne $StartDateTime -and $null -ne $EndDateTime -and $StartDateTime -ge $EndDateTime){
                    $InvalidStart=1
                    $InvalidEnd=1
                }
                if($InvalidStart -eq 1){ $LblStartDateWarnings.Content="INVALID DATE" } else { $LblStartDateWarnings.Content="" }
                if($InvalidEnd -eq 1){ $LblEndDateWarnings.Content="INVALID DATE"} else { $LblEndDateWarnings.Content="" }
                if($InvalidStart -eq 0 -and $InvalidEnd -eq 0 -and $null -ne $StartDateTime -and $null -ne $EndDateTime){
                    $Valid = $true
                }
                return [PSCustomObject]@{
                    Valid = $Valid
                    StartDateTime = $StartDateTime
                    EndDateTime = $EndDateTime
                }
            }
            function Resolve-RequiredInputs {
                if($ListboxEvilSenders.HasItems -eq $true -and ((Resolve-DateInputs).Valid) -and $ListboxExchangeAdmins.HasItems -eq $true){
                    $BtnSubmit.Content="Begin SPAM Cleanup"
                    $BtnSubmit.IsEnabled = $true
                } else {
                    $BtnSubmit.Content="[Disabled]"
                    $BtnSubmit.IsEnabled = $false
                }
            }
            function Add-ExchAdmin {
                if((-NOT [string]::IsNullOrEmpty($TxtBtnAddExchAdmin.text))){
                    $TxtBtnAddExchAdmin.Text = "$([char]0x26EF)"
                    $Animation4txtBtnAddExchAdmin.Begin()
                    if($InputExchAdmin.Text -match "^.+@.+\..+$" -and $null -eq ($Creds.Username | Where-Object { $_ -eq $CredEntry })){
                        try {
                            if(!($NoMFA)){
                                Connect-EXOPSSession -UserPrincipalName ($InputExchAdmin.Text) -ErrorAction Stop 3>$null
                                $TestPSSession = Get-PSSession
                            } else {
                                $CredEntry = Get-Credential -Username ($InputExchAdmin.Text) -Message "Please enter password for $($InputExchAdmin.Text)"
                                $TestPSSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Authentication Basic -AllowRedirection -Credential $CredEntry -ErrorAction Stop
                            }
                            Invoke-Command -Session $TestPSSession -ScriptBlock { Get-OrganizationConfig | Select-Object Name } -ErrorAction Stop -HideComputerName
                            Remove-PSSession $TestPSSession
                            if(!($NoMFA)){
                                $Creds.Add((New-Object PSCredential($InputExchAdmin.Text,(ConvertTo-SecureString " " -AsPlainText -Force)))) | Out-Null
                            } else {
                                $Creds.Add($CredEntry) | Out-Null
                            }
                            $ListboxExchangeAdmins.Items.Add(($InputExchAdmin.text))
                            $InputExchAdmin.Clear()
                        } catch {
                            $InputExchAdmin.Text="NOSOUP4U"
                        }            
                    } else {
                        $InputExchAdmin.Text="NOSOUP4U"
                    }
                    $Animation4txtBtnAddExchAdmin.Stop()
                    $TxtBtnAddExchAdmin.Text = "--$([char]0x003E)"
                    $InputExchAdmin.Focus()
                    Resolve-RequiredInputs
                }
            }
            function Add-EvilSenders {
                if((-NOT [string]::IsNullOrEmpty($InputEvilSenderTxt.text))){
                    if($InputEvilSenderTxt.text -match "^.+@.+\..+$"){
                        $ListboxEvilSenders.Items.Add(($InputEvilSenderTxt.text).Trim())
                        $InputEvilSenderTxt.Clear()
                    } else {
                        $InputEvilSenderTxt.Text="NOSOUP4U"
                    }
                    $InputEvilSenderTxt.Focus()
                    Resolve-RequiredInputs
                }
            }
            function Add-EvilSubjects {
                if((-NOT [string]::IsNullOrEmpty($InputSubjectLinesTxt.text))){
                    $ListboxEvilSubjects.Items.Add([pscustomobject]@{DisplayName='"'+$InputSubjectLinesTxt.text+'"';ValueSubject=$InputSubjectLinesTxt.text})
                    $InputSubjectLinesTxt.Clear()
                    $InputSubjectLinesTxt.Focus()
                }
            }
            $BtnAddExchAdmin.Add_Click({
                Add-ExchAdmin
            })
            $BtnRemoveExchAdmin.Add_Click({
                while($ListboxExchangeAdmins.SelectedItems.count -gt 0){
                    $ListboxExchangeAdmins.Items.RemoveAt($ListboxExchangeAdmins.SelectedIndex)
                }
                Resolve-RequiredInputs
            })
            $InputExchAdmin.Add_KeyDown({
                param ( 
                    [Parameter(Mandatory)][Object]$sender,
                    [Parameter(Mandatory)][Windows.Input.KeyEventArgs]$e
                )
                if($e.Key -eq 'Enter'){
                    Add-ExchAdmin
                }
            })
            $BtnAddEvilSender.Add_Click({
                Add-EvilSenders
            })
            $BtnRemoveEvilSender.Add_Click({
                while($ListboxEvilSenders.SelectedItems.count -gt 0){
                    $ListboxEvilSenders.Items.RemoveAt($ListboxEvilSenders.SelectedIndex)
                }
                Resolve-RequiredInputs
            })
            $BtnAddSubject.Add_Click({
                Add-EvilSubjects
            })
            $BtnRemoveSubject.Add_Click({
                while($ListboxEvilSubjects.SelectedItems.count -gt 0){
                    $ListboxEvilSubjects.Items.RemoveAt($ListboxEvilSubjects.SelectedIndex)
                }
            })
            $InputEvilSenderTxt.Add_KeyDown({
                param ( 
                    [Parameter(Mandatory)][Object]$sender,
                    [Parameter(Mandatory)][Windows.Input.KeyEventArgs]$e
                )
                if($e.Key -eq 'Enter'){
                    Add-EvilSenders
                }
            })
            $InputSubjectLinesTxt.Add_KeyDown({
                param ( 
                    [Parameter(Mandatory)][Object]$sender,
                    [Parameter(Mandatory)][Windows.Input.KeyEventArgs]$e
                )
                if($e.Key -eq 'Enter'){
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
                if($StartDateTime.Hour -gt 12){
                    $HourDropdownStart.SelectedIndex = $StartDateTime.Hour - 13
                    $AmPmDropdownStart.SelectedIndex = 1
                } else {
                    $HourDropdownStart.SelectedIndex = $StartDateTime.Hour - 1
                    if($StartDateTime.Hour -ne 12){
                        $AmPmDropdownStart.SelectedIndex = 0
                    } else {
                        $AmPmDropdownStart.SelectedIndex = 1
                    }
                }
                if($EndDateTime.Hour -gt 12){
                    $HourDropdownEnd.SelectedIndex = $EndDateTime.Hour - 13
                    $AmPmDropdownEnd.SelectedIndex = 1
                } else {
                    $HourDropdownEnd.SelectedIndex = $EndDateTime.Hour - 1
                    if($EndDateTime.Hour -ne 12){
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
                if($StartDateTime.Hour -gt 12){
                    $HourDropdownStart.SelectedIndex = $StartDateTime.Hour - 13
                    $AmPmDropdownStart.SelectedIndex = 1
                } else {
                    $HourDropdownStart.SelectedIndex = $StartDateTime.Hour - 1
                    if($StartDateTime.Hour -ne 12){
                        $AmPmDropdownStart.SelectedIndex = 0
                    } else {
                        $AmPmDropdownStart.SelectedIndex = 1
                    }
                }
                if($EndDateTime.Hour -gt 12){
                    $HourDropdownEnd.SelectedIndex = $EndDateTime.Hour - 13
                    $AmPmDropdownEnd.SelectedIndex = 1
                } else {
                    $HourDropdownEnd.SelectedIndex = $EndDateTime.Hour - 1
                    if($EndDateTime.Hour -ne 12){
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
                if($StartDateTime.Hour -gt 12){
                    $HourDropdownStart.SelectedIndex = $StartDateTime.Hour - 13
                    $AmPmDropdownStart.SelectedIndex = 1
                } else {
                    $HourDropdownStart.SelectedIndex = $StartDateTime.Hour - 1
                    if($StartDateTime.Hour -ne 12){
                        $AmPmDropdownStart.SelectedIndex = 0
                    } else {
                        $AmPmDropdownStart.SelectedIndex = 1
                    }
                }
                if($EndDateTime.Hour -gt 12){
                    $HourDropdownEnd.SelectedIndex = $EndDateTime.Hour - 13
                    $AmPmDropdownEnd.SelectedIndex = 1
                } else {
                    $HourDropdownEnd.SelectedIndex = $EndDateTime.Hour - 1
                    if($EndDateTime.Hour -ne 12){
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
                $InputExchAdmin.Focus()
            })
            $BtnSubmit.Add_Click({
                $Window.DialogResult = $true
                $Window.Close()
            })
            $GUIGood = $Window.ShowDialog()
            $Senders = ($ListboxEvilSenders.Items -join ",")
            $SubjectLineStr = $null
            if($ListboxEvilSubjects.Items.Count -gt 0){ 
                $SubjectLineStr = '(Subject:'
                foreach($subject in $ListboxEvilSubjects.Items.ValueSubject){
                    $SubjectLineStr += '"'+$subject.Replace('"','""')+'" OR Subject:'
                }    
                $SubjectLineStr = $SubjectLineStr.TrimEnd(" OR Subject:")
                $SubjectLineStr += ')'
            }
            $DatesTimes = Resolve-DateInputs
            $StartDateTime = $DatesTimes.StartDateTime
            $EndDateTime = $DatesTimes.EndDateTime
            $EmailStatus = if($textRadioEmailStatus.text -eq "Both"){
                "Pending","Delivered","FilteredAsSpam"
            } elseif($textRadioEmailStatus.text -eq "Delivered"){
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
                Creds = $Creds
                EmailStatus = $EmailStatus
            }
        }
        $GUIData = Get-GUIData -NoMFA:$NoMFA
        if($GUIData.GUIGood){
            #Set Creds to collection of creds captured
            $Creds = $GUIData.Creds

            #SET SENDERS TO AN ARRAY JOINED BY COMMAS
            $Senders = $GUIData.Senders

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
            if($Senders.IndexOf(",") -gt 0){
                $SearchQuery += "("
                foreach($EvilSender in $Senders.Split(",")){
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
            if ($null -ne $SubjectLineStr){
                $SearchQuery += " AND "+$SubjectLineStr
            }
        } else {
            read-host "GUI returned an error, press any key to stop"
            try {
                Remove-PSSession $Session -ErrorAction SilentlyContinue
            } catch {
                
            }
            exit
        }
    } else {
        do {
            if($Creds.Count -eq 0){
                $CredEntry = Read-Host "[Required]Email address of Exchange Admin account to use for this script. This prompt will repeat until you press enter with no information."
            } else {
                $CredEntry = Read-Host "[Optional]Email address of another Exchange Admin account to speed up the process. This prompt will repeat until you press enter with no information."
            }
            if($CredEntry.Length -gt 0){
                if($CredEntry -match "^.+@.+\..+$" -and $null -eq ($Creds.Username | Where-Object { $_ -eq $CredEntry })){
                    if(!($NoMFA)){
                        Connect-EXOPSSession -UserPrincipalName $CredEntry 3>$null
                        $TestPSSession = Get-PSSession
                    } else {
                        $CredEntry = Get-Credential -UserName $CredEntry -Message "Please enter password for $CredEntry"
                        Write-Host "  Attempting to connect to O365 and verify this is an Exchange admin"
                        $TestPSSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $CredEntry -Authentication Basic -AllowRedirection -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                    }
                    try {
                        Invoke-Command -Session $TestPSSession -ScriptBlock { Get-OrganizationConfig | Select-Object Name } -ErrorAction Stop
                        Remove-PSSession $TestPSSession
                        if(!($NoMFA)){
                            $Creds.Add((New-Object PSCredential($CredEntry,(ConvertTo-SecureString " " -AsPlainText -Force)))) | Out-Null
                        } else {
                            $Creds.Add($CredEntry) | Out-Null
                        }
                        Write-Host "  Successful, prompting for another one."
                    } catch {
                        Write-Host "There was an error connecting to O365: Not an admin, account cannot use basic auth, bad password, or bad email"
                    }
                } else {
                    Write-Host "That was not a valid entry, try again"
                }
            }
        } until ($CredEntry.Length -eq 0 -and $Creds.Count -gt 0)
        $SendersHash = @{}
        do {
            $BadSender = read-host "[Required]Email address of evil sender. This prompt will repeat until you press enter with no information. Do not enter quotes or empty, extra spaces"
            do {
                #Validate that each email address at least matches the typical email pattern
                $SendersGood = 1
                if($BadSender.length -gt 0){
                    if(!($BadSender -match ".*@.*\..*")){
                        $SendersGood = 0
                        $BadSender = read-host "Bad email address input, try again"
                    } else {
                        $SendersHash.add($BadSender,$null)
                    }
                }
            } until ($SendersGood -eq 1)
        } until ($BadSender.length -eq 0)
        $Senders = $SendersHash.keys -join ","

        #Get the start date and time of the range when that users should have received the spam
        $StartDate = read-host "[Required]Start Date AND TIME for range when users received message (e.g. 7/18/2018 12:20 AM)"
        do {
            $good = 0
            try {
                $StartDate = Get-Date($StartDate)
                $good = 1
            }
            catch {
                $StartDate = read-host "Start date invalid, try again"
            }
        } until ($good -eq 1)
        $SearchStartDate = (get-date $StartDate).ToUniversalTime()

        #Get the end date and time of the range that users should have received the spam
        $EndDate = read-host "[Required]End Date AND TIME for range when users received message (e.g. 7/18/2018 3:59 PM)"
        do {
            $good = 0
            try {
                $EndDate = Get-Date($EndDate)
                if($EndDate -gt $StartDate){ $good = 1 } else { throw }
            }
            catch {
                $EndDate = read-host "End date invalid or before start date, try again"
            }
        } until ($good -eq 1)
        $SearchEndDate = (get-date $EndDate).ToUniversalTime()

        #Get the subject line(s) of the evil messages to filter
        $SubjectLineFilter = @{}
        do {
            $SubjectLine = read-host "[Optional]Subject line to filter. This prompt will repeat until you press enter with no information. Do not enter any quotes or backticks unless actually in the subject"
            if($SubjectLine.length -gt 0){
                $SubjectLineFilter.add($SubjectLine,$null)
            }
        } until ($SubjectLine.length -eq 0)
        $SubjectLineStr = $null
        if($SubjectLineFilter.count -gt 0){
            $SubjectLineStr = '(Subject:'
            foreach($subject in $SubjectLineFilter.keys){
                $SubjectLineStr += '"'+$subject.Replace('"','""')+'" OR Subject:'
            }
            $SubjectLineStr = $SubjectLineStr.TrimEnd(" OR Subject:")
            $SubjectLineStr += ")"
        }

        #Get the desired Email status - Delivered, FilteredAsSpam, or both
        $EmailStatusFilter = @("Pending")
        do{
            try {
                $StatusInput = [int](Read-Host "[Required]Status of emails being processed: `n[1]Delivered, [2]FilteredAsSpam, [3]Both`n(Default is 3):" -ErrorAction Stop)
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
        if($Senders.IndexOf(",") -gt 0){
            $SearchQuery += "("
            foreach($EvilSender in $Senders.Split(",")){
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
        if($null -ne $SubjectLineStr) {
            $SearchQuery += " AND "+$SubjectLineStr
        }
    }
}

#If a string of spam recipients has not been supplied, we perform a message trace to get them; otherwise, they were supplied - probably by this script
if($Recipients.length -eq 0){
    $MyRecipients = @{}
    $Page = 0
    $SearchStartDateStr = $SearchStartDate.ToString()
    $SearchEndDateStr = $SearchEndDate.ToString()
    Write-Host "Connecting to Exchange Online..."
    try {
        if(!($NoMFA)){
            Connect-EXOPSSession -UserPrincipalName ($Creds[0].UserName) 3>$null
            $Session = Get-PSSession
        } else {
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid" -Authentication Basic -AllowRedirection -Credential ($Creds[0])
        }
    } catch {
        Read-Host "There was an error connecting to Exchange Online. Press any key to close."
        Exit
    }
    Write-Host "Done."
    Write-Host "Getting recipients of evil message... from $SearchStartDateStr (UTC) to $SearchEndDateStr (UTC) from $SenderList"
    #Since Message Traces cut off after 5k results, we use PageSize to limit it to 5k users and try another page of results till we run out
    do {
        $Page++
        Write-Host "  Getting Page $Page of results, can take up to 5 minutes..."
        $a = (Invoke-Command -Session $Session -ScriptBlock { Get-MessageTrace -SenderAddress $Using:SenderList -StartDate $Using:SearchStartDateStr -EndDate $Using:SearchEndDateStr -Pagesize 5000 -Page $Using:Page -Status $Using:EmailStatusFilter -ErrorAction Stop | select-object recipientaddress} -HideComputerName).recipientaddress
        Write-Host "  Done."
        if($null -ne $a){
            #For every person found in the trace, we look to make sure it is not already in the list and that it is an LU address to which we can actually do something
            foreach($Recipient in $a){
                if(!($MyRecipients.ContainsKey($Recipient)) -and $Recipient.IndexOf("@$EmailDomain") -gt 0){
                    $MyRecipients.Add($Recipient,$null)
                }
            }
        }
        #Just because these searches can be resource-intensive and occasionally freak out, wait a second before trying again
        Start-Sleep -Seconds 1
    } until ($null -eq $a)
    Write-Host Done
    #When done, we want to use Remove-PSSession to make sure that we properly close our powershell o365 sessions
    Remove-PSSession $Session

    #Test that you are running this as a script (needed to spawn child windows)
    Write-Host Testing that you ran this as a script and did not copy paste it...
    try{
        $ScriptPath = $MyInvocation.MyCommand.Definition
        Resolve-Path $ScriptPath | out-null
    } catch {
        Write-Host STOPPING - You are not running this as a script. Press any key to close...
        Exit;
    }
    Write-host Done

    #Just for clarity to the person running the script
    Write-Host "Spawning child powershell windows with these parameters:"
    Write-Host "  Path to script file - $ScriptPath"
    Write-Host "  Received Local Date/Time -"$SearchStartDate.ToLocalTime().ToString()"to"$SearchEndDate.ToLocalTime().ToString()
    Write-Host "  Search Query (UTC DateTime) - $SearchQuery"
    Write-Host "  Total Mailboxes being processed -"$MyRecipients.Count
    $timer = [System.Diagnostics.Stopwatch]::StartNew()

    #When sending stuff with double quotes to the child powershell windows, double quotes get lost due to how powershell works. This adds a backslash to escape them so it works correctly.
    $SearchQuery = $SearchQuery.replace('"','\"')

    #Here is where we sort the list of recipients for later chunking into smaller groups
    #$MaxMailboxesProcessedPerRound tells how many mailboxes will be processed for each round of open windows. Currently, with 3 windows per credential supplied, 1500 is expected because we want 500 per window per cred
    $Mailboxes = @($MyRecipients.keys | Sort-Object | foreach-object { $_.toString() })

    #If the total number of mailboxes to process is smaller than the number of windows per session * number of credentials provided, remove some creds because the extra are pointless
    if($Mailboxes.Count -lt ($Creds.Count*$WindowsPerCred)){
        do {
            if($Creds.Count -gt 1){
                $Creds.RemoveAt(($Creds.Count)-1)
            }
        } until ($Mailboxes.Count -ge ($Creds.Count*$WindowsPerCred) -or $Creds.Count -eq 1)
    }
    
    #Figure out how many mailboxes will be processed per credential supplied.
    $MailboxesPerCred = if($Creds.Count -gt 1){[math]::Ceiling($Mailboxes.Count/$Creds.Count)} else { $Mailboxes.Count }

    #Adjust the number of mailboxes processed in each window if we have such a small number that the (number of mailboxes per window * the windows per credential) is too big
    if($MailboxesPerCred -lt ($MailboxesPerWindow*$WindowsPerCred)){ $MailboxesPerWindow = [math]::Ceiling($MailboxesPerCred/$WindowsPerCred) }
    $MaxMailboxesProcessedPerRound = $MailboxesPerWindow * $WindowsPerCred * ($Creds.Count)
    if($MaxMailboxesProcessedPerRound -gt $MyRecipients.Count){ $MaxMailboxesProcessedPerRound = $MyRecipients.Count }

    #Tell the script-runner the results of our calculations on creds, windows, and mailboxes
    Write-Host "  Total number of accounts being used - $($Creds.Count)"
    Write-Host "  Number of mailboxes per child window - $MailboxesPerWindow"
    Write-Host "  Number of mailboxes per round of child windows - $MaxMailboxesProcessedPerRound"

    #Begin the process of spawning child windows. Number of windows will be the WindowsPerCred * number of creds you provided and waits for all windows in each round to complete before attempting another
    Write-Host "Child Window Data:"
    for($m=0;$m -lt $Mailboxes.count;$m+=$MaxMailboxesProcessedPerRound){
        $RoundMinimum = $m
        $ChildWindows = $(
            for($c=0;$c -lt $Creds.Count;$c++){
                for($w=0;$w -lt $WindowsPerCred;$w++){
                    $min=$RoundMinimum
                    $max=$RoundMinimum+$MailboxesPerWindow-1
                    if($null -eq $Mailboxes[$max]){ $max = $Mailboxes.Count-1}
                    if($null -ne $Mailboxes[$min]){
                        Write-Host "  Window$c$w will be $($Mailboxes[$min]) to $($Mailboxes[$max])"
                        $u = $Creds[$c].UserName
                        $p = ConvertFrom-SecureString $Creds[$c].Password
                        $Recipients = $Mailboxes[($min)..($max)] -join ","
                        #This is the freaking magic that opens another powershell window and supplies all the values that are set as parameters up at the top of this script
                        #!!!NOTICE THE BACKTICK CHARACTERS FOR FILE AT BEGINNING AND SEARCH QUERY AT THE END! IF YOU TAKE THEM AWAY, THE SEARCH QUERY FOR THESE SPAWNED PROCESSES IS INCOMPLETE AND DELETES LOTS MORE EMAILS!!!
                        if(!($NoMFA)){
                            Start-Process powershell -Passthru -ArgumentList "-file `"$ScriptPath`" -Recipients $Recipients -CredU $u -CredP $p -SearchQuery `"$SearchQuery`""
                        } else {
                            Start-Process powershell -Passthru -ArgumentList "-file `"$ScriptPath`" -Recipients $Recipients -CredU $u -CredP $p -SearchQuery `"$SearchQuery`" -NoMFA"
                        }
                    }
                    $RoundMinimum+=$MailboxesPerWindow
                }
            }
        )
        #Wait for all of the child windows spawned up there to complete and close before opening a new set.
        $ChildWindows | Wait-Process
    }
    $timer.Stop()
    read-host "Done, the runtime for this entire process was"($timer.Elapsed.TotalMinutes)"minutes. Press any key to complete script"
    exit
} else {
    #Go ahead and process the recipients supplied to the script already
    Write-Host "Processing recipients using this query: $SearchQuery"
    $total = ([regex]::Matches($Recipients,"@")).count
    $counter = 0
    foreach ($Recipient in ($Recipients.split(",") | Sort-Object)){
        $counter++
        $errorCounter = 0
        do{
            try{
                $SearchResults = Invoke-Command -Session $Session -Scriptblock { Search-Mailbox -Identity $Using:Recipient -SearchQuery $Using:SearchQuery -deletecontent -Force -ErrorAction Stop -WarningAction SilentlyContinue } -HideComputerName -WarningVariable SearchWarning -ErrorAction Stop
                if($SearchWarning.Message -like "*exceeded*" -or $SearchWarning.Message -like "*throttl*" -or $SearchWarning.Message -like "*frequent*" -or $SearchWarning.Message -like "*The I/O Operation*"){
                    throw $SearchWarning.Message
                }
            } catch {
                $errorCounter++
                if($_.exception.message -like "*exceeded*" -or $_.exception.message -like "*throttl*" -or $_.exception.message -like "*frequent*" -or $_.exception.message -like "*The I/O Operation*"){
                    Write-Host "O365 throttling limit hit. Pausing for 5 minutes starting now $(Get-Date)"
                    Start-Sleep -Seconds 300
                    Remove-PSSession -Session $Session
                }
                if($Session.State -ne "Opened"){
                    Get-PSSession | Remove-PSSession
                    if(!($NoMFA)){
                        Connect-EXOPSSession -UserPrincipalName $CredU 3>$null
                        $Session = Get-PSSession
                    } else {
                        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Cred -Authentication Basic -AllowRedirection -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                    }
                }
            }
        } until ( $null -ne $SearchResults.ResultItemsCount -or $errorCounter -eq 2)
        if($errorCounter -eq 2){
            Write-Host "ERROR - $Recipient - $counter of $total boxes"
        } else {
            Write-Host "$($SearchResults.ResultItemsCount) item removed from $Recipient - $counter of $total boxes"
        }
    }

    #When done, we want to use Remove-PSSession to make sure that we properly close our powershell o365 sessions
    Get-PSSession | Remove-PSSession
}
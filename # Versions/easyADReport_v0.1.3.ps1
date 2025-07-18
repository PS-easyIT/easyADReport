[xml]$Global:XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="easyADReport v0.1.3" Height="875" Width="1250"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResizeWithGrip"
        Background="#F9F9F9" FontFamily="Segoe UI">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="60"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="50"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="#0078D7" BorderThickness="0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- Title und Version -->
                <StackPanel Grid.Column="0" Orientation="Horizontal" Margin="20,0">
                    <TextBlock Text="easyADReport" Foreground="White" FontSize="22" FontWeight="SemiBold" VerticalAlignment="Center"/>
                    <TextBlock Text="v0.1.1" Foreground="#E1E1E1" FontSize="14" VerticalAlignment="Center" Margin="10,0,0,0"/>
                </StackPanel>

                <!-- Ergebnisanzahl-Anzeige -->
                <Border Grid.Column="1" Background="#0063B1" Width="300" Height="50" CornerRadius="4" Margin="0,5,20,5">
                    <Grid x:Name="ResultCountGrid" Margin="5">
                        <!-- Gesamtergebnis-Anzeige -->
                        <Border Background="#0063B1" CornerRadius="2" Margin="2">
                            <Grid>
                                <TextBlock Text="SUMMARY" Foreground="White" FontSize="14" HorizontalAlignment="Center" VerticalAlignment="Top" Margin="0,5,0,0" Width="266"/>
                                <TextBlock x:Name="TotalResultCountText" Text="0" Foreground="White" FontSize="20" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Top" Margin="0,10,0,0" Width="33" TextAlignment="Right"/>
                            </Grid>
                        </Border>
                    </Grid>
                </Border>
            </Grid>
        </Border>

        <!-- Upper Section: Filter and Attribute Selection -->
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="20,15">
            <Border Background="White" CornerRadius="4" BorderBrush="#DDDDDD" BorderThickness="1" Margin="0,0,15,0" Width="620">
                <GroupBox Header="Filter" Margin="5" BorderThickness="0">
                    <StackPanel>
                        <!-- Objekttyp-Auswahl -->
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                            <Label Content="Objekttyp:" VerticalAlignment="Center" Width="80"/>
                            <RadioButton x:Name="RadioButtonUser" Content="User" IsChecked="True" Margin="5,0" VerticalAlignment="Center" />
                            <RadioButton x:Name="RadioButtonGroup" Content="Group" Margin="15,0" VerticalAlignment="Center" />
                            <RadioButton x:Name="RadioButtonComputer" Content="Computer" Margin="5,0,10,0" VerticalAlignment="Center" />
                            <RadioButton x:Name="RadioButtonGroupMemberships" Content="Mitgliedschaften" Margin="5,0" VerticalAlignment="Center" />
                        </StackPanel>
                        
                        <!-- Filter-Attribute und Werte -->
                        <StackPanel Orientation="Horizontal">
                            <Label Content="Filter Attribute:" VerticalAlignment="Center" Width="80"/>
                            <ComboBox x:Name="ComboBoxFilterAttribute" Width="150" Margin="5,0" VerticalAlignment="Center" BorderThickness="1" BorderBrush="#CCCCCC"/>
                            <Label Content="Filter Value:" VerticalAlignment="Center"/>
                            <TextBox x:Name="TextBoxFilterValue" Width="200" Margin="5,0" VerticalAlignment="Center" BorderThickness="1" BorderBrush="#CCCCCC"/>
                        </StackPanel>
                    </StackPanel>
                </GroupBox>
            </Border>
            <Border Background="White" CornerRadius="4" BorderBrush="#DDDDDD" BorderThickness="1" Margin="0,0,15,0" Width="350">
                <GroupBox Header="Attributes to Export" Margin="5" BorderThickness="0">
                    <ListBox x:Name="ListBoxSelectAttributes" Width="300" Height="80" SelectionMode="Multiple" BorderThickness="0">
                        <ListBoxItem Content="DisplayName"/>
                        <ListBoxItem Content="SamAccountName"/>
                        <ListBoxItem Content="GivenName"/>
                        <ListBoxItem Content="Surname"/>
                        <ListBoxItem Content="mail"/>
                        <ListBoxItem Content="Department"/>
                        <ListBoxItem Content="Title"/>
                        <ListBoxItem Content="Enabled"/>
                    </ListBox>
                </GroupBox>
            </Border>
            <Button x:Name="ButtonQueryAD" Content="Query" Width="145" Height="36" Margin="0,5" VerticalAlignment="Center" 
                    Background="#0078D7" Foreground="White" BorderThickness="0">
                <Button.Resources>
                    <Style TargetType="Border">
                        <Setter Property="CornerRadius" Value="4"/>
                    </Style>
                </Button.Resources>
            </Button>
        </StackPanel>

        <!-- Middle Section: Quick Reports and Results -->
        <Grid Grid.Row="2" Margin="20,0,20,15">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="220"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <Border Background="White" CornerRadius="4" BorderBrush="#DDDDDD" BorderThickness="1" Margin="0,0,10,0">
                <GroupBox Header="Quick Reports" BorderThickness="0" Margin="5,0,0,0">
                    <StackPanel>
                        <GroupBox Header="Users" BorderThickness="0" Margin="0,5,0,0">
                            <StackPanel>
                                <Button x:Name="ButtonQuickAllUsers" Content="All Users" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickDisabledUsers" Content="Disabled Users" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickLockedUsers" Content="Locked Users" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickNeverExpire" Content="Password Never Expires" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickInactiveUsers" Content="Inactive Users (90+ days)" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickAdminUsers" Content="Administrators" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                            </StackPanel>
                        </GroupBox>
                        <GroupBox Header="Groups" BorderThickness="0" Margin="0,5,0,0">
                            <StackPanel>
                                <Button x:Name="ButtonQuickGroups" Content="All Groups" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickSecurityGroups" Content="Security Groups" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickDistributionGroups" Content="Distribution Lists" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                            </StackPanel>
                        </GroupBox>
                        <GroupBox Header="Computers" BorderThickness="0" Margin="0,5,0,0">
                            <StackPanel>
                                <Button x:Name="ButtonQuickComputers" Content="All Computers" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickInactiveComputers" Content="Inactive Computers (90+ days)" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                            </StackPanel>
                        </GroupBox>
                        <GroupBox Header="Security Audit" BorderThickness="0" Margin="0,5,0,0">
                            <StackPanel>
                                <Button x:Name="ButtonQuickWeakPasswordPolicy" Content="Weak Password Policies" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickRiskyGroupMemberships" Content="Risky Memberships" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickPrivilegedAccounts" Content="Privileged Accounts" Margin="0,2" Height="20" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                            </StackPanel>
                        </GroupBox>
                        <GroupBox Header="AD-Health" FontSize="11" BorderThickness="0" Margin="0,5,0,0">
                            <StackPanel>
                                <Button x:Name="ButtonQuickFSMORoles" Content="FSMO Role Holders" Margin="0,2" FontSize="10" Height="15" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickDCStatus" Content="Domain Controller Status" Margin="0,2" FontSize="10" Height="15" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickReplicationStatus" Content="Replication Status" Margin="0,2" FontSize="10" Height="15" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                            </StackPanel>
                        </GroupBox>
                    </StackPanel>
                </GroupBox>
            </Border>

            <Border Grid.Column="1" Background="White" CornerRadius="4" BorderBrush="#DDDDDD" BorderThickness="1">
                <GroupBox BorderThickness="0" Margin="5">
                    <DataGrid x:Name="DataGridResults" AutoGenerateColumns="True" IsReadOnly="True" BorderThickness="0"
                              Background="Transparent" GridLinesVisibility="All" RowBackground="White" AlternatingRowBackground="#F5F5F5"/>
                </GroupBox>
            </Border>
        </Grid>

        <!-- Footer -->
        <Border Grid.Row="3" Background="#F0F0F0" BorderThickness="0,1,0,0" BorderBrush="#DDDDDD">
            <Grid Margin="20,0">
                <TextBlock x:Name="TextBlockStatus" Text="Ready" VerticalAlignment="Center" Margin="0,0,460,0"/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="ButtonExportCSV" Content="Export CSV" Width="100" Height="28" Margin="5,6" Background="#F0F0F0" BorderBrush="#CCCCCC">
                        <Button.Resources>
                            <Style TargetType="Border">
                                <Setter Property="CornerRadius" Value="3"/>
                            </Style>
                        </Button.Resources>
                    </Button>
                    <Button x:Name="ButtonExportHTML" Content="Export HTML" Width="100" Height="28" Margin="5,6" Background="#F0F0F0" BorderBrush="#CCCCCC">
                        <Button.Resources>
                            <Style TargetType="Border">
                                <Setter Property="CornerRadius" Value="3"/>
                            </Style>
                        </Button.Resources>
                    </Button>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Setze die Ausgabekodierung auf UTF-8, um Probleme mit Umlauten zu vermeiden
$Script:OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

# Assembly für WPF laden
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms # Für SaveFileDialog

# --- Log-Funktion für konsistente Fehlerausgabe ---
Function Write-ADReportLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Type = 'Info',
        
        [Parameter(Mandatory=$false)]
        [switch]$GUI,
        
        [Parameter(Mandatory=$false)]
        [switch]$Terminal
    )
    
    # Standardmäßig sowohl GUI als auch Terminal, wenn nicht explizit angegeben
    if (-not $GUI -and -not $Terminal) {
        $GUI = $true
        $Terminal = $true
    }
    
    # Ausgabe in der GUI
    if ($GUI -and $Global:TextBlockStatus) {
        $Global:TextBlockStatus.Text = $Message
    }
    
    # Ausgabe im Terminal
    if ($Terminal) {
        switch ($Type) {
            'Info'    { Write-Host $Message }
            'Warning' { Write-Warning $Message }
            'Error'   { Write-Error $Message }
        }
    }
}

# --- Funktion zum Abrufen von AD-Daten ---
Function Get-ADReportData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [string]$FilterAttribute,
        [Parameter(Mandatory=$false)]
        [string]$FilterValue,
        [Parameter(Mandatory=$true)]
        [System.Collections.IList]$SelectedAttributes,
        [Parameter(Mandatory=$false)]
        [string]$CustomFilter,
        [Parameter(Mandatory=$false)]
        [string]$ObjectType = "User"
    )

    # Überprüfen, ob das AD-Modul verfügbar ist
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-ADReportLog -Message "Error: Active Directory module not found." -Type Error
        return $null
    }

    try {
        # Konvertiere SelectedAttributes zu String-Array und filtere leere/null Werte
        $PropertiesToLoad = @(
            $SelectedAttributes | ForEach-Object {
                if ($null -ne $_) {
                    # Wenn es sich um ListBoxItem handelt, Content verwenden
                    if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                        $_.Content.ToString()
                    } else {
                        $_.ToString()
                    }
                }
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        # Entferne doppelte Einträge aus der Properties-Liste
        $PropertiesToLoad = $PropertiesToLoad | Select-Object -Unique

        # Sicherstellen, dass immer einige Basiseigenschaften geladen werden
        if ('DistinguishedName' -notin $PropertiesToLoad) { # Wichtig für viele Operationen
            $PropertiesToLoad += 'DistinguishedName'
        }

        if ('ObjectClass' -notin $PropertiesToLoad) { # Wichtig für Visualisierung und Typbestimmung
            $PropertiesToLoad += 'ObjectClass'
        }

        # Sicherheitsfilter: Stelle sicher, dass 'memberOf' und 'Member' nur bei GroupMemberships-Abfragen geladen werden bzw. wenn explizit ausgewählt
        if ($ObjectType -ne "GroupMemberships") {
            $PropertiesToLoadBeforeFilter = @($PropertiesToLoad)
            $PropertiesToLoad = $PropertiesToLoad | Where-Object { $_ -notin @('memberOf', 'Member') }
            if (($PropertiesToLoadBeforeFilter.Count -ne $PropertiesToLoad.Count) -or ($PropertiesToLoadBeforeFilter -join ',') -ne ($PropertiesToLoad -join ',')) {
                 Write-ADReportLog -Message "Filtered 'memberOf' or 'Member' from PropertiesToLoad for $ObjectType query. Was: $($PropertiesToLoadBeforeFilter -join ', '), IsNow: $($PropertiesToLoad -join ', ')" -Type Debug
            }
        }
        if ($PropertiesToLoad.Count -eq 0) {
            # Default properties if nothing was selected
            $PropertiesToLoad = @("DisplayName", "SamAccountName", "ObjectClass") 
        }

        $Filter = "*"
        if ($CustomFilter) {
            $Filter = $CustomFilter
        } elseif ($FilterAttribute -and $FilterValue) {
            $Filter = "$($FilterAttribute) -like '$($FilterValue)*'"
        } elseif ($FilterValue -and (-not $FilterAttribute)) {
            $Filter = "DisplayName -like '$($FilterValue)*'"
        }
        
        $Users = $null # Initialize Users

        # Special handling for LockedOut accounts
        if ($CustomFilter -and $CustomFilter.Trim() -eq "LockedOut -eq `$true") {
            Write-Host "Führe Search-ADAccount -LockedOut -UsersOnly aus, um gesperrte Benutzer zu finden."
            # Ensure AD module is loaded; this is checked at the beginning of the function.
            $LockedOutAccounts = Search-ADAccount -LockedOut -UsersOnly -ErrorAction Stop 
            
            if ($LockedOutAccounts) {
                Write-Host "$($LockedOutAccounts.Count) gesperrte(s) Konto/Konten gefunden. Rufe Details für ausgewählte Attribute ab..."
                
                $Users = foreach ($Account in $LockedOutAccounts) {
                    try {
                        Get-ADUser -Identity $Account.SamAccountName -Properties $PropertiesToLoad -ErrorAction SilentlyContinue
                    } catch {
                        Write-Warning "Konnte Details für Benutzer $($Account.SamAccountName) nicht abrufen: $($_.Exception.Message)"
                        $null
                    }
                }
                # Filter out any null results from failed Get-ADUser calls
                $Users = $Users | Where-Object {$_ -ne $null}

            } else {
                Write-Host "Keine gesperrten Benutzerkonten gefunden via Search-ADAccount."
            }
        } else {
            # Standard AD User Abfrage für andere Filter
            Write-Host "Executing Get-ADUser with filter '$Filter' and properties '$($PropertiesToLoad -join ', ')'"
            # Sicherstellen, dass Get-ADUser immer ein Array zurückgibt, auch bei nur einem Ergebnis
            $Users = @(Get-ADUser -Filter $Filter -Properties $PropertiesToLoad -ErrorAction Stop)
            Write-ADReportLog -Message "Users found: $($Users.Count) - Type: $($Users.GetType().FullName)" -Type Info -Terminal
        }

        if ($Users) {
            Write-Host "DEBUG: Final Select. PropertiesToLoad: $($PropertiesToLoad -join ', ')"
            return $Users
        } else {
            Write-ADReportLog -Message "No users found for the specified filter." -Type Warning
            return $null
        }
    } catch {
        $ErrorMessage = "Error in AD query: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Funktion zum Abrufen der Gruppenmitgliedschaften eines Benutzers ---
Function Get-UserGroupMemberships {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$SamAccountName
    )

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-ADReportLog -Message "Fehler: Active Directory Modul nicht gefunden." -Type Error
        return $null
    }

    try {
        $User = Get-ADUser -Identity $SamAccountName -Properties SamAccountName, Name -ErrorAction Stop # Hinzugefügt: Name für UserDisplayName
        if (-not $User) {
            Write-ADReportLog -Message "Benutzer $SamAccountName nicht gefunden." -Type Warning
            return $null
        }
        
        $Groups = Get-ADPrincipalGroupMembership -Identity $User -ErrorAction Stop | 
                  Get-ADGroup -Properties Name, SamAccountName, Description, GroupCategory, GroupScope -ErrorAction SilentlyContinue # Hinzugefügt: SamAccountName für GroupSamAccountName

        if ($Groups) {
            $GroupMemberships = $Groups | ForEach-Object {
                [PSCustomObject]@{
                    UserDisplayName = $User.Name
                    UserSamAccountName = $User.SamAccountName
                    GroupName = $_.Name
                    GroupSamAccountName = $_.SamAccountName
                    GroupDescription = $_.Description
                    GroupCategory = $_.GroupCategory
                    GroupScope = $_.GroupScope
                }
            }
            return $GroupMemberships
        } else {
            Write-ADReportLog -Message "Keine Gruppenmitgliedschaften für Benutzer $SamAccountName gefunden." -Type Info
            return [System.Collections.ArrayList]::new() # Leeres Array zurückgeben, um Fehler zu vermeiden
        }
    } catch {
        $ErrorMessage = "Fehler beim Abrufen der Gruppenmitgliedschaften fuer $($SamAccountName): $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Funktion zum Abrufen von AD-Gruppendaten ---
Function Get-ADGroupReportData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [System.Collections.IList]$SelectedAttributes,
        [Parameter(Mandatory=$false)]
        [string]$CustomFilter
    )

    # Überprüfen, ob das AD-Modul verfügbar ist
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-ADReportLog -Message "Fehler: Active Directory Modul nicht gefunden." -Type Error
        return $null
    }

    try {
        # Konvertiere SelectedAttributes zu String-Array und filtere leere/null Werte
        $PropertiesToLoad = @(
            $SelectedAttributes | ForEach-Object {
                if ($null -ne $_) {
                    # Wenn es sich um ListBoxItem handelt, Content verwenden
                    if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                        $_.Content.ToString()
                    } else {
                        $_.ToString()
                    }
                }
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        if ('ObjectClass' -notin $PropertiesToLoad) { # Wichtig für Visualisierung und Typbestimmung
            $PropertiesToLoad += 'ObjectClass'
        }

        if ($PropertiesToLoad.Count -eq 0) {
            $PropertiesToLoad = @("Name", "SamAccountName", "GroupCategory", "GroupScope", "ObjectClass") 
        }

        $FilterToUse = "*"
        if ($CustomFilter) {
            $FilterToUse = $CustomFilter
        }
        
        Write-Host "Führe Get-ADGroup mit Filter '$FilterToUse' und Eigenschaften '$($PropertiesToLoad -join ', ')' aus"
        $Groups = Get-ADGroup -Filter $FilterToUse -Properties $PropertiesToLoad -ErrorAction Stop

        if ($Groups) {
            # Erstelle Array mit den bereinigten Attributnamen für Select-Object
            $SelectAttributes = @(
                $SelectedAttributes | ForEach-Object {
                    if ($null -ne $_) {
                        if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                            $_.Content.ToString()
                        } else {
                            $_.ToString()
                        }
                    }
                } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
            
            # Verwende Select-Object, um ein Array von Objekten zu erstellen
            # Dies stellt sicher, dass wir eine IEnumerable-Sammlung zurückgeben
            $Output = $Groups | Select-Object $SelectAttributes
            return $Output
        } else {
            Write-ADReportLog -Message "No groups found for the specified filter." -Type Warning
            return $null
        }
    } catch {
        $ErrorMessage = "Error in AD group query: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}
# --- Funktionen für Computer-Abfragen ---
Function Get-ADComputerReportData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [System.Collections.IList]$SelectedAttributes,
        [Parameter(Mandatory=$false)]
        [string]$CustomFilter
    )

    # Überprüfen, ob das AD-Modul verfügbar ist
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-ADReportLog -Message "Fehler: Active Directory Modul nicht gefunden." -Type Error
        return $null
    }

    try {
        # Konvertiere SelectedAttributes zu String-Array und filtere leere/null Werte
        $PropertiesToLoad = @(
            $SelectedAttributes | ForEach-Object {
                if ($null -ne $_) {
                    # Wenn es sich um ListBoxItem handelt, Content verwenden
                    if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                        $_.Content.ToString()
                    } else {
                        $_.ToString()
                    }
                }
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        if ('ObjectClass' -notin $PropertiesToLoad) { # Wichtig für Visualisierung und Typbestimmung
            $PropertiesToLoad += 'ObjectClass'
        }

        if ($PropertiesToLoad.Count -eq 0) {
            $PropertiesToLoad = @("Name", "DNSHostName", "OperatingSystem", "Enabled", "ObjectClass") 
        }

        $FilterToUse = "*"
        if ($CustomFilter) {
            $FilterToUse = $CustomFilter
        }
        
        Write-Host "Führe Get-ADComputer mit Filter '$FilterToUse' und Eigenschaften '$($PropertiesToLoad -join ', ')' aus"
        $Computers = Get-ADComputer -Filter $FilterToUse -Properties $PropertiesToLoad -ErrorAction Stop

        if ($Computers) {
            # Erstelle Array mit den bereinigten Attributnamen für Select-Object
            $SelectAttributes = @(
                $SelectedAttributes | ForEach-Object {
                    if ($null -ne $_) {
                        if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                            $_.Content.ToString()
                        } else {
                            $_.ToString()
                        }
                    }
                } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
            
            # Verwende Select-Object, um ein Array von Objekten zu erstellen
            $Output = $Computers | Select-Object $SelectAttributes
            return $Output
        } else {
            Write-ADReportLog -Message "No computers found for the specified filter." -Type Warning
            return $null
        }
    } catch {
        $ErrorMessage = "Error in AD computer query: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Funktion zum Abrufen von Gruppenmitgliedschaftsberichten ---
Function Get-ADGroupMembershipsReportData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$FilterAttribute,

        [Parameter(Mandatory=$true)]
        [string]$FilterValue
    )

    Write-ADReportLog -Message "Fetching group membership data for filter: $FilterAttribute = $FilterValue" -Type Info -Terminal
    $ReportOutput = @()

    try {
        $TargetObjects = Get-ADObject -Filter "$FilterAttribute -like '*$FilterValue*'" -Properties DisplayName, SamAccountName, MemberOf, Member, ObjectClass -ErrorAction SilentlyContinue

        if (-not $TargetObjects) {
            Write-ADReportLog -Message "No objects found for the specified filter '$FilterAttribute -like *$FilterValue*'." -Type Warning
            return $ReportOutput
        }

        foreach ($TargetObject in $TargetObjects) {
            # Korrekte Objektklasse abrufen
            $objectClassSimple = if ($TargetObject.ObjectClass -is [array]) {
                $TargetObject.ObjectClass[-1]
            } else {
                $TargetObject.ObjectClass
            }
            
            $objDisplayName = if ($TargetObject.DisplayName) { $TargetObject.DisplayName } else { $TargetObject.Name }
            $objSamAccountName = if ($TargetObject.SamAccountName) { $TargetObject.SamAccountName } else { "N/A" }

            Write-ADReportLog -Message "Processing object: $objDisplayName (Type: $objectClassSimple)" -Type Info

            if ($objectClassSimple -eq 'user' -or $objectClassSimple -eq 'computer') {
                if ($TargetObject.MemberOf) {
                    foreach ($groupDN in $TargetObject.MemberOf) {
                        try {
                            $groupObject = Get-ADGroup -Identity $groupDN -Properties DisplayName, SamAccountName -ErrorAction Stop
                            $groupDisplayName = if ($groupObject.DisplayName) { $groupObject.DisplayName } else { $groupObject.Name }
                            $ReportOutput += [PSCustomObject]@{ 
                                ObjectName = $objDisplayName
                                ObjectSAM  = $objSamAccountName
                                ObjectType = $objectClassSimple
                                Relationship = "Ist Mitglied von"
                                RelatedObject = $groupDisplayName
                                RelatedObjectSAM = $groupObject.SamAccountName
                                RelatedObjectType = "Group"
                            }
                        } catch {
                            Write-ADReportLog -Message "Error resolving group DN '$groupDN' for object '$objDisplayName': $($_.Exception.Message)" -Type Warning
                            $ReportOutput += [PSCustomObject]@{ 
                                ObjectName = $objDisplayName
                                ObjectSAM  = $objSamAccountName
                                ObjectType = $objectClassSimple
                                Relationship = "Ist Mitglied von (Fehler)"
                                RelatedObject = $groupDN 
                                RelatedObjectSAM = "N/A"
                                RelatedObjectType = "Group (Fehler)"
                            }
                        }
                    }
                } else {
                    $ReportOutput += [PSCustomObject]@{ 
                        ObjectName = $objDisplayName
                        ObjectSAM  = $objSamAccountName
                        ObjectType = $objectClassSimple
                        Relationship = "Ist Mitglied von"
                        RelatedObject = "(Keine Gruppenmitgliedschaften)"
                        RelatedObjectSAM = "N/A"
                        RelatedObjectType = "N/A"
                    }
                }
            } elseif ($objectClassSimple -eq 'group') {
                if ($TargetObject.Member) {
                    foreach ($memberDN in $TargetObject.Member) {
                        try {
                            $memberObject = Get-ADObject -Identity $memberDN -Properties DisplayName, SamAccountName, ObjectClass -ErrorAction Stop
                            $memberObjectClassSimple = if ($memberObject.ObjectClass -is [array]) {
                                $memberObject.ObjectClass[-1]
                            } else {
                                $memberObject.ObjectClass
                            }
                            $memberName = if ($memberObject.DisplayName) { $memberObject.DisplayName } else { $memberObject.Name }
                            $memberSam = if ($memberObject.SamAccountName) { $memberObject.SamAccountName } else { "N/A" }

                            $ReportOutput += [PSCustomObject]@{ 
                                ObjectName = $objDisplayName
                                ObjectSAM  = $objSamAccountName
                                ObjectType = $objectClassSimple
                                Relationship = "Hat Mitglied"
                                RelatedObject = $memberName
                                RelatedObjectSAM = $memberSam
                                RelatedObjectType = $memberObjectClassSimple
                            }
                        } catch {
                            Write-ADReportLog -Message "Error resolving member DN '$memberDN' for group '$objDisplayName': $($_.Exception.Message)" -Type Warning
                            $ReportOutput += [PSCustomObject]@{ 
                                ObjectName = $objDisplayName
                                ObjectSAM  = $objSamAccountName
                                ObjectType = $objectClassSimple
                                Relationship = "Hat Mitglied (Fehler)"
                                RelatedObject = $memberDN 
                                RelatedObjectSAM = "N/A"
                                RelatedObjectType = "Unbekannt (Fehler)"
                            }
                        }
                    }
                } else {
                    $ReportOutput += [PSCustomObject]@{ 
                        ObjectName = $objDisplayName
                        ObjectSAM  = $objSamAccountName
                        ObjectType = $objectClassSimple
                        Relationship = "Hat Mitglied"
                        RelatedObject = "(Keine Mitglieder)"
                        RelatedObjectSAM = "N/A"
                        RelatedObjectType = "N/A"
                    }
                }
            } else {
                Write-ADReportLog -Message "Object $objDisplayName is of unhandled type '$objectClassSimple' for membership report." -Type Info
            }
        }
    } catch {
        $ErrorMessage = "Error in Get-ADGroupMembershipsReportData: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
    }
    return $ReportOutput
}

# --- Security Audit Functions ---
Function Get-WeakPasswordPolicyUsers {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing users with weak password policies..." -Type Info -Terminal
        
        # Properties relevant for comprehensive password policy analysis
        $Properties = @(
            "DisplayName", "SamAccountName", "Enabled", "PasswordNeverExpires", 
            "PasswordNotRequired", "PasswordLastSet", "LastLogonDate", "AdminCount",
            "CannotChangePassword", "SmartcardLogonRequired", "TrustedForDelegation",
            "DoesNotRequirePreAuth", "UseDESKeyOnly", "AccountExpirationDate",
            "LastBadPasswordAttempt", "BadLogonCount", "LogonCount", "PrimaryGroup",
            "MemberOf", "ServicePrincipalNames", "UserAccountControl", "LockedOut",
            "TrustedToAuthForDelegation", "AllowReversiblePasswordEncryption",
            "whenCreated", "Description", "UserPrincipalName", "DistinguishedName"
        )
        
        # Retrieve Domain Password Policy for comparisons
        $DomainPasswordPolicy = Get-ADDefaultDomainPasswordPolicy
        $MinPasswordAge = $DomainPasswordPolicy.MinPasswordAge.Days
        $MaxPasswordAge = $DomainPasswordPolicy.MaxPasswordAge.Days
        $MinPasswordLength = $DomainPasswordPolicy.MinPasswordLength
        
        Write-ADReportLog -Message "Domain Password Policy - Min Age: $MinPasswordAge days, Max Age: $MaxPasswordAge days, Min Length: $MinPasswordLength chars" -Type Info -Terminal
        
        # Load all users with relevant properties
        $AllUsers = Get-ADUser -Filter * -Properties $Properties
        Write-ADReportLog -Message "$($AllUsers.Count) users loaded for analysis..." -Type Info -Terminal
        
        # Define high-risk service account patterns
        $ServiceAccountPatterns = @("svc", "service", "app", "sql", "iis", "web", "backup", "sync", "admin", "sa")
        $TestAccountPatterns = @("test", "temp", "demo", "guest", "anonymous", "trial")
        
        # Enhanced analysis for weak password policies
        $WeakPasswordUsers = foreach ($user in $AllUsers) {
            $issues = @()
            $riskLevel = 0
            $recommendations = @()
            $securityFlags = @()
            
            # 1. Password never expires
            if ($user.PasswordNeverExpires -eq $true) {
                $issues += "Password never expires"
                $riskLevel += 3
                $recommendations += "Enable password expiration"
                $securityFlags += "NO_EXPIRY"
            }
            
            # 2. Password not required
            if ($user.PasswordNotRequired -eq $true) {
                $issues += "Password not required"
                $riskLevel += 5  # Critical risk
                $recommendations += "Enforce password requirement"
                $securityFlags += "NO_PASSWORD_REQ"
            }
            
            # 3. No password set or extremely old password
            if ($user.PasswordLastSet -eq $null) {
                $issues += "No password set"
                $riskLevel += 5
                $recommendations += "Set password immediately"
                $securityFlags += "NO_PASSWORD"
            } elseif ($user.PasswordLastSet -lt (Get-Date).AddDays(-($MaxPasswordAge * 3))) {
                $issues += "Password extremely outdated (>$($MaxPasswordAge * 3) days)"
                $riskLevel += 4
                $recommendations += "Force password reset immediately"
                $securityFlags += "ANCIENT_PASSWORD"
            } elseif ($user.PasswordLastSet -lt (Get-Date).AddDays(-($MaxPasswordAge * 2))) {
                $issues += "Password very old (>$($MaxPasswordAge * 2) days)"
                $riskLevel += 3
                $recommendations += "Schedule password reset"
                $securityFlags += "OLD_PASSWORD"
            } elseif ($user.PasswordLastSet -lt (Get-Date).AddDays(-365)) {
                $issues += "Password older than 1 year"
                $riskLevel += 2
                $recommendations += "Password reset recommended"
                $securityFlags += "STALE_PASSWORD"
            }
            
            # 4. User cannot change password
            if ($user.CannotChangePassword -eq $true) {
                $issues += "Cannot change password"
                $riskLevel += 2
                $recommendations += "Allow password changes (except for service accounts)"
                $securityFlags += "NO_CHANGE_ALLOWED"
            }
            
            # 5. Kerberos Pre-Authentication disabled (ASREPRoast attack possible)
            if ($user.DoesNotRequirePreAuth -eq $true) {
                $issues += "Kerberos Pre-Auth disabled (ASREPRoast vulnerability)"
                $riskLevel += 4
                $recommendations += "Enable Kerberos Pre-Authentication"
                $securityFlags += "ASREP_ROASTABLE"
            }
            
            # 6. Weak encryption (DES)
            if ($user.UseDESKeyOnly -eq $true) {
                $issues += "Uses weak DES encryption"
                $riskLevel += 3
                $recommendations += "Disable DES encryption"
                $securityFlags += "WEAK_ENCRYPTION"
            }
            
            # 7. Reversible password encryption
            if ($user.AllowReversiblePasswordEncryption -eq $true) {
                $issues += "Reversible password encryption enabled"
                $riskLevel += 4
                $recommendations += "Disable reversible password encryption"
                $securityFlags += "REVERSIBLE_ENCRYPTION"
            }
            
            # 8. Smartcard authentication not used for privileged accounts
            if ($user.AdminCount -eq 1 -and $user.SmartcardLogonRequired -eq $false) {
                $issues += "Privileged account without smartcard requirement"
                $riskLevel += 3
                $recommendations += "Enable smartcard authentication for admin accounts"
                $securityFlags += "ADMIN_NO_SMARTCARD"
            }
            
            # 9. Delegation for normal user accounts
            if (($user.TrustedForDelegation -eq $true -or $user.TrustedToAuthForDelegation -eq $true) -and $user.AdminCount -ne 1) {
                $issues += "Delegation enabled for standard user"
                $riskLevel += 3
                $recommendations += "Restrict delegation to service accounts only"
                $securityFlags += "UNEXPECTED_DELEGATION"
            }
            
            # 10. Excessive failed logon attempts
            if ($user.BadLogonCount -gt 10) {
                $issues += "Excessive failed logon attempts ($($user.BadLogonCount))"
                $riskLevel += 2
                $recommendations += "Investigate account for potential compromise"
                $securityFlags += "HIGH_FAILED_LOGONS"
            } elseif ($user.BadLogonCount -gt 5) {
                $issues += "Multiple failed logon attempts ($($user.BadLogonCount))"
                $riskLevel += 1
                $recommendations += "Monitor account activity"
                $securityFlags += "FAILED_LOGONS"
            }
            
            # 11. Service account without SPN (potentially misconfigured)
            $isServiceAccount = $false
            foreach ($pattern in $ServiceAccountPatterns) {
                if ($user.SamAccountName -like "*$pattern*" -or $user.DisplayName -like "*$pattern*") {
                    $isServiceAccount = $true
                    break
                }
            }
            
            if ($isServiceAccount) {
                if (-not $user.ServicePrincipalNames) {
                    $issues += "Service account without SPN"
                    $riskLevel += 1
                    $recommendations += "Configure SPN for service account"
                    $securityFlags += "SERVICE_NO_SPN"
                }
                
                # Service accounts should not be interactive
                if ($user.LastLogonDate -and $user.LastLogonDate -gt (Get-Date).AddDays(-30)) {
                    $issues += "Interactive logons detected for service account"
                    $riskLevel += 2
                    $recommendations += "Review service account usage"
                    $securityFlags += "SERVICE_INTERACTIVE"
                }
            }
            
            # 12. Never logged in accounts with high privileges
            if ($user.LogonCount -eq 0 -and $user.AdminCount -eq 1) {
                $issues += "Admin account never used"
                $riskLevel += 3
                $recommendations += "Disable unused admin account"
                $securityFlags += "UNUSED_ADMIN"
            } elseif ($user.LogonCount -eq 0 -and $user.whenCreated -lt (Get-Date).AddDays(-30)) {
                $issues += "Account never used (created >30 days ago)"
                $riskLevel += 1
                $recommendations += "Consider disabling unused account"
                $securityFlags += "NEVER_USED"
            }
            
            # 13. Inactive privileged accounts
            if ($user.AdminCount -eq 1 -and $user.LastLogonDate -and $user.LastLogonDate -lt (Get-Date).AddDays(-60)) {
                $issues += "Privileged account inactive for >60 days"
                $riskLevel += 3
                $recommendations += "Review inactive admin account"
                $securityFlags += "INACTIVE_ADMIN"
            } elseif ($user.LastLogonDate -and $user.LastLogonDate -lt (Get-Date).AddDays(-180)) {
                $issues += "Account inactive for >180 days"
                $riskLevel += 2
                $recommendations += "Consider disabling inactive account"
                $securityFlags += "LONG_INACTIVE"
            }
            
            # 14. Test/temp accounts without expiration
            $isTestAccount = $false
            foreach ($pattern in $TestAccountPatterns) {
                if ($user.SamAccountName -like "*$pattern*" -or $user.DisplayName -like "*$pattern*") {
                    $isTestAccount = $true
                    break
                }
            }
            
            if ($isTestAccount) {
                if ($user.AccountExpirationDate -eq $null) {
                    $issues += "Test/temp account without expiration date"
                    $riskLevel += 2
                    $recommendations += "Set expiration date for temporary accounts"
                    $securityFlags += "TEST_NO_EXPIRY"
                }
                
                if ($user.whenCreated -lt (Get-Date).AddDays(-90)) {
                    $issues += "Old test account (>90 days)"
                    $riskLevel += 1
                    $recommendations += "Review necessity of old test account"
                    $securityFlags += "OLD_TEST_ACCOUNT"
                }
            }
            
            # 15. Locked accounts with admin privileges
            if ($user.LockedOut -eq $true -and $user.AdminCount -eq 1) {
                $issues += "Locked privileged account"
                $riskLevel += 2
                $recommendations += "Investigate locked admin account"
                $securityFlags += "LOCKED_ADMIN"
            }
            
            # 16. Accounts with suspicious creation patterns
            if ($user.whenCreated -gt (Get-Date).AddDays(-7) -and $user.AdminCount -eq 1) {
                $issues += "Recently created admin account"
                $riskLevel += 2
                $recommendations += "Verify legitimacy of new admin account"
                $securityFlags += "NEW_ADMIN"
            }
            
            # 17. Accounts with generic or weak naming
            $weakNames = @("admin", "administrator", "user", "guest", "test", "temp", "service", "default")
            foreach ($weakName in $weakNames) {
                if ($user.SamAccountName -eq $weakName -or $user.SamAccountName -like "$weakName*") {
                    $issues += "Generic/predictable account name"
                    $riskLevel += 1
                    $recommendations += "Use non-predictable account names"
                    $securityFlags += "WEAK_NAMING"
                    break
                }
            }
            
            # Only return users with identified vulnerabilities
            if ($issues.Count -gt 0) {
                # Categorize risk level
                $riskCategory = switch ([int]$riskLevel) {
                    {$_ -ge 10} { [string]"Critical" }
                    {$_ -ge 7} { [string]"High" }
                    {$_ -ge 4} { [string]"Medium" }
                    {$_ -ge 2} { [string]"Low" }
                    default    { [string]"Minimal" }
                }
                

                # Add context information
                $contextInfo = @()
                if ($user.AdminCount -eq 1) { $contextInfo += "Privileged Account" }
                if ($user.Enabled -eq $false) { $contextInfo += "Disabled" }
                if ($user.LockedOut -eq $true) { $contextInfo += "Locked" }
                if ($user.ServicePrincipalNames) { $contextInfo += "Service Account" }
                if ($isServiceAccount) { $contextInfo += "Service Pattern" }
                if ($isTestAccount) { $contextInfo += "Test/Temp Account" }
                if ($user.SamAccountName -match "^(admin|administrator|root|sa|service|svc)") { $contextInfo += "System Account" }
                
                # Calculate compliance status
                $complianceIssues = 0
                if ($user.PasswordNeverExpires) { $complianceIssues++ }
                if ($user.PasswordNotRequired) { $complianceIssues++ }
                if ($user.DoesNotRequirePreAuth) { $complianceIssues++ }
                if ($user.UseDESKeyOnly) { $complianceIssues++ }
                if ($user.AllowReversiblePasswordEncryption) { $complianceIssues++ }
                
                $complianceStatus = if ($complianceIssues -eq 0) { "Compliant" } 
                                   elseif ($complianceIssues -le 2) { "Partially Compliant" } 
                                   else { "Non-Compliant" }
                
                # Calculate password age in days
                $passwordAge = if ($user.PasswordLastSet) { 
                    [math]::Round((New-TimeSpan -Start $user.PasswordLastSet -End (Get-Date)).TotalDays) 
                } else { "Never Set" }
                
                # Calculate account age
                $accountAge = [math]::Round((New-TimeSpan -Start $user.whenCreated -End (Get-Date)).TotalDays)
                
                # Determine urgency level
                $urgencyLevel = if ($user.AdminCount -eq 1 -and $riskLevel -ge 7) { "Immediate Action Required" }
                               elseif ($riskLevel -ge 10) { "Critical" }
                               elseif ($riskLevel -ge 7) { "Urgent" }
                               elseif ($riskLevel -ge 4) { "High Priority" }
                               elseif ($riskLevel -ge 2) { "Medium Priority" }
                               else { "Low Priority" }
                
                # Enhanced output object with comprehensive analysis
                [PSCustomObject]@{
                    DisplayName = $user.DisplayName
                    SamAccountName = $user.SamAccountName
                    UserPrincipalName = $user.UserPrincipalName
                    Enabled = $user.Enabled
                    LockedOut = $user.LockedOut
                    Context = if ($contextInfo) { $contextInfo -join ", " } else { "Standard User" }
                    PasswordLastSet = $user.PasswordLastSet
                    PasswordAge = $passwordAge
                    AccountCreated = $user.whenCreated
                    AccountAge = $accountAge
                    LastLogonDate = $user.LastLogonDate
                    LogonCount = $user.LogonCount
                    BadLogonCount = $user.BadLogonCount
                    LastBadPasswordAttempt = $user.LastBadPasswordAttempt
                    Vulnerabilities = $issues -join "; "
                    SecurityFlags = $securityFlags -join ", "
                    RiskLevel = $riskCategory
                    RiskScore = $riskLevel
                    ComplianceStatus = $complianceStatus
                    ComplianceIssues = $complianceIssues
                    UrgencyLevel = $urgencyLevel
                    Recommendations = $recommendations -join "; "
                    Description = $user.Description
                    LastAssessment = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    TotalIssuesFound = $issues.Count
                    RequiresImmediateAction = ($urgencyLevel -eq "Immediate Action Required" -or $urgencyLevel -eq "Critical")
                }
            }
        }
        
        # Enhanced statistics for logging
        $totalIssues = $WeakPasswordUsers.Count
        $criticalIssues = ($WeakPasswordUsers | Where-Object { $_.RiskLevel -eq "Critical" }).Count
        $highIssues = ($WeakPasswordUsers | Where-Object { $_.RiskLevel -eq "High" }).Count
        $mediumIssues = ($WeakPasswordUsers | Where-Object { $_.RiskLevel -eq "Medium" }).Count
        $adminIssues = ($WeakPasswordUsers | Where-Object { $_.Context -like "*Privileged Account*" }).Count
        $nonCompliant = ($WeakPasswordUsers | Where-Object { $_.ComplianceStatus -eq "Non-Compliant" }).Count
        $immediateAction = ($WeakPasswordUsers | Where-Object { $_.RequiresImmediateAction -eq $true }).Count
        $serviceAccountIssues = ($WeakPasswordUsers | Where-Object { $_.Context -like "*Service*" }).Count
        $testAccountIssues = ($WeakPasswordUsers | Where-Object { $_.Context -like "*Test*" }).Count
        
        Write-ADReportLog -Message "Password Policy Analysis completed:" -Type Info -Terminal
        Write-ADReportLog -Message "  Total: $totalIssues users with vulnerabilities" -Type Info -Terminal
        Write-ADReportLog -Message "  Risk Distribution - Critical: $criticalIssues, High: $highIssues, Medium: $mediumIssues" -Type Info -Terminal
        Write-ADReportLog -Message "  Privileged accounts affected: $adminIssues" -Type Info -Terminal
        Write-ADReportLog -Message "  Service accounts with issues: $serviceAccountIssues" -Type Info -Terminal
        Write-ADReportLog -Message "  Test/temp accounts with issues: $testAccountIssues" -Type Info -Terminal
        Write-ADReportLog -Message "  Compliance violations: $nonCompliant" -Type Info -Terminal
        Write-ADReportLog -Message "  Requiring immediate action: $immediateAction" -Type Info -Terminal
        
        return $WeakPasswordUsers
        
    } catch {
        $ErrorMessage = "Error analyzing password policies: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-RiskyGroupMemberships {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analysiere riskante Gruppenmitgliedschaften..." -Type Info -Terminal
        
        # Definiere hochprivilegierte/riskante Gruppen
        $RiskyGroups = @(
            "Domain Admins", "Domänen-Admins",
            "Enterprise Admins", "Organisations-Admins", 
            "Schema Admins", "Schema-Admins",
            "Administrators", "Administratoren",
            "Account Operators", "Konten-Operatoren",
            "Server Operators", "Server-Operatoren",
            "Backup Operators", "Sicherungs-Operatoren",
            "Print Operators", "Druck-Operatoren",
            "Replicator", "Replikations-Operator",
            "Remote Desktop Users", "Remotedesktopbenutzer",
            "Power Users", "Hauptbenutzer"
        )
        
        $RiskyUsers = @()
        
        # Für jede riskante Gruppe die Mitglieder ermitteln
        foreach ($groupName in $RiskyGroups) {
            try {
                $group = Get-ADGroup -Filter "Name -eq '$groupName'" -ErrorAction SilentlyContinue
                if ($group) {
                    $members = Get-ADGroupMember -Identity $group.SamAccountName -ErrorAction SilentlyContinue | 
                              Get-ADObject -Properties DisplayName, SamAccountName, ObjectClass -ErrorAction SilentlyContinue | 
                              Where-Object { $_.objectClass -eq "user" }
                    
                    foreach ($member in $members) {
                        try {
                            $userDetails = Get-ADUser -Identity $member.SamAccountName -Properties DisplayName, Enabled, LastLogonDate, PasswordLastSet -ErrorAction SilentlyContinue
                            
                            if ($userDetails) {
                                # Erstelle Objekt mit Risikobewertung
                                $riskUser = [PSCustomObject]@{
                                    DisplayName = $userDetails.DisplayName
                                    SamAccountName = $userDetails.SamAccountName
                                    Enabled = $userDetails.Enabled
                                    LastLogonDate = $userDetails.LastLogonDate
                                    PasswordLastSet = $userDetails.PasswordLastSet
                                    RisikoGruppe = $group.Name
                                    Risikostufe = switch ($group.Name) {
                                        {$_ -match "Domain Admins|Domänen-Admins|Enterprise Admins|Organisations-Admins|Schema Admins|Schema-Admins"} { "Kritisch" }
                                        {$_ -match "Administrators|Administratoren"} { "Hoch" }
                                        {$_ -match "Account Operators|Server Operators|Backup Operators"} { "Mittel" }
                                        default { "Niedrig" }
                                    }
                                    Empfehlung = if ($userDetails.Enabled -eq $false) { 
                                        "Deactivate Account delete from group" 
                                    } elseif ($userDetails.LastLogonDate -and $userDetails.LastLogonDate -lt (Get-Date).AddDays(-90)) { 
                                        "Check inactive account" 
                                    } else { 
                                        "Regularly review permissions" 
                                    }
                                }
                                $RiskyUsers += $riskUser
                            }
                        } catch {
                            Write-Warning "Konnte Details für Benutzer $($member.SamAccountName) nicht abrufen: $($_.Exception.Message)"
                        }
                    }
                }
            } catch {
                Write-Warning "Konnte Gruppe '$groupName' nicht finden oder darauf zugreifen: $($_.Exception.Message)"
            }
        }
        
        # Entferne Duplikate (Benutzer können in mehreren riskanten Gruppen sein)
        $UniqueRiskyUsers = $RiskyUsers | Sort-Object SamAccountName -Unique
        
        Write-ADReportLog -Message "$($UniqueRiskyUsers.Count) Benutzer mit riskanten Gruppenmitgliedschaften gefunden." -Type Info -Terminal
        return $UniqueRiskyUsers
        
    } catch {
        $ErrorMessage = "Fehler beim Analysieren der Gruppenmitgliedschaften: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-PrivilegedAccounts {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analysiere Konten mit erhöhten Rechten..." -Type Info -Terminal
        
        # Eigenschaften für privilegierte Konten
        $Properties = @(
            "DisplayName", "SamAccountName", "Enabled", "AdminCount", 
            "LastLogonDate", "PasswordLastSet", "PasswordNeverExpires",
            "ServicePrincipalNames", "TrustedForDelegation", "TrustedToAuthForDelegation"
        )
        
        # Alle Benutzer mit AdminCount = 1 (historisch privilegiert)
        $AdminCountUsers = Get-ADUser -Filter "AdminCount -eq 1" -Properties $Properties
        
        # Service-Konten (Konten mit SPNs)
        $ServiceAccounts = Get-ADUser -Filter "ServicePrincipalNames -like '*'" -Properties $Properties
        
        # Konten mit Delegierungsrechten
        $DelegationAccounts = Get-ADUser -Filter "TrustedForDelegation -eq `$true -or TrustedToAuthForDelegation -eq `$true" -Properties $Properties
        
        # Alle privilegierten Konten zusammenführen
        $AllPrivilegedAccounts = @()
        $AllPrivilegedAccounts += $AdminCountUsers
        $AllPrivilegedAccounts += $ServiceAccounts
        $AllPrivilegedAccounts += $DelegationAccounts
        
        # Duplikate entfernen und analysieren
        $UniquePrivilegedAccounts = $AllPrivilegedAccounts | Sort-Object SamAccountName -Unique | ForEach-Object {
            $account = $_
            
            # Risikofaktoren analysieren
            $riskFactors = @()
            $riskLevel = 0
            
            if ($account.AdminCount -eq 1) {
                $riskFactors += "AdminCount set"
                $riskLevel += 2
            }
            
            if ($account.ServicePrincipalNames) {
                $riskFactors += "Service-Account (SPN)"
                $riskLevel += 1
            }
            
            if ($account.TrustedForDelegation) {
                $riskFactors += "Delegation activated"
                $riskLevel += 2
            }
            
            if ($account.TrustedToAuthForDelegation) {
                $riskFactors += "Constrained Delegation"
                $riskLevel += 1
            }
            
            if ($account.PasswordNeverExpires) {
                $riskFactors += "Password never expires"
                $riskLevel += 1
            }
            
            if ($account.Enabled -eq $false) {
                $riskFactors += "Account disabled"
                $riskLevel += 3  # Deactivated privileged accounts are a high risk
            }

            if ($account.LastLogonDate -and $account.LastLogonDate -lt (Get-Date).AddDays(-90)) {
                $riskFactors += "Inactive (>90 days)"
                $riskLevel += 2
            }
            
            # Privilegien-Level bestimmen
            $privilegeLevel = "Standard"
            if ($account.AdminCount -eq 1 -and ($account.TrustedForDelegation -or $account.TrustedToAuthForDelegation)) {
                $privilegeLevel = "Critical"
            } elseif ($account.AdminCount -eq 1) {
                $privilegeLevel = "High"
            } elseif ($account.ServicePrincipalNames -or $account.TrustedForDelegation) {
                $privilegeLevel = "Medium"
            }
            
            # Risikostufe bestimmen
            $overallRisk = switch ($riskLevel) {
                {$_ -ge 5} { "Critical" }
                {$_ -ge 3} { "High" }
                {$_ -ge 2} { "Medium" }
                default { "Low" }
            }
            
            # Empfehlungen generieren
            $recommendations = @()
            if ($account.Enabled -eq $false -and $account.AdminCount -eq 1) {
                $recommendations += "Reset AdminCount for deactivated account"
            }
            if ($account.LastLogonDate -and $account.LastLogonDate -lt (Get-Date).AddDays(-90)) {
                $recommendations += "Check account usage"
            }
            if ($account.PasswordNeverExpires -and $account.AdminCount -eq 1) {
                $recommendations += "Enable password expiration"
            }
            if ($account.TrustedForDelegation) {
                $recommendations += "Review delegation rights"
            }
            
            [PSCustomObject]@{
                AdminCount = $account.AdminCount
                ServiceAccount = [bool]$account.ServicePrincipalNames
                DisplayName = $account.DisplayName
                SamAccountName = $account.SamAccountName
                Enabled = $account.Enabled
                LastLogonDate = $account.LastLogonDate
                PasswordLastSet = $account.PasswordLastSet
                PrivilegeLevel = $privilegeLevel
                RiskFactors = $riskFactors -join "; "
                Recommendations = if ($recommendations) { $recommendations -join "; " } else { "Regular review" }
                Delegation = $account.TrustedForDelegation -or $account.TrustedToAuthForDelegation
            }
        }
        
        Write-ADReportLog -Message "$($UniquePrivilegedAccounts.Count) Konten mit erhöhten Rechten gefunden." -Type Info -Terminal
        return $UniquePrivilegedAccounts
        
    } catch {
        $ErrorMessage = "Fehler beim Analysieren der privilegierten Konten: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- AD-Health Funktionen ---
Function Get-FSMORoles {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Retrieving FSMO role holders..." -Type Info -Terminal
        
        $Forest = Get-ADForest
        $Domain = Get-ADDomain
        
        $FSMORoles = @()
        
        # Forest-weite FSMO-Rollen
        $FSMORoles += [PSCustomObject]@{
            Role = "Schema Master"
            Type = "Forest-wide"
            Server = $Forest.SchemaMaster
            Domain = $Forest.Name
            Status = if (Test-Connection -ComputerName $Forest.SchemaMaster -Count 1 -Quiet) { "Online" } else { "Offline" }
            Description = "Manages the Active Directory schema"
        }

        $FSMORoles += [PSCustomObject]@{
            Role = "Domain Naming Master"
            Type = "Forest-wide"
            Server = $Forest.DomainNamingMaster
            Domain = $Forest.Name
            Status = if (Test-Connection -ComputerName $Forest.DomainNamingMaster -Count 1 -Quiet) { "Online" } else { "Offline" }
            Description = "Manages adding and removing domains"
        }
        
        # Domänen-spezifische FSMO-Rollen
        $FSMORoles += [PSCustomObject]@{
            Role = "PDC Emulator"
            Type = "Domain-specific"
            Server = $Domain.PDCEmulator
            Domain = $Domain.Name
            Status = if (Test-Connection -ComputerName $Domain.PDCEmulator -Count 1 -Quiet) { "Online" } else { "Offline" }
            Description = "Time synchronization and password changes"
        }

        $FSMORoles += [PSCustomObject]@{
            Role = "RID Master"
            Type = "Domain-specific"
            Server = $Domain.RIDMaster
            Domain = $Domain.Name
            Status = if (Test-Connection -ComputerName $Domain.RIDMaster -Count 1 -Quiet) { "Online" } else { "Offline" }
            Description = "Distributes RID pools to domain controllers"
        }

        $FSMORoles += [PSCustomObject]@{
            Role = "Infrastructure Master"
            Type = "Domain-specific"
            Server = $Domain.InfrastructureMaster
            Domain = $Domain.Name
            Status = if (Test-Connection -ComputerName $Domain.InfrastructureMaster -Count 1 -Quiet) { "Online" } else { "Offline" }
            Description = "Manages cross-domain references"
        }
        
        Write-ADReportLog -Message "$($FSMORoles.Count) FSMO roles found." -Type Info -Terminal
        return $FSMORoles
        
    } catch {
        $ErrorMessage = "Error retrieving FSMO roles: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}
Function Get-DomainControllerStatus {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Collect AD Data ..." -Type Info -Terminal
        
        # AD-Health Sammlung - ähnlich wie dxdiag
        $ADHealthReport = @()
        
        # 1. Forest-Informationen
        try {
            $Forest = Get-ADForest
            $ForestInfo = [PSCustomObject]@{
            Category = "Forest-Information"
            Parameter = "Forest Name"
            Value = $Forest.Name
            Status = "OK"
            Details = "Forest Functional Level: $($Forest.ForestMode)"
            }
            $ADHealthReport += $ForestInfo
            
            # Schema Version korrekt auslesen
            $schemaVersion = "Unknown"
            if ($Forest.SchemaVersion) {
            # Prüfe ob es eine Collection ist und hole den ersten Wert
            if ($Forest.SchemaVersion -is [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
                $schemaVersion = $Forest.SchemaVersion[0]
            } else {
                $schemaVersion = $Forest.SchemaVersion.ToString()
            }
            }
            
            $ADHealthReport += [PSCustomObject]@{
            Category = "Forest-Information"
            Parameter = "Schema Version"
            Value = $schemaVersion
            Status = "OK"
            Details = "Current schema version"
            }
        } catch {
            $ADHealthReport += [PSCustomObject]@{
            Category = "Forest-Information"
            Parameter = "Forest Access"
            Value = "Error"
            Status = "Critical"
            Details = $_.Exception.Message
            }
        }
        
        # 2. Domain-Informationen
        try {
            $Domain = Get-ADDomain
            $ADHealthReport += [PSCustomObject]@{
                Category = "Domain-Information"
                Parameter = "Domain Name"
                Value = $Domain.NetBIOSName
                Status = "OK"
                Details = "FQDN: $($Domain.DNSRoot), Level: $($Domain.DomainMode)"
            }
            
            # PDC Emulator Test
            $pdcTest = Test-Connection -ComputerName $Domain.PDCEmulator -Count 1 -Quiet -ErrorAction SilentlyContinue
            $ADHealthReport += [PSCustomObject]@{
                Category = "Domain-Information"
                Parameter = "PDC Emulator"
                Value = $Domain.PDCEmulator
                Status = if ($pdcTest) { "OK" } else { "Warning" }
                Details = if ($pdcTest) { "Reachable" } else { "Not reachable" }
            }
        } catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "Domain-Information"
                Parameter = "Domain Access"
                Value = "Error"
                Status = "Critical"
                Details = $_.Exception.Message
            }
        }
        
        # 3. Domain Controller Diagnose
        try {
            $DomainControllers = Get-ADDomainController -Filter * -ErrorAction Stop
            
            # Konvertiere zu Array und zähle dann
            $DCArray = @($DomainControllers)
            $DCCount = $DCArray.Count
            $ADHealthReport += [PSCustomObject]@{
                Category = "Domain Controller"
                Parameter = "Numbers of DCs"
                Value = $DCCount
                Status = if ($DCCount -ge 2) { "OK" } else { "Warning" }
                Details = "Recommended: Min. 2 DCs for redundancy"
            }

            foreach ($DC in $DomainControllers) {
                # Ping-Test
                $pingResult = Test-Connection -ComputerName $DC.HostName -Count 1 -Quiet -ErrorAction SilentlyContinue
                $ADHealthReport += [PSCustomObject]@{
                    Category = "Domain Controller"
                    Parameter = "DC Reachability"
                    Value = $DC.Name
                    Status = if ($pingResult) { "OK" } else { "Critical" }
                    Details = if ($pingResult) { "Ping successful" } else { "Ping failed" }
                }

                # LDAP-Port Test
                $ldapTest = Test-NetConnection -ComputerName $DC.HostName -Port 389 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                $ADHealthReport += [PSCustomObject]@{
                    Category = "Domain Controller"
                    Parameter = "LDAP Service"
                    Value = "$($DC.Name):389"
                    Status = if ($ldapTest) { "OK" } else { "Critical" }
                    Details = if ($ldapTest) { "LDAP Port reachable" } else { "LDAP Port not reachable" }
                }
                
                # Global Catalog Test
                if ($DC.IsGlobalCatalog) {
                    $gcTest = Test-NetConnection -ComputerName $DC.HostName -Port 3268 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                    $ADHealthReport += [PSCustomObject]@{
                        Category = "Domain Controller"
                        Parameter = "Global Catalog"
                        Value = "$($DC.Name):3268"
                        Status = if ($gcTest) { "OK" } else { "Warning" }
                        Details = if ($gcTest) { "GC Port reachable" } else { "GC Port not reachable    " }
                    }
                }
            }
        } catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "Domain Controller"
                Parameter = "DC Enumeration"
                Value = "Error"
                Status = "Critical"
                Details = $_.Exception.Message
            }
        }
        
        # 4. FSMO-Rollen Diagnose
        try {
            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO-Rollen"
                Parameter = "Schema Master"
                Value = $Forest.SchemaMaster
                Status = if (Test-Connection -ComputerName $Forest.SchemaMaster -Count 1 -Quiet -ErrorAction SilentlyContinue) { "OK" } else { "Critical" }
                Details = "Forest-wide role"
            }

            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO-Rollen"
                Parameter = "Domain Naming Master"
                Value = $Forest.DomainNamingMaster
                Status = if (Test-Connection -ComputerName $Forest.DomainNamingMaster -Count 1 -Quiet -ErrorAction SilentlyContinue) { "OK" } else { "Critical" }
                Details = "Forest-wide role"
            }
            
            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO-Rollen"
                Parameter = "PDC Emulator"
                Value = $Domain.PDCEmulator
                Status = if (Test-Connection -ComputerName $Domain.PDCEmulator -Count 1 -Quiet -ErrorAction SilentlyContinue) { "OK" } else { "Critical" }
                Details = "Time synchronization, password changes"
            }

            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO-Rollen"
                Parameter = "RID Master"
                Value = $Domain.RIDMaster
                Status = if (Test-Connection -ComputerName $Domain.RIDMaster -Count 1 -Quiet -ErrorAction SilentlyContinue) { "OK" } else { "Critical" }
                Details = "RID pool distribution"
            }
            
            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO-Rollen"
                Parameter = "Infrastructure Master"
                Value = $Domain.InfrastructureMaster
                Status = if (Test-Connection -ComputerName $Domain.InfrastructureMaster -Count 1 -Quiet -ErrorAction SilentlyContinue) { "OK" } else { "Critical" }
                Details = "Cross-domain references"
            }
        } catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO-Rollen"
                Parameter = "FSMO Access"
                Value = "Error"
                Status = "Critical"
                Details = $_.Exception.Message
            }
        }
        
        # 5. DNS-Diagnose
        try {
            $dnsResult = Resolve-DnsName -Name $Domain.DNSRoot -Type A -ErrorAction SilentlyContinue
            $ADHealthReport += [PSCustomObject]@{
                Category = "DNS-Diagnose"
                Parameter = "Domain DNS Resolution"
                Value = $Domain.DNSRoot
                Status = if ($dnsResult) { "OK" } else { "Warning" }
                Details = if ($dnsResult) { "DNS resolution successful" } else { "DNS resolution failed" }
            }
            
            # SRV-Records testen
            $srvTest = Resolve-DnsName -Name "_ldap._tcp.$($Domain.DNSRoot)" -Type SRV -ErrorAction SilentlyContinue
            $ADHealthReport += [PSCustomObject]@{
                Category = "DNS-Diagnose"
                Parameter = "LDAP SRV Records"
                Value = "_ldap._tcp.$($Domain.DNSRoot)"
                Status = if ($srvTest) { "OK" } else { "Warning" }
                Details = if ($srvTest) { "$($srvTest.Count) SRV Records found" } else { "No SRV Records found" }
            }
        } catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "DNS-Diagnose"
                Parameter = "DNS Test"
                Value = "Error"
                Status = "Warning"
                Details = $_.Exception.Message
            }
        }
        
        # 6. Replikations-Schnelltest
        try {
            $replPartners = Get-ADReplicationPartnerMetadata -Target (Get-ADDomainController).HostName -Partition (Get-ADDomain).DistinguishedName -ErrorAction SilentlyContinue | Select-Object -First 3
            
            if ($replPartners) {
                $recentRepl = $replPartners | Where-Object { $_.LastReplicationSuccess -gt (Get-Date).AddHours(-24) }
                $ADHealthReport += [PSCustomObject]@{
                    Category = "Replication"
                    Parameter = "Last Replication"
                    Value = "$($recentRepl.Count)/$($replPartners.Count) Partner"
                    Status = if ($recentRepl.Count -eq $replPartners.Count) { "OK" } elseif ($recentRepl.Count -gt 0) { "Warning" } else { "Critical" }
                    Details = "Replication at last 24h"
                }
            } else {
                $ADHealthReport += [PSCustomObject]@{
                    Category = "Replication"
                    Parameter = "Replication Partners"
                    Value = "None found"
                    Status = "Warning"
                    Details = "No replication partners found"
                }
            }
        } catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "Replication"
                Parameter = "Replication Test"
                Value = "Error"
                Status = "Warning"
                Details = $_.Exception.Message
            }
        }
        
        # 7. System-Gesundheit
        try {
            $ADHealthReport += [PSCustomObject]@{
                Category = "System-Info"
                Parameter = "Active Server"
                Value = $env:COMPUTERNAME
                Status = "Info"
                Details = "Ausfuehrungskontext: $($env:USERDOMAIN)\$($env:USERNAME)"
            }

            $ADHealthReport += [PSCustomObject]@{
                Category = "System-Info"
                Parameter = "PowerShell Version"
                Value = $PSVersionTable.PSVersion.ToString()
                Status = "Info"
                Details = "AD-Modul available: $(if (Get-Module -ListAvailable -Name ActiveDirectory) { 'Ja' } else { 'Nein' })"
            }

            # Zeitzone und Zeit
            $ADHealthReport += [PSCustomObject]@{
                Category = "System-Info"
                Parameter = "Systemtime"
                Value = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                Status = "Info"
                Details = "Zeitzone: UTC$((Get-TimeZone).BaseUtcOffset.ToString('hh\:mm'))"
            }
        } catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "System-Info"
                Parameter = "System-Check"
                Value = "Teilweise Fehler"
                Status = "Info"
                Details = $_.Exception.Message
            }
        }
        
        # Zusammenfassung erstellen
        $kritisch = ($ADHealthReport | Where-Object { $_.Status -eq "Critical" }).Count
        $warnungen = ($ADHealthReport | Where-Object { $_.Status -eq "Warning" }).Count
        $ok = ($ADHealthReport | Where-Object { $_.Status -eq "OK" }).Count
        
        $summary = [PSCustomObject]@{
            Category = "=== SUMMARY ==="
            Parameter = "AD-Health Status"
            Value = if ($kritisch -eq 0 -and $warnungen -eq 0) { "Healthy" } elseif ($kritisch -eq 0) { "Minor Issues" } else { "Critical Issues" }
            Status = if ($kritisch -eq 0 -and $warnungen -eq 0) { "OK" } elseif ($kritisch -eq 0) { "Warning" } else { "Critical" }
            Details = "OK: $ok, Warnings: $warnungen, Critical: $kritisch"
        }
        
        # Zusammenfassung an den Anfang setzen
        $FinalReport = @($summary) + $ADHealthReport
        
        Write-ADReportLog -Message "AD-Health Diagnose abgeschlossen: $($FinalReport.Count) Checks durchgeführt." -Type Info -Terminal
        return $FinalReport
        
    } catch {
        $ErrorMessage = "Schwerwiegender Fehler bei AD-Health Diagnose: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return @([PSCustomObject]@{
            Category = "FEHLER"
            Parameter = "AD-Health Check"
            Value = "Fehlgeschlagen"
            Status = "Kritisch"
            Details = $ErrorMessage
        })
    }
}
Function Get-ReplicationStatus {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Checking AD replication status..." -Type Info -Terminal
        
        $ReplicationData = @()
        $DomainControllers = Get-ADDomainController -Filter *
        
        # 1. Collect replication failures for all DCs
        Write-ADReportLog -Message "Checking replication failures..." -Type Info -Terminal
        foreach ($DC in $DomainControllers) {
            try {
                $ReplicationFailures = Get-ADReplicationFailure -Target $DC.HostName -ErrorAction SilentlyContinue
                
                if ($ReplicationFailures) {
                    foreach ($failure in $ReplicationFailures) {
                        $ReplicationData += [PSCustomObject]@{
                            Category = "Replication Failures"
                            SourceDC = $DC.Name
                            TargetDC = $failure.Partner
                            Partition = $failure.PartitionDN
                            Status = "ERROR"
                            ErrorType = $failure.FailureType
                            ErrorCount = $failure.FailureCount
                            FirstFailure = $failure.FirstFailureTime
                            LastFailure = $failure.LastFailureTime
                            Details = $failure.LastError
                        }
                    }
                } else {
                    # No errors for this DC
                    $ReplicationData += [PSCustomObject]@{
                        Category = "Replication Status"
                        SourceDC = $DC.Name
                        TargetDC = "All Partners"
                        Partition = "All Partitions"
                        Status = "OK"
                        ErrorType = "None"
                        ErrorCount = 0
                        FirstFailure = $null
                        LastFailure = $null
                        Details = "No replication failures found"
                    }
                }
            } catch {
                $ReplicationData += [PSCustomObject]@{
                    Category = "System Error"
                    SourceDC = $DC.Name
                    TargetDC = "N/A"
                    Partition = "N/A"
                    Status = "CRITICAL"
                    ErrorType = "Query Error"
                    ErrorCount = 1
                    FirstFailure = Get-Date
                    LastFailure = Get-Date
                    Details = "Error retrieving replication data: $($_.Exception.Message)"
                }
            }
        }
        
        # 2. Collect replication partner metadata
        Write-ADReportLog -Message "Collecting replication partner metadata..." -Type Info -Terminal
        foreach ($DC in $DomainControllers) {
            try {
                $PartnerMetadata = Get-ADReplicationPartnerMetadata -Target $DC.HostName -ErrorAction SilentlyContinue
                
                if ($PartnerMetadata) {
                    foreach ($partner in $PartnerMetadata) {
                        $timeSinceLastSync = "Unknown"
                        $syncStatus = "Unknown"
                        
                        if ($partner.LastReplicationSuccess) {
                            $timeSince = (Get-Date) - $partner.LastReplicationSuccess
                            $timeSinceLastSync = "$([math]::Round($timeSince.TotalHours, 1)) hours"
                            
                            # Status based on time since last sync
                            if ($timeSince.TotalHours -lt 1) {
                                $syncStatus = "EXCELLENT"
                            } elseif ($timeSince.TotalHours -lt 6) {
                                $syncStatus = "GOOD"
                            } elseif ($timeSince.TotalHours -lt 24) {
                                $syncStatus = "WARNING"
                            } else {
                                $syncStatus = "CRITICAL"
                            }
                        }
                        
                        $ReplicationData += [PSCustomObject]@{
                            Category = "Partner Metadata"
                            SourceDC = $DC.Name
                            TargetDC = $partner.Partner -replace ".*CN=NTDS Settings,CN=([^,]+),.*", '$1'
                            Partition = $partner.Partition
                            Status = $syncStatus
                            ErrorType = if ($partner.ConsecutiveReplicationFailures -gt 0) { "Consecutive Failures" } else { "None" }
                            ErrorCount = $partner.ConsecutiveReplicationFailures
                            FirstFailure = $partner.LastReplicationAttempt
                            LastFailure = $partner.LastReplicationSuccess
                            Details = "Last sync: $timeSinceLastSync ago | USN: $($partner.LastReplicationSuccess)"
                        }
                    }
                }
            } catch {
                Write-Warning "Error retrieving partner metadata for $($DC.Name): $($_.Exception.Message)"
            }
        }
        
        # 3. Additional replication health checks
        Write-ADReportLog -Message "Performing extended replication diagnostics..." -Type Info -Terminal
        
        # DC-to-DC connectivity tests
        foreach ($SourceDC in $DomainControllers) {
            foreach ($TargetDC in $DomainControllers) {
                if ($SourceDC.Name -ne $TargetDC.Name) {
                    try {
                        # Test LDAP connectivity
                        $ldapTest = Test-NetConnection -ComputerName $TargetDC.HostName -Port 389 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                        
                        # Test RPC connectivity (for replication)
                        $rpcTest = Test-NetConnection -ComputerName $TargetDC.HostName -Port 135 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                        
                        $connectionStatus = "OK"
                        $connectionDetails = "LDAP: OK, RPC: OK"
                        
                        if (-not $ldapTest -and -not $rpcTest) {
                            $connectionStatus = "CRITICAL"
                            $connectionDetails = "LDAP: FAILED, RPC: FAILED"
                        } elseif (-not $ldapTest) {
                            $connectionStatus = "WARNING"
                            $connectionDetails = "LDAP: FAILED, RPC: OK"
                        } elseif (-not $rpcTest) {
                            $connectionStatus = "WARNING"
                            $connectionDetails = "LDAP: OK, RPC: FAILED"
                        }
                        
                        $ReplicationData += [PSCustomObject]@{
                            Category = "DC Connectivity"
                            SourceDC = $SourceDC.Name
                            TargetDC = $TargetDC.Name
                            Partition = "Network Connectivity"
                            Status = $connectionStatus
                            ErrorType = if ($connectionStatus -ne "OK") { "Network Issue" } else { "None" }
                            ErrorCount = if ($connectionStatus -eq "CRITICAL") { 2 } elseif ($connectionStatus -eq "WARNING") { 1 } else { 0 }
                            FirstFailure = if ($connectionStatus -ne "OK") { Get-Date } else { $null }
                            LastFailure = if ($connectionStatus -ne "OK") { Get-Date } else { $null }
                            Details = $connectionDetails
                        }
                    } catch {
                        Write-Warning "Error testing connectivity between $($SourceDC.Name) and $($TargetDC.Name)"
                    }
                }
            }
        }
        
        # 4. Create replication summary
        $totalErrors = ($ReplicationData | Where-Object { $_.Status -eq "CRITICAL" -or $_.Status -eq "ERROR" }).Count
        $warnings = ($ReplicationData | Where-Object { $_.Status -eq "WARNING" }).Count
        $okCount = ($ReplicationData | Where-Object { $_.Status -eq "OK" -or $_.Status -eq "GOOD" -or $_.Status -eq "EXCELLENT" }).Count
        
        # Add summary to the beginning
        $summary = [PSCustomObject]@{
            Category = "SUMMARY"
            SourceDC = "All DCs"
            TargetDC = "All Partners"
            Partition = "Overall Status"
            Status = if ($totalErrors -eq 0 -and $warnings -eq 0) { "HEALTHY" } elseif ($totalErrors -eq 0) { "MINOR ISSUES" } else { "CRITICAL ISSUES" }
            ErrorType = "Overview"
            ErrorCount = $totalErrors
            FirstFailure = $null
            LastFailure = $null
            Details = "OK/Good: $okCount | Warnings: $warnings | Critical/Error: $totalErrors | DCs: $($DomainControllers.Count)"
        }
        
        # Final report with summary
        $FinalReplicationData = @($summary) + $ReplicationData
        
        Write-ADReportLog -Message "Replication analysis completed: $($FinalReplicationData.Count) entries found." -Type Info -Terminal
        return $FinalReplicationData
        
    } catch {
        $ErrorMessage = "Critical error during replication analysis: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return @([PSCustomObject]@{
            Category = "SYSTEM ERROR"
            SourceDC = "Unknown"
            TargetDC = "Unknown"
            Partition = "N/A"
            Status = "CRITICAL"
            ErrorType = "System Error"
            ErrorCount = 1
            FirstFailure = Get-Date
            LastFailure = Get-Date
            Details = $ErrorMessage
        })
    }
}

# Funktion zur Aktualisierung der Ergebniszähler im Header
Function Update-ResultCounters {
    param (
        [Parameter(Mandatory=$true)]
        $Results
    )
    
    Write-ADReportLog -Message "Updating result display..." -Type Info -Terminal
    
    try {
        # Determine result count
        $totalCount = 0
        
        if ($null -ne $Results) {
            $totalCount = $Results.Count
            
            # Determine object type for logging (for logging purposes only)
            if ($totalCount -gt 0) {
                $firstObject = $Results[0]
                $objectType = "Unknown"
                
                if ($firstObject.PSObject.Properties.Name -contains "ObjectClass") {
                    switch ($firstObject.ObjectClass) {
                        "user" { $objectType = "User" }
                        "computer" { $objectType = "Computer" }
                        "group" { $objectType = "Group" }
                    }
                }
                
                Write-ADReportLog -Message "$objectType results: $totalCount" -Type Info -Terminal
            } else {
                Write-ADReportLog -Message "No results found" -Type Info -Terminal
            }
        } else {
            Write-ADReportLog -Message "No results (Null)" -Type Info -Terminal
        }
        
        # Check if UI elements exist before trying to update them
        if ($null -ne $Global:TotalResultCountText) {
            $Global:TotalResultCountText.Text = $totalCount.ToString()
        }
        
        # Update the status in the status TextBlock
        if ($null -ne $Global:TextBlockStatus) {
            $Global:TextBlockStatus.Text = "Query completed. $totalCount result(s) found."
        }
    }
    catch {
        Write-ADReportLog -Message "Error updating result display: $($_.Exception.Message)" -Type Error
    }
}


# Funktion zum Initialisieren der Ergebnisanzeige
Function Initialize-ResultCounters {
    # Gesamtergebniszähler auf 0 setzen
    if ($null -ne $Global:TotalResultCountText) {
        $Global:TotalResultCountText.Text = "0"
    }
    
    # Sicherstellen, dass alle Zähler zurückgesetzt werden
    if ($null -ne $Global:UserCountText) {
        $Global:UserCountText.Text = "0"
    }
    
    if ($null -ne $Global:ComputerCountText) {
        $Global:ComputerCountText.Text = "0"
    }
    
    if ($null -ne $Global:GroupCountText) {
        $Global:GroupCountText.Text = "0"
    }
    
    # Status zurücksetzen
    if ($null -ne $Global:TextBlockStatus) {
        $Global:TextBlockStatus.Text = "Ready for query..."
    }
    
    # DataGrid leeren
    if ($null -ne $Global:DataGridResults) {
        $Global:DataGridResults.ItemsSource = $null
    }
}

# Funktion zur universellen Aktualisierung der Ergebnisanzeige und DataGrid
Function Update-ADReportResults {
    param (
        [Parameter(Mandatory=$false)]
        $Results = $null,
        
        [Parameter(Mandatory=$false)]
        [string]$StatusMessage
    )
    # Statusmeldung anzeigen, wenn vorhanden
    if (-not [string]::IsNullOrWhiteSpace($StatusMessage) -and $null -ne $Global:TextBlockStatus) {
        $Global:TextBlockStatus.Text = $StatusMessage
    }
}

    # --- Globale Variablen für UI Elemente --- 
Function Initialize-ADReportForm {
    param($XAMLContent)
    # Überprüfen, ob das Window-Objekt bereits existiert und zurücksetzen
    if ($Global:Window) {
        Remove-Variable -Name Window -Scope Global -ErrorAction SilentlyContinue
    }
    
    $reader = New-Object System.Xml.XmlNodeReader $XAMLContent
    $Global:Window = [Windows.Markup.XamlReader]::Load( $reader )

    # --- UI Elemente zu globalen Variablen zuweisen --- 
    # Objekttyp Radio Buttons
    $Global:RadioButtonUser = $Window.FindName("RadioButtonUser")
    $Global:RadioButtonGroup = $Window.FindName("RadioButtonGroup")
    $Global:RadioButtonComputer = $Window.FindName("RadioButtonComputer")
    $Global:RadioButtonGroupMemberships = $Window.FindName("RadioButtonGroupMemberships")

    # Filter und Attribute
    $Global:ComboBoxFilterAttribute = $Window.FindName("ComboBoxFilterAttribute")
    $Global:TextBoxFilterValue = $Window.FindName("TextBoxFilterValue")
    $Global:ListBoxSelectAttributes = $Window.FindName("ListBoxSelectAttributes")
    $Global:ButtonQueryAD = $Window.FindName("ButtonQueryAD")

    # Quick Reports - Benutzer
    $Global:ButtonQuickAllUsers = $Window.FindName("ButtonQuickAllUsers")
    $Global:ButtonQuickDisabledUsers = $Window.FindName("ButtonQuickDisabledUsers")
    $Global:ButtonQuickLockedUsers = $Window.FindName("ButtonQuickLockedUsers")
    $Global:ButtonQuickNeverExpire = $Window.FindName("ButtonQuickNeverExpire")
    $Global:ButtonQuickInactiveUsers = $Window.FindName("ButtonQuickInactiveUsers")
    $Global:ButtonQuickAdminUsers = $Window.FindName("ButtonQuickAdminUsers")
    
    # Quick Reports - Gruppen
    $Global:ButtonQuickGroups = $Window.FindName("ButtonQuickGroups")
    $Global:ButtonQuickSecurityGroups = $Window.FindName("ButtonQuickSecurityGroups")
    $Global:ButtonQuickDistributionGroups = $Window.FindName("ButtonQuickDistributionGroups")
    
    # Quick Reports - Computer
    $Global:ButtonQuickComputers = $Window.FindName("ButtonQuickComputers")
    $Global:ButtonQuickInactiveComputers = $Window.FindName("ButtonQuickInactiveComputers")
    
    # Security Audit
    $Global:ButtonQuickWeakPasswordPolicy = $Window.FindName("ButtonQuickWeakPasswordPolicy")
    $Global:ButtonQuickRiskyGroupMemberships = $Window.FindName("ButtonQuickRiskyGroupMemberships")
    $Global:ButtonQuickPrivilegedAccounts = $Window.FindName("ButtonQuickPrivilegedAccounts")
    
    # AD-Health
    $Global:ButtonQuickFSMORoles = $Window.FindName("ButtonQuickFSMORoles")
    $Global:ButtonQuickDCStatus = $Window.FindName("ButtonQuickDCStatus")
    $Global:ButtonQuickReplicationStatus = $Window.FindName("ButtonQuickReplicationStatus")

    # Ergebnisanzahl-Anzeige
    $Global:ResultCountGrid = $Window.FindName("ResultCountGrid")
    $Global:TotalResultCountText = $Window.FindName("TotalResultCountText")
    $Global:UserCountText = $Window.FindName("UserCountText")
    $Global:ComputerCountText = $Window.FindName("ComputerCountText")
    $Global:GroupCountText = $Window.FindName("GroupCountText")
    
    # Ergebnisse
    $Global:DataGridResults = $Window.FindName("DataGridResults")

    # Status und Export
    $Global:TextBlockStatus = $Window.FindName("TextBlockStatus")
    $Global:ButtonExportCSV = $Window.FindName("ButtonExportCSV")
    $Global:ButtonExportHTML = $Window.FindName("ButtonExportHTML")

    # --- UI Elemente füllen ---
    # Standardattribute für die Auswahl-ListBox füllen
    $DefaultAttributes = @("DisplayName", "SamAccountName", "GivenName", "Surname", "mail", "Department", "Title", "Enabled", "LastLogonDate", "whenCreated")
    $DefaultAttributes | ForEach-Object {
        $listBoxItem = New-Object System.Windows.Controls.ListBoxItem
        $listBoxItem.Content = $_ 
        $Global:ListBoxSelectAttributes.Items.Add($listBoxItem)
        if ($_ -in @("DisplayName", "SamAccountName", "Enabled")) { # Standardmäßig ausgewählte Attribute
            $Global:ListBoxSelectAttributes.SelectedItems.Add($listBoxItem)
        }
    }

    # Standardattribute für die Filter-ComboBox füllen
    $FilterableAttributes = @("DisplayName", "SamAccountName", "GivenName", "Surname", "mail", "Department", "Title")
    $FilterableAttributes | ForEach-Object { $Global:ComboBoxFilterAttribute.Items.Add($_) }
    if ($Global:ComboBoxFilterAttribute.Items.Count -gt 0) { $Global:ComboBoxFilterAttribute.SelectedIndex = 0 }

    # Sicherstellen, dass alle Zähler zurückgesetzt werden
    if ($null -ne $Global:UserCountText) {
        $Global:UserCountText.Text = "0"
    }
    
    if ($null -ne $Global:ComputerCountText) {
        $Global:ComputerCountText.Text = "0"
    }
    
    if ($null -ne $Global:GroupCountText) {
        $Global:GroupCountText.Text = "0"
    }
    
    # Status zurücksetzen
    if ($null -ne $Global:TextBlockStatus) {
        $Global:TextBlockStatus.Text = "Ready for query..."
    }
    
    # DataGrid leeren
    if ($null -ne $Global:DataGridResults) {
        $Global:DataGridResults.ItemsSource = $null
    }
}
Function Update-ResultCounters {
    param (
        [Parameter(Mandatory=$true)]
        $Results
    )
    
    Write-ADReportLog -Message "Updating result display..." -Type Info -Terminal
    
    try {
        # Determine result count
        $totalCount = 0
        
        if ($null -ne $Results) {
            $totalCount = $Results.Count
            
            # Determine object type for logging (for logging purposes only)
            if ($totalCount -gt 0) {
                $firstObject = $Results[0]
                $objectType = "Unknown"
                
                if ($firstObject.PSObject.Properties.Name -contains "ObjectClass") {
                    switch ($firstObject.ObjectClass) {
                        "user" { $objectType = "User" }
                        "computer" { $objectType = "Computer" }
                        "group" { $objectType = "Group" }
                    }
                }
                
                Write-ADReportLog -Message "$objectType results: $totalCount" -Type Info -Terminal
            } else {
                Write-ADReportLog -Message "No results found" -Type Info -Terminal
            }
        } else {
            Write-ADReportLog -Message "No results (Null)" -Type Info -Terminal
        }
        
        # Check if UI elements exist before trying to update them
        if ($null -ne $Global:TotalResultCountText) {
            $Global:TotalResultCountText.Text = $totalCount.ToString()
        }
        
        # Update the status in the status TextBlock
        if ($null -ne $Global:TextBlockStatus) {
            $Global:TextBlockStatus.Text = "Query completed. $totalCount result(s) found."
        }
    }
    catch {
        Write-ADReportLog -Message "Error updating result display: $($_.Exception.Message)" -Type Error
    }
}


# Funktion zum Initialisieren der Ergebnisanzeige
Function Initialize-ResultCounters {
    # Gesamtergebniszähler auf 0 setzen
    if ($null -ne $Global:TotalResultCountText) {
        $Global:TotalResultCountText.Text = "0"
    }
    
    # Sicherstellen, dass alle Zähler zurückgesetzt werden
    if ($null -ne $Global:UserCountText) {
        $Global:UserCountText.Text = "0"
    }
    
    if ($null -ne $Global:ComputerCountText) {
        $Global:ComputerCountText.Text = "0"
    }
    
    if ($null -ne $Global:GroupCountText) {
        $Global:GroupCountText.Text = "0"
    }
    
    # Status zurücksetzen
    if ($null -ne $Global:TextBlockStatus) {
        $Global:TextBlockStatus.Text = "Ready for query..."
    }
    
    # DataGrid leeren
    if ($null -ne $Global:DataGridResults) {
        $Global:DataGridResults.ItemsSource = $null
    }
}

# --- Visualisierungsfunktionen ---
# Funktion zum Erstellen von Mini-Donut-Charts in Canvas-Elementen
Function New-MiniDonutChart {
    param (
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.Canvas]$Canvas,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Data,  # z.B. @{"Enabled" = 25; "Disabled" = 5}
        
        [Parameter(Mandatory=$false)]
        [hashtable]$Colors = @{"Enabled" = "#4CAF50"; "Disabled" = "#F44336"; "Default" = "#2196F3"}
    )
    
    # Canvas leeren
    $Canvas.Children.Clear()
    
    # Dimensionen berechnen
    $canvasWidth = $Canvas.Width
    if (-not $canvasWidth -or $canvasWidth -eq 0) {
        $canvasWidth = 80  # Standardbreite
    }
    $canvasHeight = $Canvas.Height
    if (-not $canvasHeight -or $canvasHeight -eq 0) {
        $canvasHeight = 25  # Standardhöhe
    }
    
    $center = New-Object System.Windows.Point($canvasWidth/2, $canvasHeight/2)
    $radius = [Math]::Min($canvasWidth, $canvasHeight) / 2 - 2
    
    # Gesamtsumme berechnen
    $total = 0
    foreach ($value in $Data.Values) {
        $total += $value
    }
    
    if ($total -eq 0) {
        # Wenn keine Daten vorhanden sind, zeige einen grauen Kreis
        $ellipse = New-Object System.Windows.Shapes.Ellipse
        $ellipse.Width = $radius * 2
        $ellipse.Height = $radius * 2
        $ellipse.Fill = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(200, 200, 200))
        $Canvas.Children.Add($ellipse) | Out-Null
        [System.Windows.Controls.Canvas]::SetLeft($ellipse, $center.X - $radius)
        [System.Windows.Controls.Canvas]::SetTop($ellipse, $center.Y - $radius)
        
        # Textanzeige für "No Data"
        $textBlock = New-Object System.Windows.Controls.TextBlock
        $textBlock.Text = "$total"
        $textBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::White)
        $textBlock.FontSize = 8
        $textBlock.FontWeight = [System.Windows.FontWeights]::Bold
        $textBlock.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
        $textBlock.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $Canvas.Children.Add($textBlock) | Out-Null
        [System.Windows.Controls.Canvas]::SetLeft($textBlock, $center.X - 8)
        [System.Windows.Controls.Canvas]::SetTop($textBlock, $center.Y - 8)
        
        return
    }
    
    # Startwinkel für die Segmente
    $startAngle = 0
    
    # Für jeden Datenpunkt ein Segment zeichnen
    foreach ($key in $Data.Keys) {
        $value = $Data[$key]
        $sweepAngle = 360 * ($value / $total)
        
        # Segment nur zeichnen, wenn der Wert > 0 ist
        if ($value -gt 0) {
            # Farbe ermitteln
            $colorBrush = $null
            if ($Colors.ContainsKey($key)) {
                $colorCode = $Colors[$key]
                $colorBrush = New-Object System.Windows.Media.SolidColorBrush((ConvertFrom-HexColor $colorCode))
            }
            else {
                # Standardfarbe verwenden
                $colorCode = $Colors["Default"]
                $colorBrush = New-Object System.Windows.Media.SolidColorBrush((ConvertFrom-HexColor $colorCode))
            }
            
            # Segment zeichnen
            $segment = New-PieSegment -CenterX $center.X -CenterY $center.Y -Radius $radius -StartAngle $startAngle -SweepAngle $sweepAngle -Fill $colorBrush
            $Canvas.Children.Add($segment) | Out-Null
            
            # Startwinkel für das nächste Segment aktualisieren
            $startAngle += $sweepAngle
        }
    }
    
    # Inneres Loch für Donut-Effekt
    $innerRadius = $radius * 0.6
    $innerEllipse = New-Object System.Windows.Shapes.Ellipse
    $innerEllipse.Width = $innerRadius * 2
    $innerEllipse.Height = $innerRadius * 2
    $innerEllipse.Fill = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Transparent)
    $innerEllipse.Fill = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 99, 177)) # #0063B1
    $Canvas.Children.Add($innerEllipse) | Out-Null
    [System.Windows.Controls.Canvas]::SetLeft($innerEllipse, $center.X - $innerRadius)
    [System.Windows.Controls.Canvas]::SetTop($innerEllipse, $center.Y - $innerRadius)
    
    # Textanzeige für Gesamtzahl
    $textBlock = New-Object System.Windows.Controls.TextBlock
    $textBlock.Text = "$total"
    $textBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::White)
    $textBlock.FontSize = 8
    $textBlock.FontWeight = [System.Windows.FontWeights]::Bold
    $textBlock.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
    $textBlock.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    $Canvas.Children.Add($textBlock) | Out-Null
    [System.Windows.Controls.Canvas]::SetLeft($textBlock, $center.X - 8)
    [System.Windows.Controls.Canvas]::SetTop($textBlock, $center.Y - 8)
}

# Hilfsfunktion zum Konvertieren von Hex-Farbcodes in Color-Objekte
Function ConvertFrom-HexColor {
    param(
        [Parameter(Mandatory=$true)]
        [string]$HexColor
    )
    
    $HexColor = $HexColor.TrimStart('#')
    if ($HexColor.Length -eq 6) {
        $r = [Convert]::ToInt32($HexColor.Substring(0, 2), 16)
        $g = [Convert]::ToInt32($HexColor.Substring(2, 2), 16)
        $b = [Convert]::ToInt32($HexColor.Substring(4, 2), 16)
        return [System.Windows.Media.Color]::FromRgb($r, $g, $b)
    }
    return [System.Windows.Media.Colors]::Black  # Fallback
}

# Hilfsfunktion zum Erstellen eines Kreissegments
Function New-PieSegment {
    param (
        [double]$CenterX,
        [double]$CenterY,
        [double]$Radius,
        [double]$StartAngle,
        [double]$SweepAngle,
        [System.Windows.Media.Brush]$Fill
    )
    
    # Winkel in Radianten umrechnen
    $startAngleRad = $StartAngle * [Math]::PI / 180
    $endAngleRad = ($StartAngle + $SweepAngle) * [Math]::PI / 180
    
    # Groß-Kreis Flag (true wenn SweepAngle > 180°)
    $isLargeArc = $SweepAngle -gt 180
    
    # Punkte für den Pfad berechnen
    $startPoint = New-Object System.Windows.Point($CenterX + $Radius * [Math]::Cos($startAngleRad), $CenterY + $Radius * [Math]::Sin($startAngleRad))
    $endPoint = New-Object System.Windows.Point($CenterX + $Radius * [Math]::Cos($endAngleRad), $CenterY + $Radius * [Math]::Sin($endAngleRad))
    
    # PathFigure erstellen
    $pathFigure = New-Object System.Windows.Media.PathFigure
    $pathFigure.StartPoint = $CenterX, $CenterY
    $pathFigure.IsClosed = $true
    
    # Line zum Startpunkt
    $lineSegment1 = New-Object System.Windows.Media.LineSegment
    $lineSegment1.Point = $startPoint
    $pathFigure.Segments.Add($lineSegment1) | Out-Null
    
    # Kreisbogen zum Endpunkt
    $arcSegment = New-Object System.Windows.Media.ArcSegment
    $arcSegment.Point = $endPoint
    $arcSegment.Size = New-Object System.Windows.Size($Radius, $Radius)
    $arcSegment.IsLargeArc = $isLargeArc
    $arcSegment.SweepDirection = [System.Windows.Media.SweepDirection]::Clockwise
    $pathFigure.Segments.Add($arcSegment) | Out-Null
    
    # Line zurück zum Zentrum
    $lineSegment2 = New-Object System.Windows.Media.LineSegment
    $lineSegment2.Point = $CenterX, $CenterY
    $pathFigure.Segments.Add($lineSegment2) | Out-Null
    
    # PathGeometry erstellen
    $pathGeometry = New-Object System.Windows.Media.PathGeometry
    $pathGeometry.Figures.Add($pathFigure) | Out-Null
    
    # Path erstellen
    $path = New-Object System.Windows.Shapes.Path
    $path.Data = $pathGeometry
    $path.Fill = $Fill
    
    return $path
}

# Funktion zum Aktualisieren der Ergebnisvisualisierung
Function Update-ResultVisualization {
    param (
        [Parameter(Mandatory=$true)]
        [array]$Results
    )
    
    try {
        # Analysieren der Suchergebnisse basierend auf dem Objekttyp
        # Standardwerte für die Visualisierungen
        $userData = @{"Total" = 0}
        $computerData = @{"Total" = 0}
        $groupData = @{"Total" = 0}
        
        # Farben für die Diagramme
        $userColors = @{
            "Enabled" = "#4CAF50";
            "Disabled" = "#F44336";
            "Locked" = "#FFC107";
            "Admin" = "#2196F3";
            "Total" = "#607D8B";
            "Default" = "#9E9E9E"
        }
        
        $computerColors = @{
            "Enabled" = "#4CAF50";
            "Disabled" = "#F44336";
            "Inactive" = "#FFC107";
            "Total" = "#607D8B";
            "Default" = "#9E9E9E"
        }
        
        $groupColors = @{
            "Security" = "#FF9800";
            "Distribution" = "#9C27B0";
            "Total" = "#607D8B";
            "Default" = "#9E9E9E"
        }
        
        # Objekttyp bestimmen basierend auf Eigenschaften des ersten Elements, falls vorhanden
        if ($Results.Count -gt 0) {
            $firstObject = $Results[0]
            
            # Überprüfen auf Benutzer, Computer oder Gruppe basierend auf ObjectClass Eigenschaft
            if ($firstObject.PSObject.Properties.Name -contains "ObjectClass") {
                $objectClass = $firstObject.ObjectClass
                
                if ($objectClass -eq "user") {
                    # Benutzerdetails analysieren
                    $enabledUsers = $Results | Where-Object {$_.Enabled -eq $true}
                    $disabledUsers = $Results | Where-Object {$_.Enabled -eq $false}
                    
                    $userData = @{
                        "Total" = $Results.Count
                    }
                    
                    if ($enabledUsers.Count -gt 0) {
                        $userData["Enabled"] = $enabledUsers.Count
                    }
                    
                    if ($disabledUsers.Count -gt 0) {
                        $userData["Disabled"] = $disabledUsers.Count
                    }
                    
                    # Wenn die Eigenschaften LockedOut oder AdminCount vorhanden sind, diese auch berücksichtigen
                    if ($firstObject.PSObject.Properties.Name -contains "LockedOut") {
                        $lockedUsers = $Results | Where-Object {$_.LockedOut -eq $true}
                        if ($lockedUsers.Count -gt 0) {
                            $userData["Locked"] = $lockedUsers.Count
                        }
                    }
                    
                    if ($firstObject.PSObject.Properties.Name -contains "AdminCount") {
                        $adminUsers = $Results | Where-Object {$_.AdminCount -eq 1}
                        if ($adminUsers.Count -gt 0) {
                            $userData["Admin"] = $adminUsers.Count
                        }
                    }
                }
                elseif ($objectClass -eq "computer") {
                    # Computerdetails analysieren
                    $enabledComputers = $Results | Where-Object {$_.Enabled -eq $true}
                    $disabledComputers = $Results | Where-Object {$_.Enabled -eq $false}
                    
                    $computerData = @{
                        "Total" = $Results.Count
                    }
                    
                    if ($enabledComputers.Count -gt 0) {
                        $computerData["Enabled"] = $enabledComputers.Count
                    }
                    
                    if ($disabledComputers.Count -gt 0) {
                        $computerData["Disabled"] = $disabledComputers.Count
                    }
                    
                    # Wenn InactiveDays vorhanden ist, auch inaktive Computer anzeigen
                    if ($firstObject.PSObject.Properties.Name -contains "InactiveDays") {
                        $inactiveComputers = $Results | Where-Object {$_.InactiveDays -gt 30}
                        if ($inactiveComputers.Count -gt 0) {
                            $computerData["Inactive"] = $inactiveComputers.Count
                        }
                    }
                }
                elseif ($objectClass -eq "group") {
                    # Gruppendetails analysieren
                    if ($firstObject.PSObject.Properties.Name -contains "GroupCategory") {
                        $securityGroups = $Results | Where-Object {$_.GroupCategory -eq "Security"}
                        $distributionGroups = $Results | Where-Object {$_.GroupCategory -eq "Distribution"}
                        
                        $groupData = @{
                            "Total" = $Results.Count
                        }
                        
                        if ($securityGroups.Count -gt 0) {
                            $groupData["Security"] = $securityGroups.Count
                        }
                        
                        if ($distributionGroups.Count -gt 0) {
                            $groupData["Distribution"] = $distributionGroups.Count
                        }
                    }
                    else {
                        # Wenn GroupCategory nicht verfügbar ist, nur Gesamtzahl anzeigen
                        $groupData = @{
                            "Total" = $Results.Count
                        }
                    }
                }
                else {
                    # Für andere Objekttypen nur Gesamtzahl anzeigen
                    $userData = @{"Total" = $Results.Count}
                }
            }
            else {
                # Wenn keine ObjectClass-Eigenschaft vorhanden ist, zeige nur die Gesamtzahl an
                $userData = @{"Total" = $Results.Count}
            }
        }
        
        # Visualisierungen erstellen mit Null-Check
        if ($null -ne $Global:UserStatusCanvas) {
            New-MiniDonutChart -Canvas $Global:UserStatusCanvas -Data $userData -Colors $userColors
        }
        
        if ($null -ne $Global:ComputerStatusCanvas) {
            New-MiniDonutChart -Canvas $Global:ComputerStatusCanvas -Data $computerData -Colors $computerColors
        }
        
        if ($null -ne $Global:GroupsStatusCanvas) {
            New-MiniDonutChart -Canvas $Global:GroupsStatusCanvas -Data $groupData -Colors $groupColors
        }
        
        Write-ADReportLog -Message "Result visualization updated successfully." -Type Info -Terminal
    }
    catch {
        Write-ADReportLog -Message "Error updating result visualization: $($_.Exception.Message)" -Type Error
    }
}

# Funktion zur universellen Aktualisierung der Ergebnisanzeige und DataGrid
Function Update-ADReportResults {
    param (
        [Parameter(Mandatory=$false)]
        $Results = $null,
        
        [Parameter(Mandatory=$false)]
        [string]$StatusMessage
    )
    # Statusmeldung anzeigen, wenn vorhanden
    if (-not [string]::IsNullOrWhiteSpace($StatusMessage) -and $null -ne $Global:TextBlockStatus) {
        $Global:TextBlockStatus.Text = $StatusMessage
    }
}

    # --- Event Handler zuweisen --- 
    # Event Handler für RadioButtons zum Umschalten zwischen Objekttypen
    # Event Handler für ListBoxSelectAttributes zur Steuerung der Attributauswahl
    # Define the SelectionChanged event handler script block
    $script:ListBoxSelectionChangedHandler = {
        param($sourceControl, $e)

        # Get all items from the ListBox.
        $allItems = $Global:ListBoxSelectAttributes.Items
        # Get currently selected items. These are expected to be ListBoxItem objects.
        $selectedListBoxItems = $Global:ListBoxSelectAttributes.SelectedItems

        $memberOfSelected = $false
        if ($null -ne $selectedListBoxItems) {
            foreach ($selectedItem_obj in $selectedListBoxItems) {
                if ($selectedItem_obj.Content -eq "MemberOf") {
                    $memberOfSelected = $true
                    break
                }
            }
        }

        if ($memberOfSelected) {
            # "MemberOf" is selected. Disable all other items.
            foreach ($item_iter in $allItems) {

                if ($item_iter.Content -ne "MemberOf") {
                    $item_iter.IsEnabled = $false
                } else {
                    $item_iter.IsEnabled = $true # Ensure "MemberOf" itself remains enabled
                }
            }
        } else {
            # "MemberOf" is NOT selected.
            $otherAttributeSelected = $false
            if ($null -ne $selectedListBoxItems -and $selectedListBoxItems.Count -gt 0) {
                $otherAttributeSelected = $true
            }

            if ($otherAttributeSelected) {
                # Other attributes are selected. Disable "MemberOf". Enable all others.
                foreach ($item_iter in $allItems) {

                    if ($item_iter.Content -eq "MemberOf") {
                        $item_iter.IsEnabled = $false
                        if ($item_iter.IsSelected) {
                            $item_iter.IsSelected = $false
                        }
                    } else {
                        $item_iter.IsEnabled = $true
                    }
                }
            } else {
                # No attributes are selected (or only "MemberOf" was selected and now it's not).
                # Enable all attributes.
                foreach ($item_iter in $allItems) {
                    if ($item_iter -isnot [System.Windows.Controls.ListBoxItem]) {
                        Write-Warning "SelectionChanged: Item '$($item_iter)' (no selection) is not a ListBoxItem. Skipping."
                        continue
                    }
                    $item_iter.IsEnabled = $true
                }
            }
        }
    }
    # Add the event handler
    $Global:ListBoxSelectAttributes.add_SelectionChanged($script:ListBoxSelectionChangedHandler)

    $RadioButtonUser.add_Checked({
        $Global:ListBoxSelectAttributes.IsEnabled = $true
        # Funktion wird ausgeführt, wenn RadioButtonUser ausgewählt wird
        Write-ADReportLog -Message "Object type changed to User" -Type Info -Terminal
        
        # ComboBoxFilterAttribute leeren und mit benutzerspezifischen Attributen füllen
        $Global:ComboBoxFilterAttribute.Items.Clear()
        $UserFilterAttributes = @("DisplayName", "SamAccountName", "GivenName", "Surname", "mail", 
                                  "Department", "Title", "EmployeeID", "EmployeeNumber", "UserPrincipalName")
        $UserFilterAttributes | ForEach-Object { $Global:ComboBoxFilterAttribute.Items.Add($_) }
        if ($Global:ComboBoxFilterAttribute.Items.Count -gt 0) { 
            $Global:ComboBoxFilterAttribute.SelectedIndex = 0 
        }
        
        # ListBoxSelectAttributes leeren und mit benutzerspezifischen Attributen füllen
        $Global:ListBoxSelectAttributes.Items.Clear()
        $UserAttributes = @("DisplayName", "SamAccountName", "GivenName", "Surname", "mail", 
                            "Department", "Title", "Enabled", "LastLogonDate", "whenCreated", 
                            "PasswordLastSet", "PasswordNeverExpires", "LockedOut", "Description",
                            "Office", "OfficePhone", "MobilePhone", "Company")
        $UserAttributes | ForEach-Object {
            $newItem = New-Object System.Windows.Controls.ListBoxItem
            $newItem.Content = $_ 
            $Global:ListBoxSelectAttributes.Items.Add($newItem)
        }
        
        # Status aktualisieren
        $Global:TextBlockStatus.Text = "Bereit für Benutzerabfrage"
    })

    $RadioButtonGroup.add_Checked({
        $Global:ListBoxSelectAttributes.IsEnabled = $true
        # Funktion wird ausgeführt, wenn RadioButtonGroup ausgewählt wird
        Write-ADReportLog -Message "Object type changed to Group" -Type Info -Terminal
        
        # ComboBoxFilterAttribute leeren und mit gruppenspezifischen Attributen füllen
        $Global:ComboBoxFilterAttribute.Items.Clear()
        $GroupFilterAttributes = @("Name", "SamAccountName", "Description", "GroupCategory", "GroupScope")
        $GroupFilterAttributes | ForEach-Object { $Global:ComboBoxFilterAttribute.Items.Add($_) }
        if ($Global:ComboBoxFilterAttribute.Items.Count -gt 0) { 
            $Global:ComboBoxFilterAttribute.SelectedIndex = 0 
        }
        
        # ListBoxSelectAttributes leeren und mit gruppenspezifischen Attributen füllen
        $Global:ListBoxSelectAttributes.Items.Clear()
        $GroupAttributes = @("Name", "SamAccountName", "Description", "GroupCategory", "GroupScope", 
                            "whenCreated", "whenChanged", "ManagedBy", "mail", "info", "MemberOf")
        $GroupAttributes | ForEach-Object {
            $newItem = New-Object System.Windows.Controls.ListBoxItem
            $newItem.Content = $_ 
            $Global:ListBoxSelectAttributes.Items.Add($newItem)
        }
        
        # Status aktualisieren
        $Global:TextBlockStatus.Text = "Bereit für Gruppenabfrage"
    })

    $RadioButtonComputer.add_Checked({
        $Global:ListBoxSelectAttributes.IsEnabled = $true
        # Funktion wird ausgeführt, wenn RadioButtonComputer ausgewählt wird
        Write-ADReportLog -Message "Object type changed to Computer" -Type Info -Terminal
        
        # ComboBoxFilterAttribute leeren und mit computerspezifischen Attributen füllen
        $Global:ComboBoxFilterAttribute.Items.Clear()
        $ComputerFilterAttributes = @("Name", "DNSHostName", "OperatingSystem", "Description")
        $ComputerFilterAttributes | ForEach-Object { $Global:ComboBoxFilterAttribute.Items.Add($_) }
        if ($Global:ComboBoxFilterAttribute.Items.Count -gt 0) { 
            $Global:ComboBoxFilterAttribute.SelectedIndex = 0 
        }
        
        # ListBoxSelectAttributes leeren und mit computerspezifischen Attributen füllen
        $Global:ListBoxSelectAttributes.Items.Clear()
        $ComputerAttributes = @("Name", "DNSHostName", "OperatingSystem", "OperatingSystemVersion", 
                              "Enabled", "LastLogonDate", "whenCreated", "IPv4Address", "Description",
                              "Location", "ManagedBy", "PasswordLastSet")
        $ComputerAttributes | ForEach-Object {
            $newItem = New-Object System.Windows.Controls.ListBoxItem
            $newItem.Content = $_ 
            $Global:ListBoxSelectAttributes.Items.Add($newItem)
        }
        
        # Status aktualisieren
        $Global:TextBlockStatus.Text = "Bereit für Computerabfrage"
    })

    $RadioButtonGroupMemberships.add_Checked({
        Write-ADReportLog -Message "Object type changed to GroupMemberships" -Type Info -Terminal
        $Global:ListBoxSelectAttributes.IsEnabled = $false
        # Optional: ComboBoxFilterAttribute anpassen oder leeren, falls erforderlich
        # $Global:ComboBoxFilterAttribute.Items.Clear()
        # $Global:ComboBoxFilterAttribute.Items.Add("SamAccountName") # Beispiel
        # if ($Global:ComboBoxFilterAttribute.Items.Count -gt 0) { $Global:ComboBoxFilterAttribute.SelectedIndex = 0 }
        Write-ADReportLog -Message "Attribute selection disabled for GroupMemberships query." -Type Info
    })

    # Event Handler für ButtonQueryAD anpassen, um Objekttyp zu berücksichtigen
    $ButtonQueryAD.add_Click({
        Write-ADReportLog -Message "Executing query..." -Type Info
        try {
            $SelectedFilterAttribute = $Global:ComboBoxFilterAttribute.SelectedItem.ToString()
            $FilterValue = $Global:TextBoxFilterValue.Text
            $SelectedAttributes = $Global:ListBoxSelectAttributes.SelectedItems
            Write-Host "DEBUG: SelectedItems in ButtonClickHandler: $($Global:ListBoxSelectAttributes.SelectedItems | ForEach-Object {$_.Content -join '; '})"
            $isUserSearch = $Global:RadioButtonUser.IsChecked

            if ($SelectedAttributes.Count -eq 0) {
                Write-ADReportLog -Message "Please select at least one attribute for export." -Type Warning
                return
            }

            # Bestimme den aktuell ausgewählten Objekttyp
            $ObjectType = if ($Global:RadioButtonUser.IsChecked) { "User" } 
                        elseif ($Global:RadioButtonGroup.IsChecked) { "Group" } 
                        elseif ($Global:RadioButtonGroupMemberships.IsChecked) { "GroupMemberships" } 
                        else { "Computer" }
            
            # AD-Abfrage basierend auf Objekttyp durchführen
            $ReportData = $null
            switch ($ObjectType) {
                "User" {
                    $ReportData = Get-ADReportData -FilterAttribute $SelectedFilterAttribute -FilterValue $FilterValue -SelectedAttributes $SelectedAttributes
                }
                "Group" {
                    # Für Gruppen wird der Filter erst nach dem Abrufen aller Gruppen angewendet
                    $ReportData = Get-ADGroupReportData -CustomFilter "*" -SelectedAttributes $SelectedAttributes
                    if ($FilterValue -and $SelectedFilterAttribute) {
                        $ReportData = $ReportData | Where-Object { $_.$SelectedFilterAttribute -like "*$FilterValue*" }
                    }
                }
                "Computer" {
                    # Für Computer wird der Filter erst nach dem Abrufen aller Computer angewendet
                    $ReportData = Get-ADComputerReportData -CustomFilter "*" -SelectedAttributes $SelectedAttributes
                    if ($FilterValue -and $SelectedFilterAttribute) {
                        $ReportData = $ReportData | Where-Object { $_.$SelectedFilterAttribute -like "*$FilterValue*" }
                    }
                }
                "GroupMemberships" {
                    if (-not ([string]::IsNullOrWhiteSpace($FilterValue)) -and -not ([string]::IsNullOrWhiteSpace($SelectedFilterAttribute))) {
                        $ReportData = Get-ADGroupMembershipsReportData -FilterAttribute $SelectedFilterAttribute -FilterValue $FilterValue
                    } else {
                        Write-ADReportLog -Message "Filter attribute or value is empty for GroupMemberships query. Please specify a filter." -Type Warning
                        [System.Windows.Forms.MessageBox]::Show("Bitte geben Sie einen Filter (Attribut und Wert) für die Mitgliedschaftsabfrage an.", "Hinweis", "OK", "Information")
                        $ReportData = @()
                    }
                }
            }
            
            if ($ReportData) {
                try { # Inner try for processing $ReportData
                    if ($Global:DataGridResults) {
                        $Global:DataGridResults.ItemsSource = $null # Clear items first
                        if ($Global:DataGridResults.Columns) { $Global:DataGridResults.Columns.Clear() } # Clear columns
                    }
                    # Debug-Informationen
                    Write-ADReportLog -Message "ReportData Typ: $($ReportData.GetType().FullName)" -Type Info -Terminal
                    
                    # Wir brauchen sicherzustellen, dass wir immer eine Liste/Sammlung haben, auch bei einzelnen Objekten
                    # Verwende @() um es als Array zu erzwingen
                    $ReportCollection = @($ReportData)
                    
                    # Direkte Zuweisung an DataGrid
                    $Global:DataGridResults.ItemsSource = $ReportCollection
                    
                    # Zähle die Anzahl der Ergebnisse
                    $Count = $ReportCollection.Count
                    Write-ADReportLog -Message "Query completed. $Count result(s) found for $ObjectType." -Type Info
                    
                    # Ergebniszähler im Header aktualisieren
                    Update-ResultCounters -Results $ReportCollection
                    
                    if ($isUserSearch -and $ReportCollection.Count -eq 1 -and $ReportCollection[0].PSObject.Properties['SamAccountName']) {
                        $userSamAccountName = $ReportCollection[0].SamAccountName
                        
                        # Check if the "Mitgliedschaften" RadioButton is checked
                        if ($Global:RadioButtonGroupMemberships.IsChecked -eq $true) {
                            Write-ADReportLog -Message "Rufe Gruppenmitgliedschaften für Benutzer $($userSamAccountName) ab (RadioButton 'Mitgliedschaften' ist aktiv)..." -Type Info
                            $GroupMemberships = Get-UserGroupMemberships -SamAccountName $userSamAccountName
                            
                            if ($GroupMemberships -and $GroupMemberships.Count -gt 0) {
                                Write-ADReportLog -Message "$($GroupMemberships.Count) Gruppenmitgliedschaften gefunden." -Type Info
                                $DisplayData = $GroupMemberships | Select-Object @{
                                    Name = 'Benutzer';
                                    Expression = {$_.UserDisplayName}
                                }, @{
                                    Name = 'Benutzer (SAM)';
                                    Expression = {$_.UserSamAccountName}
                                }, @{
                                    Name = 'Gruppe';
                                    Expression = {$_.GroupName}
                                }, @{
                                    Name = 'Gruppe (SAM)';
                                    Expression = {$_.GroupSamAccountName}
                                }, @{
                                    Name = 'Gruppentyp';
                                    Expression = {"$($_.GroupCategory) / $($_.GroupScope)"}
                                }, @{
                                    Name = 'Beschreibung';
                                    Expression = {$_.GroupDescription}
                                }
                                $Global:DataGridResults.ItemsSource = $DisplayData
                                $Global:TextBlockStatus.Text = "Gruppenmitgliedschaften für Benutzer: $userSamAccountName"
                                # For memberships, we might not want to overwrite general counters with group details
                                # Update-ResultCounters -Results $ReportData # Keep original user count
                                # Update-ResultVisualization -Results $ReportData # Keep original visualization
                            } else {
                                # RadioButton is checked, but no memberships found or error during retrieval
                                Write-ADReportLog -Message "Benutzer $userSamAccountName gefunden. Keine Gruppenmitgliedschaften oder Fehler beim Abruf (RadioButton 'Mitgliedschaften' war aktiv)." -Type Info
                                $Global:DataGridResults.ItemsSource = $ReportCollection # Show user data as fallback
                                $Global:TextBlockStatus.Text = "Benutzer $userSamAccountName gefunden. Keine Gruppenmitgliedschaften (Abfrage aktiv)."
                                Update-ResultCounters -Results $ReportCollection
                                Update-ResultVisualization -Results $ReportCollection
                            }
                        } else {
                            # RadioButton "Mitgliedschaften" is NOT checked. Display user data as usual.
                            Write-ADReportLog -Message "RadioButton 'Mitgliedschaften' ist nicht aktiv. Zeige Benutzerdaten für $userSamAccountName." -Type Info
                            $Global:DataGridResults.ItemsSource = $ReportCollection
                            $Global:TextBlockStatus.Text = "Benutzer $userSamAccountName gefunden. Mitgliedschaften nicht abgefragt."
                            Update-ResultCounters -Results $ReportCollection
                            Update-ResultVisualization -Results $ReportCollection
                        }
                    } else {
                        # This is the original 'else' for cases other than single user search.
                        # It should remain as is, displaying the $ReportCollection.
                        $Global:DataGridResults.ItemsSource = $ReportCollection
                        Update-ResultCounters -Results $ReportCollection
                        Update-ResultVisualization -Results $ReportCollection # Ensure visualization is updated here too.
                    }
                } catch {
                    $ErrorMessage = "Error in query process: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    if ($Global:DataGridResults) {
                        $Global:DataGridResults.ItemsSource = $null
                        if ($Global:DataGridResults.Columns) { $Global:DataGridResults.Columns.Clear() }
                    }
                    Update-ResultCounters -Results @() # Leeres Array für die Zähler
                } 
            } else {
                Write-ADReportLog -Message "No data returned from query for $ObjectType." -Type Info
                if ($Global:DataGridResults) {
                    $Global:DataGridResults.ItemsSource = $null
                    if ($Global:DataGridResults.Columns) { $Global:DataGridResults.Columns.Clear() }
                }
                Update-ResultCounters -Results @()
                if ($Global:TextBlockStatus) { $Global:TextBlockStatus.Text = "No results found for $ObjectType." }
            }
        } catch { # Outer catch for the whole ButtonQueryAD.add_Click
            $OuterCatchErrorMessage = "An unexpected error occurred during query execution: $($_.Exception.Message)"
            Write-ADReportLog -Message $OuterCatchErrorMessage -Type Error
            try {
                if ($Global:DataGridResults) {
                    $Global:DataGridResults.ItemsSource = $null
                    if ($Global:DataGridResults.Columns) { $Global:DataGridResults.Columns.Clear() }
                }
                Update-ResultCounters -Results @()
                if ($Global:TextBlockStatus) { $Global:TextBlockStatus.Text = "Error: Query failed." } 
            } catch {
                Write-ADReportLog -Message "CRITICAL: Error within the main query button's outer catch block: $($_.Exception.Message)" -Type Error
            }
        }
    }) # End of .add_Click({
            $ButtonQuickAllUsers.add_Click({
        Write-ADReportLog -Message "Loading all users..." -Type Info
        try {
            if ($script:ListBoxSelectionChangedHandler) {
                $Global:ListBoxSelectAttributes.remove_SelectionChanged($script:ListBoxSelectionChangedHandler)
            }
            $Global:RadioButtonUser.IsChecked = $true # Sicherstellen, dass Benutzermodus aktiv ist
            $Global:ComboBoxFilterAttribute.SelectedIndex = -1
            $Global:TextBoxFilterValue.Text = ""
            
            $QuickReportAttributes = @("DisplayName", "SamAccountName", "GivenName", "Surname", "mail", "Enabled", "LastLogonDate", "whenCreated", "LockedOut")
            $Global:ListBoxSelectAttributes.SelectedItems.Clear()
            foreach ($attr in $QuickReportAttributes) {
                $itemsFound = $Global:ListBoxSelectAttributes.Items | Where-Object { $_.Content -eq $attr }
                if ($itemsFound) {
                    foreach ($singleItemInFoundList in $itemsFound) {
                        if ($singleItemInFoundList -is [System.Windows.Controls.ListBoxItem]) {
                            $singleItemInFoundList.IsSelected = $true
                        }
                    }
                }
            }

            $ReportData = Get-ADReportData -CustomFilter "*" -SelectedAttributes $QuickReportAttributes
            if ($null -eq $ReportData) { $ReportData = @() }
            
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
            $Global:DataGridResults.ItemsSource = $ReportData
            Update-ResultCounters -Results $ReportData
            $Global:TextBlockStatus.Text = "All users loaded. $($ReportData.Count) record(s) found."
            Write-ADReportLog -Message "All users loaded. $($ReportData.Count) result(s) found." -Type Info
        } catch {
            $ErrorMessage = "Error loading all users: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
            Update-ResultCounters -Results @()
            $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
        }
    })

    $ButtonQuickDisabledUsers.add_Click({
        Write-ADReportLog -Message "Loading disabled users..." -Type Info
        try {
            if ($script:ListBoxSelectionChangedHandler) {
                $Global:ListBoxSelectAttributes.remove_SelectionChanged($script:ListBoxSelectionChangedHandler)
            }
            $Global:RadioButtonUser.IsChecked = $true
            $Global:ComboBoxFilterAttribute.SelectedIndex = -1
            $Global:TextBoxFilterValue.Text = ""

            $QuickReportAttributes = @("DisplayName", "SamAccountName", "Enabled", "LastLogonDate")
            $Global:ListBoxSelectAttributes.SelectedItems.Clear()
            foreach ($attr in $QuickReportAttributes) {
                $itemsFound = 1 # DEBUG
                if ($itemsFound) {
                    foreach ($singleItemInFoundList in $itemsFound) {
                        if ($singleItemInFoundList -is [System.Windows.Controls.ListBoxItem]) {
                            $singleItemInFoundList.IsSelected = $true
                        }
                    }
                }
            }
            
            $ReportData = Get-ADReportData -CustomFilter "Enabled -eq `$false" -SelectedAttributes $QuickReportAttributes
            if ($null -eq $ReportData) { $ReportData = @() }
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
            $Global:DataGridResults.ItemsSource = $ReportData
            Update-ResultCounters -Results $ReportData
            $Global:TextBlockStatus.Text = "Disabled users loaded. $($ReportData.Count) record(s) found."
            Write-ADReportLog -Message "Disabled users loaded. $($ReportData.Count) result(s) found." -Type Info
        } catch {
            $ErrorMessage = "Error loading disabled users: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
            Update-ResultCounters -Results @()
            $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
        }
    })

    $ButtonQuickLockedUsers.add_Click({
        Write-ADReportLog -Message "Loading locked out users..." -Type Info
        try {
            if ($script:ListBoxSelectionChangedHandler) {
                $Global:ListBoxSelectAttributes.remove_SelectionChanged($script:ListBoxSelectionChangedHandler)
            }
            $Global:RadioButtonUser.IsChecked = $true
            $Global:ComboBoxFilterAttribute.SelectedIndex = -1
            $Global:TextBoxFilterValue.Text = ""

            $QuickReportAttributes = @("DisplayName", "SamAccountName", "LockedOut", "LastLogonDate", "BadLogonCount")
            $Global:ListBoxSelectAttributes.SelectedItems.Clear()
            foreach ($attr in $QuickReportAttributes) {
                $itemsFound = $Global:ListBoxSelectAttributes.Items | Where-Object { $_.Content -eq $attr }
                if ($itemsFound) {
                    foreach ($singleItemInFoundList in $itemsFound) {
                        if ($singleItemInFoundList -is [System.Windows.Controls.ListBoxItem]) {
                            $singleItemInFoundList.IsSelected = $true
                        }
                    }
                }
            }

            $ReportData = Get-ADReportData -CustomFilter "LockedOut -eq `$true" -SelectedAttributes $QuickReportAttributes
            if ($null -eq $ReportData) { $ReportData = @() }
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
            $Global:DataGridResults.ItemsSource = $ReportData
            Update-ResultCounters -Results $ReportData
            $Global:TextBlockStatus.Text = "Locked out users loaded. $($ReportData.Count) record(s) found."
            Write-ADReportLog -Message "Locked out users loaded. $($ReportData.Count) result(s) found." -Type Info
        } catch {
            $ErrorMessage = "Error loading locked out users: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
            Update-ResultCounters -Results @()
            $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
        }
    })

    $ButtonQuickGroups.add_Click({
        Write-ADReportLog -Message "Loading all groups..." -Type Info
        try {
            if ($script:ListBoxSelectionChangedHandler) {
                $Global:ListBoxSelectAttributes.remove_SelectionChanged($script:ListBoxSelectionChangedHandler)
            }
            $Global:RadioButtonGroup.IsChecked = $true
            $Global:ComboBoxFilterAttribute.SelectedIndex = -1
            $Global:TextBoxFilterValue.Text = ""

            $QuickReportAttributes = @("Name", "SamAccountName", "GroupCategory", "GroupScope")
            $Global:ListBoxSelectAttributes.SelectedItems.Clear()
            foreach ($attr in $QuickReportAttributes) {
                $itemsFound = $Global:ListBoxSelectAttributes.Items | Where-Object { $_.Content -eq $attr }
                if ($itemsFound) {
                    foreach ($singleItemInFoundList in $itemsFound) {
                        if ($singleItemInFoundList -is [System.Windows.Controls.ListBoxItem]) {
                            $singleItemInFoundList.IsSelected = $true
                        }
                    }
                }
            }

            $ReportData = Get-ADGroupReportData -CustomFilter "*" -SelectedAttributes $QuickReportAttributes
            if ($null -eq $ReportData) { $ReportData = @() }
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
            $Global:DataGridResults.ItemsSource = $ReportData
            Update-ResultCounters -Results $ReportData
            $Global:TextBlockStatus.Text = "All groups loaded. $($ReportData.Count) record(s) found."
            Write-ADReportLog -Message "All groups loaded. $($ReportData.Count) result(s) found." -Type Info
        } catch {
            $ErrorMessage = "Error loading all groups: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
            Update-ResultCounters -Results @()
            $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
        }
    })

    $ButtonQuickSecurityGroups.add_Click({
        Write-ADReportLog -Message "Loading security groups..." -Type Info
        try {
            $QuickReportAttributes = @("Name", "SamAccountName", "GroupCategory", "GroupScope")
            $ReportData = Get-ADGroupReportData -CustomFilter "GroupCategory -eq 'Security'" -SelectedAttributes $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "Security groups loaded. $($ReportData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReportData
            } else {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
            }
        } catch {
            $ErrorMessage = "Error loading security groups: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
        }
    })

    $ButtonQuickDistributionGroups.add_Click({
        Write-ADReportLog -Message "Loading distribution lists..." -Type Info
        try {
            $QuickReportAttributes = @("Name", "SamAccountName", "GroupCategory", "GroupScope")
            $ReportData = Get-ADGroupReportData -CustomFilter "GroupCategory -eq 'Distribution'" -SelectedAttributes $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "Distribution lists loaded. $($ReportData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReportData
            } else {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
            }
        } catch {
            $ErrorMessage = "Error loading distribution lists: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
        }
    })

    # Neue Funktionen für Benutzer
    $ButtonQuickNeverExpire.add_Click({
        Write-ADReportLog -Message "Loading users with non-expiring passwords..." -Type Info
        try {
            $QuickReportAttributes = @("DisplayName", "SamAccountName", "Enabled", "PasswordNeverExpires", "LastLogonDate")
            $ReportData = Get-ADReportData -CustomFilter "PasswordNeverExpires -eq `$true" -SelectedAttributes $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "Users with non-expiring passwords loaded. $($ReportData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReportData
            } else {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
            }
        } catch {
            $ErrorMessage = "Error loading users with non-expiring passwords: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
        }
    })

    $ButtonQuickInactiveUsers.add_Click({
        Write-ADReportLog -Message "Loading inactive users (>90 days)..." -Type Info
        try {
            $QuickReportAttributes = @("DisplayName", "SamAccountName", "Enabled", "LastLogonDate")
            $InactiveCutoffDate = (Get-Date).AddDays(-90)
            $ReportData = Get-ADReportData -CustomFilter "(LastLogonTimeStamp -lt '$($InactiveCutoffDate.ToFileTime())') -or (LastLogonTimeStamp -notlike '*')" -SelectedAttributes $QuickReportAttributes
            if ($ReportData) {
                # Nachbearbeitung: Nur wirklich inaktive Benutzer behalten
                $FilteredData = $ReportData | Where-Object { 
                    !$_.LastLogonDate -or ($_.LastLogonDate -lt $InactiveCutoffDate)
                }
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                $Global:DataGridResults.ItemsSource = $FilteredData
                Write-ADReportLog -Message "Inactive users loaded. $($FilteredData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $FilteredData
            } else {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
            }
        } catch {
            $ErrorMessage = "Error loading inactive users: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
        }
    })

    $ButtonQuickAdminUsers.add_Click({
        Write-ADReportLog -Message "Loading administrators..." -Type Info
        try {
            # Verbesserte Methode zum Finden von Admin-Benutzern
            # Zuerst alle Benutzer laden
            $AllUsers = Get-ADUser -Filter * -Properties DisplayName, SamAccountName, Enabled, LastLogonDate
            
            # Bekannte Admin-Gruppenbezeichnungen (deutsch und englisch)
            $AdminGroups = @(
                "Domain Admins", "Domänen-Admins",
                "Enterprise Admins", "Organisations-Admins",
                "Administrators", "Administratoren", 
                "Schema Admins", "Schema-Admins",
                "AAD DC Administrators", "Azure AD-DC-Administratoren",
                "Server Operators", "Server-Operatoren",
                "Account Operators", "Konten-Operatoren",
                "Backup Operators", "Sicherungs-Operatoren"
            )
            
            $AdminUsers = @()
            # Prüfe Benutzer auf Admin-Rechte - erst SIDs der Admin-Gruppen ermitteln
            $AdminGroupSIDs = @()
            foreach ($groupName in $AdminGroups) {
                try {
                    $group = Get-ADGroup -Filter "Name -eq '$groupName'" -ErrorAction SilentlyContinue
                    if ($group) {
                        $AdminGroupSIDs += $group.SID.Value
                    }
                } catch {
                    # Ignoriere Fehler bei nicht existierenden Gruppen
                }
            }
            
            # Nur fortfahren, wenn wir Admin-Gruppen gefunden haben
            if ($AdminGroupSIDs.Count -gt 0) {
                # Für jeden Benutzer die Gruppenmitgliedschaften prüfen
                foreach ($user in $AllUsers) {
                    $memberOf = Get-ADPrincipalGroupMembership -Identity $user.SamAccountName -ErrorAction SilentlyContinue
                    foreach ($group in $memberOf) {
                        if ($AdminGroupSIDs -contains $group.SID.Value) {
                            # Benutzer ist Mitglied einer Admin-Gruppe
                            $adminUser = $user | Select-Object DisplayName, SamAccountName, Enabled, LastLogonDate, 
                                @{Name='AdminGroups'; Expression={
                                    ($memberOf | Where-Object { $AdminGroupSIDs -contains $_.SID.Value } | Select-Object -ExpandProperty Name) -join ", "
                                }}
                            $AdminUsers += $adminUser
                            break # Genug, wenn wir eine Admin-Gruppe gefunden haben
                        }
                    }
                }
            }
            
            if ($AdminUsers.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                $Global:DataGridResults.ItemsSource = $AdminUsers
                Write-ADReportLog -Message "Administrators loaded. $($AdminUsers.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $AdminUsers
            } else {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                Write-ADReportLog -Message "No administrator accounts found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading administrators: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
        }
    })

    # Neue Funktionen für Computer
    $ButtonQuickComputers.add_Click({
        Write-ADReportLog -Message "Loading all computers..." -Type Info
        try {
            $QuickReportAttributes = @("Name", "DNSHostName", "OperatingSystem", "Enabled", "LastLogonDate")
            $ReportData = Get-ADComputer -Filter * -Properties $QuickReportAttributes | Select-Object $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "All computers loaded. $($ReportData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReportData
            } else {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                Write-ADReportLog -Message "No computers found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading all computers: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
        }
    })

    $ButtonQuickInactiveComputers.add_Click({
        Write-ADReportLog -Message "Loading inactive computers (>90 days)..." -Type Info
        try {
            $QuickReportAttributes = @("Name", "DNSHostName", "OperatingSystem", "Enabled", "LastLogonDate", "PasswordLastSet")
            
            # Direktes Anfordern der spezifischen Computer-Eigenschaften, die wir brauchen
            $AllComputers = Get-ADComputer -Filter * -Properties Name, DNSHostName, OperatingSystem, Enabled, LastLogonDate, LastLogonTimeStamp, PasswordLastSet
            
            # Filtern nach inaktiven Computern basierend auf LastLogonDate und LastLogonTimeStamp
            $InactiveComputers = $AllComputers | Where-Object {
                ($_.LastLogonDate -lt ((Get-Date).AddDays(-90)) -or $_.LastLogonDate -eq $null) -and
                ($_.LastLogonTimeStamp -lt ((Get-Date).AddDays(-90)).ToFileTime() -or $_.LastLogonTimeStamp -eq $null)
            } | Select-Object Name, DNSHostName, OperatingSystem, Enabled, LastLogonDate, PasswordLastSet,
              @{Name='InactiveDays'; Expression={
                if ($_.LastLogonDate) {
                    [math]::Round((New-TimeSpan -Start $_.LastLogonDate -End (Get-Date)).TotalDays)
                } else { "N/A" }
              }}
            
            if ($InactiveComputers -and $InactiveComputers.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                $Global:DataGridResults.ItemsSource = $InactiveComputers
                Write-ADReportLog -Message "Inactive computers loaded. $($InactiveComputers.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $InactiveComputers
            } else {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                Write-ADReportLog -Message "No inactive computers found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading inactive computers: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
        }
    })

    # --- Security Audit Event Handler ---
    $ButtonQuickWeakPasswordPolicy.add_Click({
        Write-ADReportLog -Message "Loading users with weak password policies..." -Type Info
        try {
            $WeakPasswordUsers = Get-WeakPasswordPolicyUsers
            if ($WeakPasswordUsers -and $WeakPasswordUsers.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                $Global:DataGridResults.ItemsSource = $WeakPasswordUsers
                Write-ADReportLog -Message "Users with weak password policies loaded. $($WeakPasswordUsers.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $WeakPasswordUsers
            } else {
                $Global:DataGridResults.ItemsSource = $null
                $Global:DataGridResults.Columns.Clear()
                Write-ADReportLog -Message "No users with weak password policies found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading users with weak password policies: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:DataGridResults.Columns.Clear()
        }
    })

    $ButtonQuickRiskyGroupMemberships.add_Click({
        Write-ADReportLog -Message "Loading users with risky group memberships..." -Type Info
        try {
            $RiskyUsers = Get-RiskyGroupMemberships
            if ($RiskyUsers -and $RiskyUsers.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $RiskyUsers
                Write-ADReportLog -Message "Users with risky group memberships loaded. $($RiskyUsers.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $RiskyUsers
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No users with risky group memberships found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading users with risky group memberships: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickPrivilegedAccounts.add_Click({
        Write-ADReportLog -Message "Loading privileged accounts..." -Type Info
        try {
            $PrivilegedAccounts = Get-PrivilegedAccounts
            if ($PrivilegedAccounts -and $PrivilegedAccounts.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $PrivilegedAccounts
                Write-ADReportLog -Message "Privileged accounts loaded. $($PrivilegedAccounts.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $PrivilegedAccounts
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No privileged accounts found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading privileged accounts: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    # --- AD-Health Event Handler ---
    $ButtonQuickFSMORoles.add_Click({
        Write-ADReportLog -Message "Loading FSMO role holders..." -Type Info
        try {
            $FSMORoles = Get-FSMORoles
            if ($FSMORoles -and $FSMORoles.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $FSMORoles
                Write-ADReportLog -Message "FSMO role holders loaded. $($FSMORoles.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $FSMORoles
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No FSMO role information found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading FSMO role holders: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickDCStatus.add_Click({
        Write-ADReportLog -Message "Loading domain controller status..." -Type Info
        try {
            $DomainControllers = Get-DomainControllerStatus
            if ($DomainControllers -and $DomainControllers.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $DomainControllers
                Write-ADReportLog -Message "Domain controller status loaded. $($DomainControllers.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $DomainControllers
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No domain controllers found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading domain controller status: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickReplicationStatus.add_Click({
        Write-ADReportLog -Message "Loading AD replication status..." -Type Info
        try {
            $ReplicationStatus = Get-ReplicationStatus
            if ($ReplicationStatus -and $ReplicationStatus.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $ReplicationStatus
                Write-ADReportLog -Message "AD replication status loaded. $($ReplicationStatus.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReplicationStatus
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No replication status information found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading AD replication status: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonExportCSV.add_Click({
        Write-ADReportLog -Message "Preparing CSV export..." -Type Info
        if ($null -eq $Global:DataGridResults.ItemsSource -or $Global:DataGridResults.Items.Count -eq 0) {
            Write-ADReportLog -Message "No data available for export." -Type Warning
            [System.Windows.Forms.MessageBox]::Show("Es sind keine Daten zum Exportieren vorhanden.", "Hinweis", "OK", "Information")
            return
        }

        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "CSV (Comma delimited) (*.csv)|*.csv"
        $SaveFileDialog.Title = "Save CSV file as"
        $SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")

        if ($SaveFileDialog.ShowDialog() -eq "OK") {
            $FilePath = $SaveFileDialog.FileName
            try {
                $Global:DataGridResults.ItemsSource | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
                $Global:TextBlockStatus.Text = "Daten erfolgreich nach $FilePath exportiert."
                [System.Windows.Forms.MessageBox]::Show("Data successfully exported!", "CSV Export", "OK", "Information")
            } catch {
                $ErrorMessage = "Error in CSV export: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren der Daten: $($_.Exception.Message)", "Export Fehler", "OK", "Error")
            }
        } else {
            Write-ADReportLog -Message "CSV export canceled by user." -Type Info
        }
    })

    $ButtonExportHTML.add_Click({
        Write-ADReportLog -Message "Preparing HTML export..." -Type Info
        if ($null -eq $Global:DataGridResults.ItemsSource -or $Global:DataGridResults.Items.Count -eq 0) {
            $Global:TextBlockStatus.Text = "No data available for export."
            [System.Windows.Forms.MessageBox]::Show("Es sind keine Daten zum Exportieren vorhanden.", "Hinweis", "OK", "Information")
            return
        }

        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "HTML File (*.html;*.htm)|*.html;*.htm"
        $SaveFileDialog.Title = "Save HTML file as"
        $SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")

        if ($SaveFileDialog.ShowDialog() -eq "OK") {
            $FilePath = $SaveFileDialog.FileName
            try {
                # Create a more attractive HTML header
                $HtmlHead = @"
<meta http-equiv='Content-Type' content='text/html; charset=utf-8' />
<title>Active Directory Report</title>
<style>
  body { font-family: Segoe UI, Arial, sans-serif; margin: 20px; }
  table { border-collapse: collapse; width: 90%; margin: 20px auto; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
  th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
  th { background-color: #0078D4; color: white; }
  tr:nth-child(even) { background-color: #f2f2f2; }
  h1 { text-align: center; color: #333; }
</style>
"@
                $DateTimeNow = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
                $ReportTitle = "Active Directory Report - Created on $DateTimeNow"

                $Global:DataGridResults.ItemsSource | ConvertTo-Html -Head $HtmlHead -Body "<h1>$ReportTitle</h1>" | Out-File -FilePath $FilePath -Encoding UTF8
                $Global:TextBlockStatus.Text = "Daten erfolgreich nach $FilePath exportiert."
                [System.Windows.Forms.MessageBox]::Show("Data successfully exported!", "HTML Export", "OK", "Information")
            } catch {
                $ErrorMessage = "Error in HTML export: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren der Daten: $($_.Exception.Message)", "Export Fehler", "OK", "Error")
            }
        } else {
            Write-ADReportLog -Message "HTML export canceled by user." -Type Info
        }
    })

    # Fenster anzeigen
    $null = $Window.ShowDialog()

# Platzhalter-Funktion für initiale Visualisierung beim Start
Function Initialize-ResultVisualization {
    [CmdletBinding()]
    param()
    
    # Erstellt eine leere Visualisierung, bis Suchergebnisse vorliegen
    $emptyData = @{"Total" = 0}
    
    # Visualisierungen mit leeren Daten initialisieren
    New-MiniDonutChart -Canvas $Global:UserStatusCanvas -Data $emptyData
    New-MiniDonutChart -Canvas $Global:ComputerStatusCanvas -Data $emptyData
    New-MiniDonutChart -Canvas $Global:GroupsStatusCanvas -Data $emptyData
}

# --- Hauptlogik ---
Function Start-ADReportGUI {
    # Bereinige alle alten globalen UI-Variablen vor dem Start
    $UiVariables = @("Window", "ComboBoxFilterAttribute", "TextBoxFilterValue", "ListBoxSelectAttributes",
                  "ButtonQueryAD", "ButtonQuickAllUsers", "ButtonQuickDisabledUsers", "ButtonQuickLockedUsers",
                  "ButtonQuickGroups", "ButtonQuickSecurityGroups", "ButtonQuickDistributionGroups",
                  "ButtonQuickNeverExpire", "ButtonQuickInactiveUsers", "ButtonQuickAdminUsers",
                  "ButtonQuickComputers", "ButtonQuickInactiveComputers", 
                  "ButtonQuickWeakPasswordPolicy", "ButtonQuickRiskyGroupMemberships", "ButtonQuickPrivilegedAccounts",
                  "DataGridResults", "TextBlockStatus", "ButtonExportCSV", "ButtonExportHTML",
                  "ResultCountGrid", "UserCountText", "ComputerCountText", "GroupCountText")
    
    foreach ($var in $UiVariables) {
        Remove-Variable -Name $var -Scope Global -ErrorAction SilentlyContinue
    }

    # Ruft Initialize-ADReportForm auf, welche die UI lädt, Elemente zuweist und füllt.
    Initialize-ADReportForm -XAMLContent $Global:XAML
}

# --- Skriptstart ---
Start-ADReportGUI
# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCELl2uhBv72JEC
# SSpRaY2Ia+SnjgtsehqUV0ezT6vF9KCCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
# jEi+sBC2rBMTMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFVBoaW5JVC1QU3Nj
# cmlwdHNfU2lnbjAeFw0yNTA3MDUwODI4MTZaFw0yNzA3MDUwODM4MTZaMCAxHjAc
# BgNVBAMMFVBoaW5JVC1QU3NjcmlwdHNfU2lnbjCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBALmz3o//iDA5MvAndTjGX7/AvzTSACClfuUR9WYK0f6Ut2dI
# mPxn+Y9pZlLjXIpZT0H2Lvxq5aSI+aYeFtuJ8/0lULYNCVT31Bf+HxervRBKsUyi
# W9+4PH6STxo3Pl4l56UNQMcWLPNjDORWRPWHn0f99iNtjI+L4tUC/LoWSs3obzxN
# 3uTypzlaPBxis2qFSTR5SWqFdZdRkcuI5LNsJjyc/QWdTYRrfmVqp0QrvcxzCv8u
# EiVuni6jkXfiE6wz+oeI3L2iR+ywmU6CUX4tPWoS9VTtmm7AhEpasRTmrrnSg20Q
# jiBa1eH5TyLAH3TcYMxhfMbN9a2xDX5pzM65EJUCAwEAAaNGMEQwDgYDVR0PAQH/
# BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQO7XOqiE/EYi+n
# IaR6YO5M2MUuVTANBgkqhkiG9w0BAQsFAAOCAQEAjYOKIwBu1pfbdvEFFaR/uY88
# peKPk0NnvNEc3dpGdOv+Fsgbz27JPvItITFd6AKMoN1W48YjQLaU22M2jdhjGN5i
# FSobznP5KgQCDkRsuoDKiIOTiKAAknjhoBaCCEZGw8SZgKJtWzbST36Thsdd/won
# ihLsuoLxfcFnmBfrXh3rTIvTwvfujob68s0Sf5derHP/F+nphTymlg+y4VTEAijk
# g2dhy8RAsbS2JYZT7K5aEJpPXMiOLBqd7oTGfM7y5sLk2LIM4cT8hzgz3v5yPMkF
# H2MdR//K403e1EKH9MsGuGAJZddVN8ppaiESoPLoXrgnw2SY5KCmhYw1xRFdjTCC
# BY0wggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3
# LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290
# IENBMB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0
# LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkq
# hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE9
# 8orYWcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9S
# H8aeFaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g
# 1ckgHWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RY
# jgwrt0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgD
# EI3Y1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNA
# vwjXWkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDg
# ohIbZpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQA
# zH0clcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOk
# GLimdwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHF
# ynIWIgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gd
# LfXZqbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYE
# FOzX44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUH
# MAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDov
# L2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# dDBFBgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGln
# aUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkq
# hkiG9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7
# IviHGmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/5
# 9PesMHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0
# POz3A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISf
# b8rbII01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhU
# LSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBq4wggSWoAMCAQICEAc2
# N7ckVHzYR6z9KGYqXlswDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEh
# MB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTIyMDMyMzAwMDAw
# MFoXDTM3MDMyMjIzNTk1OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYg
# U0hBMjU2IFRpbWVTdGFtcGluZyBDQTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBAMaGNQZJs8E9cklRVcclA8TykTepl1Gh1tKD0Z5Mom2gsMyD+Vr2EaFE
# FUJfpIjzaPp985yJC3+dH54PMx9QEwsmc5Zt+FeoAn39Q7SE2hHxc7Gz7iuAhIoi
# GN/r2j3EF3+rGSs+QtxnjupRPfDWVtTnKC3r07G1decfBmWNlCnT2exp39mQh0YA
# e9tEQYncfGpXevA3eZ9drMvohGS0UvJ2R/dhgxndX7RUCyFobjchu0CsX7LeSn3O
# 9TkSZ+8OpWNs5KbFHc02DVzV5huowWR0QKfAcsW6Th+xtVhNef7Xj3OTrCw54qVI
# 1vCwMROpVymWJy71h6aPTnYVVSZwmCZ/oBpHIEPjQ2OAe3VuJyWQmDo4EbP29p7m
# O1vsgd4iFNmCKseSv6De4z6ic/rnH1pslPJSlRErWHRAKKtzQ87fSqEcazjFKfPK
# qpZzQmiftkaznTqj1QPgv/CiPMpC3BhIfxQ0z9JMq++bPf4OuGQq+nUoJEHtQr8F
# nGZJUlD0UfM2SU2LINIsVzV5K6jzRWC8I41Y99xh3pP+OcD5sjClTNfpmEpYPtMD
# iP6zj9NeS3YSUZPJjAw7W4oiqMEmCPkUEBIDfV8ju2TjY+Cm4T72wnSyPx4Jduyr
# XUZ14mCjWAkBKAAOhFTuzuldyF4wEr1GnrXTdrnSDmuZDNIztM2xAgMBAAGjggFd
# MIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBS6FtltTYUvcyl2mi91
# jGogj57IbzAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNVHQ8B
# Af8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBpMCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKG
# NWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290
# RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQC
# MAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAfVmOwJO2b5ipRCIBfmbW
# 2CFC4bAYLhBNE88wU86/GPvHUF3iSyn7cIoNqilp/GnBzx0H6T5gyNgL5Vxb122H
# +oQgJTQxZ822EpZvxFBMYh0MCIKoFr2pVs8Vc40BIiXOlWk/R3f7cnQU1/+rT4os
# equFzUNf7WC2qk+RZp4snuCKrOX9jLxkJodskr2dfNBwCnzvqLx1T7pa96kQsl3p
# /yhUifDVinF2ZdrM8HKjI/rAJ4JErpknG6skHibBt94q6/aesXmZgaNWhqsKRcnf
# xI2g55j7+6adcq/Ex8HBanHZxhOACcS2n82HhyS7T6NJuXdmkfFynOlLAlKnN36T
# U6w7HQhJD5TNOXrd/yVjmScsPT9rp/Fmw0HNT7ZAmyEhQNC3EyTN3B14OuSereU0
# cZLXJmvkOHOrpgFPvT87eK1MrfvElXvtCl8zOYdBeHo46Zzh3SP9HSjTx/no8Zhf
# +yvYfvJGnXUsHicsJttvFXseGYs2uJPU5vIXmVnKcPA3v5gA3yAWTyf7YGcWoWa6
# 3VXAOimGsJigK+2VQbc61RWYMbRiCQ8KvYHZE/6/pNHzV9m8BPqC3jLfBInwAM1d
# wvnQI38AC+R2AibZ8GV2QqYphwlHK+Z/GqSFD/yYlvZVVCsfgPrA8g4r5db7qS9E
# FUrnEw4d2zc4GqEr9u3WfPwwgga8MIIEpKADAgECAhALrma8Wrp/lYfG+ekE4zME
# MA0GCSqGSIb3DQEBCwUAMGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2Vy
# dCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNI
# QTI1NiBUaW1lU3RhbXBpbmcgQ0EwHhcNMjQwOTI2MDAwMDAwWhcNMzUxMTI1MjM1
# OTU5WjBCMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNlcnQxIDAeBgNVBAMT
# F0RpZ2lDZXJ0IFRpbWVzdGFtcCAyMDI0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8A
# MIICCgKCAgEAvmpzn/aVIauWMLpbbeZZo7Xo/ZEfGMSIO2qZ46XB/QowIEMSvgjE
# dEZ3v4vrrTHleW1JWGErrjOL0J4L0HqVR1czSzvUQ5xF7z4IQmn7dHY7yijvoQ7u
# jm0u6yXF2v1CrzZopykD07/9fpAT4BxpT9vJoJqAsP8YuhRvflJ9YeHjes4fduks
# THulntq9WelRWY++TFPxzZrbILRYynyEy7rS1lHQKFpXvo2GePfsMRhNf1F41nyE
# g5h7iOXv+vjX0K8RhUisfqw3TTLHj1uhS66YX2LZPxS4oaf33rp9HlfqSBePejlY
# eEdU740GKQM7SaVSH3TbBL8R6HwX9QVpGnXPlKdE4fBIn5BBFnV+KwPxRNUNK6lY
# k2y1WSKour4hJN0SMkoaNV8hyyADiX1xuTxKaXN12HgR+8WulU2d6zhzXomJ2Ple
# I9V2yfmfXSPGYanGgxzqI+ShoOGLomMd3mJt92nm7Mheng/TBeSA2z4I78JpwGpT
# RHiT7yHqBiV2ngUIyCtd0pZ8zg3S7bk4QC4RrcnKJ3FbjyPAGogmoiZ33c1HG93V
# p6lJ415ERcC7bFQMRbxqrMVANiav1k425zYyFMyLNyE1QulQSgDpW9rtvVcIH7Wv
# G9sqYup9j8z9J1XqbBZPJ5XLln8mS8wWmdDLnBHXgYly/p1DhoQo5fkCAwEAAaOC
# AYswggGHMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQM
# MAoGCCsGAQUFBwMIMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATAf
# BgNVHSMEGDAWgBS6FtltTYUvcyl2mi91jGogj57IbzAdBgNVHQ4EFgQUn1csA3cO
# KBWQZqVjXu5Pkh92oFswWgYDVR0fBFMwUTBPoE2gS4ZJaHR0cDovL2NybDMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVTdGFt
# cGluZ0NBLmNybDCBkAYIKwYBBQUHAQEEgYMwgYAwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBYBggrBgEFBQcwAoZMaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVT
# dGFtcGluZ0NBLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAPa0eH3aZW+M4hBJH2UOR
# 9hHbm04IHdEoT8/T3HuBSyZeq3jSi5GXeWP7xCKhVireKCnCs+8GZl2uVYFvQe+p
# PTScVJeCZSsMo1JCoZN2mMew/L4tpqVNbSpWO9QGFwfMEy60HofN6V51sMLMXNTL
# fhVqs+e8haupWiArSozyAmGH/6oMQAh078qRh6wvJNU6gnh5OruCP1QUAvVSu4kq
# VOcJVozZR5RRb/zPd++PGE3qF1P3xWvYViUJLsxtvge/mzA75oBfFZSbdakHJe2B
# VDGIGVNVjOp8sNt70+kEoMF+T6tptMUNlehSR7vM+C13v9+9ZOUKzfRUAYSyyEmY
# tsnpltD/GWX8eM70ls1V6QG/ZOB6b6Yum1HvIiulqJ1Elesj5TMHq8CWT/xrW7tw
# ipXTJ5/i5pkU5E16RSBAdOp12aw8IQhhA/vEbFkEiF2abhuFixUDobZaA0VhqAsM
# HOmaT3XThZDNi5U2zHKhUs5uHHdG6BoQau75KiNbh0c+hatSF+02kULkftARjsyE
# pHKsF7u5zKRbt5oK5YGwFvgc4pEVUNytmB3BpIiowOIIuDgP5M9WArHYSAR16gc0
# dP2XdkMEP5eBsX7bf/MGN4K3HP50v/01ZHo/Z5lGLvNwQ7XHBx1yomzLP8lx4Q1z
# ZKDyHcp4VQJLu2kWTsKsOqQxggUKMIIFBgIBATA0MCAxHjAcBgNVBAMMFVBoaW5J
# VC1QU3NjcmlwdHNfU2lnbgIQd487Ml/QoIxIvrAQtqwTEzANBglghkgBZQMEAgEF
# AKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgor
# BgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3
# DQEJBDEiBCCty6vpBhm6sKcZ6L3Jbq3ENvnvCFMgajQqOT5QjSN8OjANBgkqhkiG
# 9w0BAQEFAASCAQAmqirJxdUq4/tu5Z62dB2e7x9jVEuIm+8TOhe5rCK5qIrdyGFT
# JSrA9dw/iW+H2FJFiJc2aJYKDHfpwPwwHsM60YciepCjb7mapi9s/b6CyC+eoleF
# x0KR5VkMP+GWrVwFgKRuNxEsiHyxB+y+uShJuSxNfgR9pCb6W7berv4uW7kKUSvG
# 18GS0OO83mzinRroWh0b+5t9lMtMO7gXsE7ue/sGlsXTgVT6vN/HsSkPAtJ10WgZ
# 1z8cviiDqrfc1KDK7+U4luLYHWha4BKZwGT9412QmwVXvPuK3xl3reAMDhv+1seG
# muz2mPVI60WKcg9tPzV1zaNlZx27eTl/rzqcoYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTE0NDMwMVowLwYJKoZIhvcNAQkEMSIEIFyDJOQboQM8nmLhMb99KAj8J2QV
# xmNAvVw8+t8BeJ2WMA0GCSqGSIb3DQEBAQUABIICAG6hpF/rjsrSi28BMlpcafBV
# brykDF+kAqGaufj641g4rHWQiUpWrTP2YQKSRzm92jmTm8RRhOjhFNU/a+rnW2PG
# yyXGEdSkVCAgzXDCeek5LH7ujpaxTy2RUgUvyW+ySi2FSQtjVbEHiatzHBdusH/6
# zFS0qY7GnLmuSTQ+4Y3ukq3hMPKODbKVkl3iWmBerk/G9kgrpEiAPJgdfo53iIwJ
# /PM1az17UUGRmN/3fW0P3XW+fKqzpuZbl/JGhprwFAJmfZo+/a5IXl0Qhuw6gHMj
# OTMGylzxeECEDz/EeJyQHZIKS0pjZ0bM0uoPuK+B2l5UKbDNfph4UhyPXmEsfLMO
# AlzHfqCQqrYE7WTaLnFuuLzZ/ryN4B8FEZzdXF+Dy5m3mLsgI7kShDP+YaxHRLy6
# +f/XGJ19/6q/9uh71eBFeSA5uPSMjaij5te3jcW789/y2/rCjdhmsSUDWXxGIDn8
# AhkddBhy6lSCUws3GVeYdMaqNmUT1KfcCmmx7hUiE2MipPUt2zITJRQu7ala3Xjb
# r6xOjVFoSJNkOK7OzTtJ4xe4izuKrTZ6VmsQBXk9ZWLOjncuxcXBRBxO5psOr374
# FCmS28LILja6plzsPetbPKYI8gqttc0ffgQGheGLmLsa32ysusOpaxyyP9tWgRlq
# NL5h5zVlj3e09lHnw5vC
# SIG # End signature block

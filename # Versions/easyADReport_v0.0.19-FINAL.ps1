[xml]$Global:XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="easyADReport v0.0.17" Height="850" Width="1200"
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
                    <TextBlock Text="v0.0.19" Foreground="#E1E1E1" FontSize="14" VerticalAlignment="Center" Margin="10,0,0,0"/>
                </StackPanel>
                
                <!-- Ergebnisanzahl-Anzeige -->
                <Border Grid.Column="1" Background="#0063B1" Width="300" Height="50" CornerRadius="4" Margin="0,5,20,5">
                    <Grid x:Name="ResultCountGrid" Margin="5">
                        <!-- Gesamtergebnis-Anzeige -->
                        <Border Background="#0063B1" CornerRadius="2" Margin="2">
                            <Grid>
                                <TextBlock Text="Ergebnisse gesamt" Foreground="White" FontSize="14" HorizontalAlignment="Center" VerticalAlignment="Top" Margin="0,5,0,0"/>
                                <TextBlock x:Name="TotalResultCountText" Text="0" Foreground="White" FontSize="20" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="0,10,0,0"/>
                            </Grid>
                        </Border>
                    </Grid>
                </Border>
            </Grid>
        </Border>

        <!-- Upper Section: Filter and Attribute Selection -->
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="20,15">
            <Border Background="White" CornerRadius="4" BorderBrush="#DDDDDD" BorderThickness="1" Margin="0,0,15,0">
                <GroupBox Header="Filter" Margin="5" BorderThickness="0">
                    <StackPanel Orientation="Horizontal">
                        <Label Content="Filter Attribute:" VerticalAlignment="Center"/>
                        <ComboBox x:Name="ComboBoxFilterAttribute" Width="150" Margin="5,0" VerticalAlignment="Center" BorderThickness="1" BorderBrush="#CCCCCC"/>
                        <Label Content="Filter Value:" VerticalAlignment="Center"/>
                        <TextBox x:Name="TextBoxFilterValue" Width="200" Margin="5,0" VerticalAlignment="Center" BorderThickness="1" BorderBrush="#CCCCCC"/>
                    </StackPanel>
                </GroupBox>
            </Border>
            <Border Background="White" CornerRadius="4" BorderBrush="#DDDDDD" BorderThickness="1" Margin="0,0,15,0">
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
            <Button x:Name="ButtonQueryAD" Content="Query" Width="100" Height="36" Margin="0,5" VerticalAlignment="Center" 
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

            <Border Background="White" CornerRadius="4" BorderBrush="#DDDDDD" BorderThickness="1" Margin="0,0,15,0">
                <GroupBox Header="Quick Reports" BorderThickness="0" Margin="5,0,0,0">
                    <StackPanel>
                        <GroupBox Header="Users" BorderThickness="0" Margin="5,10,0,0">
                            <StackPanel>
                                <Button x:Name="ButtonQuickAllUsers" Content="All Users" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickDisabledUsers" Content="Disabled Users" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickLockedUsers" Content="Locked Users" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickNeverExpire" Content="Password Never Expires" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickInactiveUsers" Content="Inactive Users (90+ days)" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickAdminUsers" Content="Administrators" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                            </StackPanel>
                        </GroupBox>
                        <GroupBox Header="Groups" BorderThickness="0" Margin="5,15,0,0">
                            <StackPanel>
                                <Button x:Name="ButtonQuickGroups" Content="All Groups" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickSecurityGroups" Content="Security Groups" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickDistributionGroups" Content="Distribution Lists" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                            </StackPanel>
                        </GroupBox>
                        <GroupBox Header="Computers" BorderThickness="0" Margin="5,15,0,0">
                            <StackPanel>
                                <Button x:Name="ButtonQuickComputers" Content="All Computers" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
                                <Button x:Name="ButtonQuickInactiveComputers" Content="Inactive Computers (90+ days)" Margin="0,2" Height="28" Background="#F5F5F5" BorderBrush="#DDDDDD" HorizontalAlignment="Stretch" HorizontalContentAlignment="Left"/>
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
                <TextBlock x:Name="TextBlockStatus" Text="Ready" VerticalAlignment="Center"/>
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
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

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
        [string]$CustomFilter
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

        # Sicherstellen, dass immer einige Basiseigenschaften geladen werden
        if ($FilterAttribute -and (-not ($PropertiesToLoad -contains $FilterAttribute))) {
            $PropertiesToLoad += $FilterAttribute
        }
        if ($PropertiesToLoad.Count -eq 0) {
            # Default properties if nothing was selected
            $PropertiesToLoad = @("DisplayName", "SamAccountName") 
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
            $Output = $Users | Select-Object $SelectAttributes
            return $Output
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

        if ($PropertiesToLoad.Count -eq 0) {
            $PropertiesToLoad = @("Name", "SamAccountName", "GroupCategory", "GroupScope") 
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
    $DefaultAttributes = @("DisplayName", "SamAccountName", "GivenName", "Surname", "mail", "Department", "Title", "Enabled", "LastLogonDate", "whenCreated", "MemberOf")
    $DefaultAttributes | ForEach-Object { $Global:ListBoxSelectAttributes.Items.Add($_) }

    # Standardattribute für die Filter-ComboBox füllen
    $FilterableAttributes = @("DisplayName", "SamAccountName", "GivenName", "Surname", "mail", "Department", "Title")
    $FilterableAttributes | ForEach-Object { $Global:ComboBoxFilterAttribute.Items.Add($_) }
    if ($Global:ComboBoxFilterAttribute.Items.Count -gt 0) { $Global:ComboBoxFilterAttribute.SelectedIndex = 0 }

    # Funktion zur Aktualisierung der Ergebniszähler im Header
Function Update-ResultCounters {
    param (
        [Parameter(Mandatory=$true)]
        $Results
    )
    
    Write-ADReportLog -Message "Aktualisiere Ergebnisanzeige..." -Type Info -Terminal
    
    try {
        # Ergebnisanzahl ermitteln
        $totalCount = 0
        
        if ($null -ne $Results) {
            $totalCount = $Results.Count
            
            # Objekttyp für das Log bestimmen (nur für die Protokollierung)
            if ($totalCount -gt 0) {
                $firstObject = $Results[0]
                $objectType = "Unbekannt"
                
                if ($firstObject.PSObject.Properties.Name -contains "ObjectClass") {
                    switch ($firstObject.ObjectClass) {
                        "user" { $objectType = "Benutzer" }
                        "computer" { $objectType = "Computer" }
                        "group" { $objectType = "Gruppen" }
                    }
                }
                
                Write-ADReportLog -Message "$objectType-Ergebnisse: $totalCount" -Type Info -Terminal
            } else {
                Write-ADReportLog -Message "Keine Ergebnisse gefunden" -Type Info -Terminal
            }
        } else {
            Write-ADReportLog -Message "Keine Ergebnisse (Null)" -Type Info -Terminal
        }
        
        # Überprüfen, ob die UI-Elemente existieren, bevor wir versuchen, sie zu aktualisieren
        if ($null -ne $Global:TotalResultCountText) {
            $Global:TotalResultCountText.Text = $totalCount.ToString()
        }
        
        # Aktualisiere den Status im Status-Textblock
        if ($null -ne $Global:TextBlockStatus) {
            $Global:TextBlockStatus.Text = "Abfrage abgeschlossen. $totalCount Ergebnis(se) gefunden."
        }
    }
    catch {
        Write-ADReportLog -Message "Fehler beim Aktualisieren der Ergebnisanzeige: $($_.Exception.Message)" -Type Error
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
        $Global:TextBlockStatus.Text = "Bereit für Abfrage..."
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
        [string]$StatusMessage = ""
    )
    
    # DataGrid aktualisieren
    if ($null -ne $Global:DataGridResults) {
        $Global:DataGridResults.ItemsSource = $Results
    }
    
    # Ergebniszähler aktualisieren
    Update-ResultCounters -Results $Results
    
    # Statusmeldung anzeigen, wenn vorhanden
    if (-not [string]::IsNullOrWhiteSpace($StatusMessage) -and $null -ne $Global:TextBlockStatus) {
        $Global:TextBlockStatus.Text = $StatusMessage
    }
}

    # --- Event Handler zuweisen --- 
    $ButtonQueryAD.add_Click({
        Write-ADReportLog -Message "Executing query..." -Type Info
        try {
            $SelectedFilterAttribute = $Global:ComboBoxFilterAttribute.SelectedItem.ToString()
            $FilterValue = $Global:TextBoxFilterValue.Text
            $SelectedExportAttributes = $Global:ListBoxSelectAttributes.SelectedItems

            if ($SelectedExportAttributes.Count -eq 0) {
                Write-ADReportLog -Message "Please select at least one attribute for export." -Type Warning
                return
            }

            # AD-Abfrage durchführen
            $ReportData = Get-ADReportData -FilterAttribute $SelectedFilterAttribute -FilterValue $FilterValue -SelectedAttributes $SelectedExportAttributes
            
            if ($ReportData) {
                try {
                    # Debug-Informationen
                    Write-ADReportLog -Message "ReportData Typ: $($ReportData.GetType().FullName)" -Type Info -Terminal
                    
                    # Wir brauchen sicherzustellen, dass wir immer eine Liste/Sammlung haben, auch bei einzelnen Objekten
                    # Verwende @() um es als Array zu erzwingen
                    $ReportCollection = @($ReportData)
                    
                    # Direkte Zuweisung an DataGrid
                    $Global:DataGridResults.ItemsSource = $ReportCollection
                    
                    # Zähle die Anzahl der Ergebnisse
                    $Count = $ReportCollection.Count
                    Write-ADReportLog -Message "Query completed. $Count result(s) found." -Type Info
                    
                    # Ergebniszähler im Header aktualisieren
                    Update-ResultCounters -Results $ReportCollection
                } catch {
                    $ErrorMessage = "Error processing query result: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    $Global:DataGridResults.ItemsSource = $null
                }
            } else {
                $Global:DataGridResults.ItemsSource = $null # DataGrid leeren
                # Status wird in Get-ADReportData gesetzt oder bleibt auf "Keine Benutzer gefunden"
            }
        } catch {
            $ErrorMessage = "Error in query process: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickAllUsers.add_Click({
        Write-ADReportLog -Message "Loading all users..." -Type Info
        try {
            $QuickReportAttributes = @("DisplayName", "SamAccountName", "Enabled", "LastLogonDate")
            $ReportData = Get-ADReportData -CustomFilter "*" -SelectedAttributes $QuickReportAttributes
            
            # Verwende die neue universelle Funktion zur Aktualisierung der Ergebnisse
            if ($ReportData) {
                $StatusMessage = "All users loaded. $($ReportData.Count) result(s) found."
            } else {
                $StatusMessage = "Keine Benutzer gefunden."
                $ReportData = @() # Leeres Array bei keinen Ergebnissen
            }
            
            Write-ADReportLog -Message $StatusMessage -Type Info
            Update-ADReportResults -Results $ReportData -StatusMessage $StatusMessage
        } catch {
            $ErrorMessage = "Error loading all users: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            Update-ADReportResults -Results @() -StatusMessage $ErrorMessage
        }
    })

    $ButtonQuickDisabledUsers.add_Click({
        Write-ADReportLog -Message "Loading disabled users..." -Type Info
        try {
            $QuickReportAttributes = @("DisplayName", "SamAccountName", "Enabled", "LastLogonDate")
            # Der Filter für deaktivierte Benutzer ist 'Enabled -eq $false'
            $ReportData = Get-ADReportData -CustomFilter "Enabled -eq `$false" -SelectedAttributes $QuickReportAttributes
            
            # Verwende die neue universelle Funktion zur Aktualisierung der Ergebnisse
            if ($ReportData) {
                $StatusMessage = "Disabled users loaded. $($ReportData.Count) result(s) found."
            } else {
                $StatusMessage = "Keine deaktivierten Benutzer gefunden."
                $ReportData = @() # Leeres Array bei keinen Ergebnissen
            }
            
            Write-ADReportLog -Message $StatusMessage -Type Info
            Update-ADReportResults -Results $ReportData -StatusMessage $StatusMessage
        } catch {
            $ErrorMessage = "Error loading disabled users: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            Update-ADReportResults -Results @() -StatusMessage $ErrorMessage
        }
    })

    $ButtonQuickLockedUsers.add_Click({
        Write-ADReportLog -Message "Loading locked users..." -Type Info
        try {
            # Verwende Search-ADAccount, da es zuverlässiger für gesperrte Konten ist
            Write-ADReportLog -Message "Executing Search-ADAccount -LockedOut -UsersOnly to find locked users." -Type Info -Terminal
            $LockedUsers = Search-ADAccount -LockedOut -UsersOnly
            
            if ($LockedUsers -and $LockedUsers.Count -gt 0) {
                # Hole zusätzliche Eigenschaften für die gefundenen Benutzer
                $LockedUsersData = $LockedUsers | ForEach-Object {
                    Get-ADUser -Identity $_.SamAccountName -Properties DisplayName, SamAccountName, Enabled, LockedOut, LastLogonDate
                } | Select-Object DisplayName, SamAccountName, Enabled, LockedOut, LastLogonDate
                
                $Global:DataGridResults.ItemsSource = $LockedUsersData
                Write-ADReportLog -Message "Locked users loaded. $($LockedUsersData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $LockedUsersData
            } else {
                # Versuche es mit der Filter-Methode als Fallback
                $QuickReportAttributes = @("DisplayName", "SamAccountName", "Enabled", "LockedOut", "LastLogonDate")
                $ReportData = Get-ADReportData -CustomFilter "LockedOut -eq `$true" -SelectedAttributes $QuickReportAttributes
                
                if ($ReportData -and $ReportData.Count -gt 0) {
                    $Global:DataGridResults.ItemsSource = $ReportData
                    Write-ADReportLog -Message "Locked users loaded. $($ReportData.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $ReportData
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No locked user accounts found." -Type Warning
                }
            }
        } catch {
            $ErrorMessage = "Error loading locked users: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickGroups.add_Click({
        Write-ADReportLog -Message "Loading all groups..." -Type Info
        try {
            $QuickReportAttributes = @("Name", "SamAccountName", "GroupCategory", "GroupScope")
            $ReportData = Get-ADGroupReportData -CustomFilter "*" -SelectedAttributes $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "All groups loaded. $($ReportData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReportData
            } else {
                $Global:DataGridResults.ItemsSource = $null
                # Status wird in Get-ADGroupReportData gesetzt
            }
        } catch {
            $ErrorMessage = "Error loading all groups: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickSecurityGroups.add_Click({
        Write-ADReportLog -Message "Loading security groups..." -Type Info
        try {
            $QuickReportAttributes = @("Name", "SamAccountName", "GroupCategory", "GroupScope")
            $ReportData = Get-ADGroupReportData -CustomFilter "GroupCategory -eq 'Security'" -SelectedAttributes $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "Security groups loaded. $($ReportData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReportData
            } else {
                $Global:DataGridResults.ItemsSource = $null
            }
        } catch {
            $ErrorMessage = "Error loading security groups: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickDistributionGroups.add_Click({
        Write-ADReportLog -Message "Loading distribution lists..." -Type Info
        try {
            $QuickReportAttributes = @("Name", "SamAccountName", "GroupCategory", "GroupScope")
            $ReportData = Get-ADGroupReportData -CustomFilter "GroupCategory -eq 'Distribution'" -SelectedAttributes $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "Distribution lists loaded. $($ReportData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReportData
            } else {
                $Global:DataGridResults.ItemsSource = $null
            }
        } catch {
            $ErrorMessage = "Error loading distribution lists: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    # Neue Funktionen für Benutzer
    $ButtonQuickNeverExpire.add_Click({
        Write-ADReportLog -Message "Loading users with non-expiring passwords..." -Type Info
        try {
            $QuickReportAttributes = @("DisplayName", "SamAccountName", "Enabled", "PasswordNeverExpires", "LastLogonDate")
            $ReportData = Get-ADReportData -CustomFilter "PasswordNeverExpires -eq `$true" -SelectedAttributes $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "Users with non-expiring passwords loaded. $($ReportData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReportData
            } else {
                $Global:DataGridResults.ItemsSource = $null
            }
        } catch {
            $ErrorMessage = "Error loading users with non-expiring passwords: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
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
                $Global:DataGridResults.ItemsSource = $FilteredData
                Write-ADReportLog -Message "Inactive users loaded. $($FilteredData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $FilteredData
            } else {
                $Global:DataGridResults.ItemsSource = $null
            }
        } catch {
            $ErrorMessage = "Error loading inactive users: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
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
                $Global:DataGridResults.ItemsSource = $AdminUsers
                Write-ADReportLog -Message "Administrators loaded. $($AdminUsers.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $AdminUsers
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No administrator accounts found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading administrators: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    # Neue Funktionen für Computer
    $ButtonQuickComputers.add_Click({
        Write-ADReportLog -Message "Loading all computers..." -Type Info
        try {
            $QuickReportAttributes = @("Name", "DNSHostName", "OperatingSystem", "Enabled", "LastLogonDate")
            $ReportData = Get-ADComputer -Filter * -Properties $QuickReportAttributes | Select-Object $QuickReportAttributes
            if ($ReportData) {
                $Global:DataGridResults.ItemsSource = $ReportData
                Write-ADReportLog -Message "All computers loaded. $($ReportData.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReportData
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No computers found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading all computers: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
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
                $Global:DataGridResults.ItemsSource = $InactiveComputers
                Write-ADReportLog -Message "Inactive computers loaded. $($InactiveComputers.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $InactiveComputers
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No inactive computers found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading inactive computers: $($_.Exception.Message)"
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
        
        # Visualisierungen erstellen
        New-MiniDonutChart -Canvas $Global:UserStatusCanvas -Data $userData -Colors $userColors
        New-MiniDonutChart -Canvas $Global:ComputerStatusCanvas -Data $computerData -Colors $computerColors
        New-MiniDonutChart -Canvas $Global:GroupsStatusCanvas -Data $groupData -Colors $groupColors
        
        Write-ADReportLog -Message "Result visualization updated successfully." -Type Info -Terminal
    }
    catch {
        Write-ADReportLog -Message "Error updating result visualization: $($_.Exception.Message)" -Type Error
    }
}

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
                  "ButtonQuickComputers", "ButtonQuickInactiveComputers", "DataGridResults",
                  "TextBlockStatus", "ButtonExportCSV", "ButtonExportHTML",
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDUbBRUqFzIp24C
# F9Tx7guadLiuGRoVFTbG+uTiDWfrc6CCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# DQEJBDEiBCAPnV2qMTY8hbEIAQGiMToEfBt5AGWyNqsYt3lH442u5zANBgkqhkiG
# 9w0BAQEFAASCAQC3NWXqFItCJKS0/7n/3H8UAsP6DB072bF4b1epb+W2DmIM4dUU
# k+0rp1pN6IVFh44TjUYGqGJSigCkfU3dn9h/KOlQtSWPUWQVQ8GSrNtYWSZ7ysGt
# dHI1i6MJHu3OgaetSbOKfOitKYwr63VxenqDfOlogs6CiDDNU+jW2cqv/KxOBAFw
# gu+RHJw0KEtp4ZO63X+Z3Ox3H6WXM5VpRLuKlIFuN3v90Z05asfx5bcmYoXjkHLx
# 0HErelDRB03Pulq5/bt63TnIHuvJsRbihy64b2Cwz7RRWacveN6QZzgGUYAC0Hxo
# TmvHHVQDowOMAiMdNaiyq5gZqhV3i1r3O18SoYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTE0NDMwMFowLwYJKoZIhvcNAQkEMSIEIFxfjBjg9Pckf/WxzEL7rEUe0JAl
# X9ktIb5BjOCzjPvvMA0GCSqGSIb3DQEBAQUABIICAJiVo18RIsIUkcU1cDotTHM3
# oxoznlH9dK1QsgwTbnXPkX6decv2mV8W9hUlF+efrj8JbUGwKZfSGh+BEpgzYu8j
# Db6cDRAbS9mdcqc1R4w6YIUCw7G3b51Fz4JkaCtcWIavTakajWhWv2i1x9OS0hcf
# WVURvPHnJuIPJu7QTwRrJ9SjjNbjOW0nW8skJ421sOj4AWpSOwVee7HmY6q9AehE
# gfyyHkBf9co7NEZFSZ55aZkGkgkpnCcj7xyah2JR97zDqQzo1DFUcDa9DOUjyOnn
# /RKTCFWJ06ycYL6Ic/hcePUQhbNshHuvAzwdrZ7ymDyJI0/fdMWi90MEa7FZqRYk
# jVvROH3nOX20paoq84CUJyUHj0X2GV2HVR6plQEKEVYVIqPy9fPB4B6ugasLELtr
# /8TRYVhtg5AfynGL99pidxLC/Q4aOoYjyfSAKTJm6bg/OO2qcY1CqIRuUse4ERk2
# oUX+YkOIOqD5Tc2321Kcau1QJkMUpNhlUgYfuO+pfzBBxkk6zrhCInp8hQPc2mqW
# xuL2iVzUOa1M37Oz4Ly5QVGAIEPHrI2rj52wH1TUqDDHcsdo06sPUx5iYFMfJ+zc
# 22SdsPQkypYhz13dy0n+4IfMoa66RkTLgK5rNaFrnuI3joS+HXXnitdT3EfdWoxz
# s+k29eBwxJx0kZADh8/Y
# SIG # End signature block

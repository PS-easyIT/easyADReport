﻿[xml]$Global:XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    Title="easyADReport v0.3.0"
    Width="1720"
    Height="1000"
    Background="#F8F9FA"
    FontFamily="Segoe UI"
    ResizeMode="CanResizeWithGrip"
    WindowStartupLocation="CenterScreen">

    <!-- Window Resources for Modern Styling -->
    <Window.Resources>
        <!-- Modern Card Style -->
        <Style x:Key="ModernCard" TargetType="Border">
            <Setter Property="Background" Value="White"/>
            <Setter Property="CornerRadius" Value="12"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>

        <!-- Modern Button Style -->
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#E5E7EB"/>
            <Setter Property="Padding" Value="16,10"/>
            <Setter Property="FontWeight" Value="Medium"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" 
                                Background="{TemplateBinding Background}" 
                                BorderBrush="{TemplateBinding BorderBrush}" 
                                BorderThickness="{TemplateBinding BorderThickness}" 
                                CornerRadius="8">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"
                                              Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#F9FAFB"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="#3B82F6"/>
                                <Setter TargetName="border" Property="Effect">
                                    <Setter.Value>
                                        <DropShadowEffect BlurRadius="6" ShadowDepth="1" Color="#103B82F6" Opacity="0.1"/>
                                    </Setter.Value>
                                </Setter>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#F3F4F6"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="#1D4ED8"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Primary Button Style -->
        <Style x:Key="PrimaryButton" TargetType="Button">
            <Setter Property="Background" Value="#3B82F6"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="24,14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" 
                                Background="{TemplateBinding Background}" 
                                CornerRadius="8">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"
                                              Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#2563EB"/>
                                <Setter TargetName="border" Property="Effect">
                                    <Setter.Value>
                                        <DropShadowEffect BlurRadius="8" ShadowDepth="2" Color="#103B82F6" Opacity="0.3"/>
                                    </Setter.Value>
                                </Setter>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#1D4ED8"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Category Header Style -->
        <Style x:Key="CategoryHeader" TargetType="TextBlock">
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="#2D3748"/>
            <Setter Property="Margin" Value="0,8,0,4"/>
        </Style>

        <!-- Sidebar Menu Button Style -->
        <Style x:Key="SidebarButton" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="FontWeight" Value="Normal"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="HorizontalAlignment" Value="Stretch"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="Margin" Value="0,1"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" 
                                Background="{TemplateBinding Background}" 
                                CornerRadius="6"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" 
                                              VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#F1F5F9"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#E2E8F0"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Active Sidebar Button Style -->
        <Style x:Key="SidebarButtonActive" TargetType="Button" BasedOn="{StaticResource SidebarButton}">
            <Setter Property="Background" Value="#EBF4FF"/>
            <Setter Property="Foreground" Value="#1E40AF"/>
            <Setter Property="FontWeight" Value="Medium"/>
        </Style>

        <!-- Expandable Section Style -->
        <Style x:Key="ExpanderStyle" TargetType="Expander">
            <Setter Property="IsExpanded" Value="True"/>
            <Setter Property="Margin" Value="0,4,0,8"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Expander">
                        <Border Background="#FAFBFC" CornerRadius="8" Margin="0,2">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <ToggleButton x:Name="HeaderSite" 
                                              Grid.Row="0" 
                                              IsChecked="{Binding IsExpanded, RelativeSource={RelativeSource TemplatedParent}}"
                                              Background="Transparent"
                                              BorderThickness="0"
                                              Padding="12,10"
                                              HorizontalAlignment="Stretch"
                                              HorizontalContentAlignment="Left">
                                    <StackPanel Orientation="Horizontal">
                                        <Path x:Name="arrow" 
                                              Fill="#6B7280" 
                                              Margin="0,0,8,0"
                                              VerticalAlignment="Center"
                                              Data="M4,6 L8,10 L4,14" 
                                              Stroke="#6B7280" 
                                              StrokeThickness="1.2"
                                              Width="12"
                                              Height="12">
                                            <Path.RenderTransform>
                                                <RotateTransform x:Name="arrowTransform" Angle="0" CenterX="6" CenterY="10"/>
                                            </Path.RenderTransform>
                                        </Path>
                                        <ContentPresenter Content="{TemplateBinding Header}" 
                                                          TextBlock.FontWeight="SemiBold" 
                                                          TextBlock.FontSize="13"
                                                          TextBlock.Foreground="#374151"/>
                                    </StackPanel>
                                </ToggleButton>
                                <ContentPresenter x:Name="ExpandSite" 
                                                  Grid.Row="1" 
                                                  Visibility="Collapsed" 
                                                  Margin="8,0,8,8"/>
                            </Grid>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsExpanded" Value="True">
                                <Setter TargetName="ExpandSite" Property="Visibility" Value="Visible"/>
                                <Trigger.EnterActions>
                                    <BeginStoryboard>
                                        <Storyboard>
                                            <DoubleAnimation Storyboard.TargetName="arrowTransform" 
                                                             Storyboard.TargetProperty="Angle" 
                                                             To="90" Duration="0:0:0.2"/>
                                        </Storyboard>
                                    </BeginStoryboard>
                                </Trigger.EnterActions>
                                <Trigger.ExitActions>
                                    <BeginStoryboard>
                                        <Storyboard>
                                            <DoubleAnimation Storyboard.TargetName="arrowTransform" 
                                                             Storyboard.TargetProperty="Angle" 
                                                             To="0" Duration="0:0:0.2"/>
                                        </Storyboard>
                                    </BeginStoryboard>
                                </Trigger.ExitActions>
                            </Trigger>
                            <Trigger SourceName="HeaderSite" Property="IsMouseOver" Value="True">
                                <Setter TargetName="HeaderSite" Property="Background" Value="#F1F5F9"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="70"/>
            <!-- Modern Header -->
            <RowDefinition Height="*"/>
            <!-- Content Area -->
            <RowDefinition Height="60"/>
            <!-- Enhanced Footer -->
        </Grid.RowDefinitions>

        <!-- Modern Header (Grid.Row="0") -->
        <Border Grid.Row="0" Background="White" BorderThickness="0,0,0,1" BorderBrush="#E1E5E9">
            <Grid Margin="24,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- App Title -->
                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <Ellipse Width="36" Height="36" Fill="#0078D7" Margin="0,0,12,0">
                        <Ellipse.Effect>
                            <DropShadowEffect BlurRadius="6" ShadowDepth="1" Color="#0D000000" Opacity="0.15"/>
                        </Ellipse.Effect>
                    </Ellipse>
                    <StackPanel VerticalAlignment="Center">
                        <TextBlock Text="easyADReport" FontSize="20" FontWeight="SemiBold" Foreground="#1A202C"/>
                        <TextBlock Text="easy Active Directory Reporting" FontSize="11" Foreground="#718096" Margin="0,-2,0,0"/>
                    </StackPanel>
                </StackPanel>

                <!-- Statistics Card -->
                <Border Grid.Column="2" Style="{StaticResource ModernCard}" Padding="16,8" MinWidth="140">
                    <StackPanel>
                        <TextBlock Text="Total Results" FontSize="11" Foreground="#718096" HorizontalAlignment="Center"/>
                        <TextBlock x:Name="TotalResultCountText" Text="0" FontSize="24" FontWeight="Bold" 
                                   Foreground="#0078D7" HorizontalAlignment="Center" Margin="0,2,0,0"/>
                    </StackPanel>
                </Border>
            </Grid>
        </Border>

        <!-- Content Area (Grid.Row="1") -->
        <Grid Grid.Row="1" Margin="24,20,24,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="320"/>
                <!-- Enhanced Sidebar - erweitert für bessere Lesbarkeit -->
                <ColumnDefinition Width="24"/>
                <!-- Spacing - vergrößert -->
                <ColumnDefinition Width="*"/>
                <!-- Main Content -->
            </Grid.ColumnDefinitions>

            <!-- Modern Sidebar with Categorized Reports -->
            <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Grid.ColumnSpan="2" Margin="0,0,10,0">
                <Border Style="{StaticResource ModernCard}" Padding="16,20">
                    <StackPanel>
                        <!-- Sidebar Header -->
                        <Grid Margin="0,0,0,20">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <TextBlock Text="📊 Quick Reports" Style="{StaticResource CategoryHeader}" 
                                       FontSize="16" FontWeight="Bold" Grid.Row="0" Margin="0,0,0,8"/>
                            <TextBlock Text="Select a predefined report to execute instantly" 
                                       FontSize="12" Foreground="#6B7280" Grid.Row="1" TextWrapping="Wrap"/>
                        </Grid>

                        <!-- Users Category -->
                        <Expander Header="👤 Users" Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonQuickAllUsers" Content="All Users" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickDisabledUsers" Content="Disabled Users" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickLockedUsers" Content="Locked Users" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickNeverExpire" Content="Password Never Expires" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickInactiveUsers" Content="Inactive Users (90+ days)" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickAdminUsers" Content="Administrators" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickRecentlyCreatedUsers" Content="Recently Created (30d)" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickPasswordExpiringSoon" Content="Password Expiring (7d)" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickExpiredPasswords" Content="Expired Passwords" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickNeverLoggedOn" Content="Never Logged On" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickRecentlyDeletedUsers" Content="Recently Deleted" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickRecentlyModifiedUsers" Content="Recently Modified" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickInactiveUsersXDays" Content="Inactive Users (X days)" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickUsersWithoutManager" Content="Users without Manager" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickUsersMissingRequiredAttributes" Content="Missing Required Attributes" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickUsersDuplicateLogonNames" Content="Duplicate Logon Names" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickOrphanedSIDsUsers" Content="Orphaned SIDs" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickStalePasswords" Content="Stale Passwords" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickNeverChangingPasswords" Content="Never Changing Passwords" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickExpiringAccounts" Content="Expiring Accounts" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickReversibleEncryption" Content="Reversible Encryption" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickKerberosDES" Content="Kerberos DES Users" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickUsersWithSPN" Content="Users with SPN" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickGuestAccountStatus" Content="Guest Account Status" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickUsersByDepartment" Content="Users by Department" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickUsersByManager" Content="Users by Manager" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickRemoteAccessUsers" Content="Remote Access Users" Style="{StaticResource SidebarButton}"/>
                            </StackPanel>
                        </Expander>

                        <!-- Groups Category -->
                        <Expander Header="👥 Groups" Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonQuickGroups" Content="All Groups" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickSecurityGroups" Content="Security Groups" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickDistributionGroups" Content="Distribution Lists" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickEmptyGroups" Content="Empty Groups" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickNestedGroups" Content="Nested Groups" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickCircularGroups" Content="Circular References" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickGroupsByTypeScope" Content="Groups by Type/Scope" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickDynamicDistGroups" Content="Dynamic Distribution Groups" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickMailEnabledGroups" Content="Mail Enabled Groups" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickGroupsWithoutOwners" Content="Groups without Owners" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickLargeGroups" Content="Large Groups" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickRecentlyModifiedGroups" Content="Recently Modified Groups" Style="{StaticResource SidebarButton}"/>
                            </StackPanel>
                        </Expander>

                        <!-- Computers Category -->
                        <Expander Header="💻 Computers" Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonQuickComputers" Content="All Computers" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickInactiveComputers" Content="Inactive Computers (90+ days)" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickOSSummary" Content="Operating System Summary" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickComputersByOSVersion" Content="Computers by OS Version" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickBitLockerStatus" Content="BitLocker Status" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickComputersNeverLoggedOn" Content="Computers Never Logged On" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickDuplicateComputerNames" Content="Duplicate Computer Names" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickComputersByLocation" Content="Computers by Location" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickVirtualVsPhysical" Content="Virtual vs Physical" Style="{StaticResource SidebarButton}"/>
                            </StackPanel>
                        </Expander>

                        <!-- Security Audit Category -->
                        <Expander Header="🛡️ Security Audit" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickWeakPasswordPolicy" Content="Weak Password Policies" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickRiskyGroupMemberships" Content="Risky Memberships" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickPrivilegedAccounts" Content="Privileged Accounts" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickKerberoastable" Content="Kerberoastable Accounts" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickASREPRoastable" Content="ASREPRoastable Accounts" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickServiceAccountsOverview" Content="Service Accounts Overview" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickGPOOverview" Content="GPO Overview" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickHoneyTokens" Content="Honey Tokens" Style="{StaticResource SidebarButton}"/>
                            </StackPanel>
                        </Expander>

                        <!-- AD Health Category -->
                        <Expander Header="🏥 AD Health" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickFSMORoles" Content="FSMO Role Holders" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickDCStatus" Content="Domain Controller Status" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickReplicationStatus" Content="Replication Status" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickSYSVOLHealth" Content="SYSVOL Health Check" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickDNSHealth" Content="DNS Health Analysis" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickBackupStatus" Content="Backup Readiness" Style="{StaticResource SidebarButton}"/>
                            </StackPanel>
                        </Expander>

                        <!-- Advanced Analysis Category -->
                        <Expander Header="🔬 Advanced Analysis" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickOUHierarchy" Content="OU Hierarchy" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickSitesSubnets" Content="Sites &amp; Subnets" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickTrustRelationships" Content="Trust Relationships" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickSchemaAnalysis" Content="Schema Extensions" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickCertificateAnalysis" Content="Certificate Security" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickQuotasLimits" Content="Quotas &amp; Limits" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickDCSyncRights" Content="DCSync Rights Analysis" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickSchemaAdmins" Content="Schema Admin Paths" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickDelegation" Content="Delegation Analysis" Style="{StaticResource SidebarButton}"/>
                            </StackPanel>
                        </Expander>

                        <!-- Service Accounts Category -->
                        <Expander Header="⚙️ Service Accounts" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickManagedServiceAccounts" Content="Managed Service Accounts" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickServiceAccountsSPN" Content="Service Accounts with SPN" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickHighPrivServiceAccounts" Content="High Privilege Service Accounts" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickServiceAccountPasswordAge" Content="Service Account Password Age" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickUnusedServiceAccounts" Content="Unused Service Accounts" Style="{StaticResource SidebarButton}"/>
                            </StackPanel>
                        </Expander>

                        <!-- GPO & Policies Category -->
                        <Expander Header="📋 GPO &amp; Policies" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickUnlinkedGPOs" Content="Unlinked GPOs" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickEmptyGPOs" Content="Empty GPOs" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickGPOPermissions" Content="GPO Permissions" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickPasswordPolicySummary" Content="Password Policy Summary" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickAccountLockoutPolicies" Content="Account Lockout Policies" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickFineGrainedPasswordPolicies" Content="Fine-Grained Password Policies" Style="{StaticResource SidebarButton}"/>
                            </StackPanel>
                        </Expander>

                        <!-- Advanced Security Category -->
                        <Expander Header="🔐 Advanced Security" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickPrivilegeEscalation" Content="Privilege Escalation Paths" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickExposedCredentials" Content="Exposed Credentials" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickSuspiciousLogons" Content="Suspicious Logons" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickForeignSecurityPrincipals" Content="Foreign Security Principals" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickSIDHistoryAbuse" Content="SID History Abuse" Style="{StaticResource SidebarButton}"/>
                            </StackPanel>
                        </Expander>

                        <!-- Permissions & ACL Category -->
                        <Expander Header="🔑 Permissions &amp; ACL" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickACLAnalysis" Content="ACL Analysis" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickInheritanceBreaks" Content="Inheritance Breaks" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickAdminSDHolderObjects" Content="AdminSDHolder Objects" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickAdvancedDelegation" Content="Advanced Delegation" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickSchemaPermissions" Content="Schema Permissions" Style="{StaticResource SidebarButton}"/>
                            </StackPanel>
                        </Expander>

                        <!-- Compliance & Audit Category -->
                        <Expander Header="📊 Compliance &amp; Audit" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickGDPRCompliance" Content="GDPR Compliance" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickISO27001Readiness" Content="ISO 27001 Readiness" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickNISTFramework" Content="NIST Framework" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickPCIDSSRequirements" Content="PCI DSS Requirements" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickHIPAACompliance" Content="HIPAA Compliance" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickCustomComplianceTemplates" Content="Custom Compliance Templates" Style="{StaticResource SidebarButton}"/>
                            </StackPanel>
                        </Expander>

                        <!-- Statistics & Reports Category -->
                        <Expander Header="📈 Statistics &amp; Reports" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickDepartmentStats" Content="Department Statistics" Style="{StaticResource SidebarButton}"/>
                                <Button x:Name="ButtonQuickDepartmentSecurity" Content="Department Security" Style="{StaticResource SidebarButton}"/>
                            </StackPanel>
                        </Expander>
                    </StackPanel>
                </Border>
            </ScrollViewer>

            <!-- Main Content Area -->
            <Grid Grid.Column="2">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <!-- Filter Section -->
                    <RowDefinition Height="20"/>
                    <!-- Spacing -->
                    <RowDefinition Height="*"/>
                    <!-- Results -->
                </Grid.RowDefinitions>

                <!-- Enhanced Filter Section -->
                <Grid Grid.Row="0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                        <ColumnDefinition Width="25"/>
                        <ColumnDefinition Width="Auto" MinWidth="435"/>
                        <ColumnDefinition Width="20"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <!-- Advanced Filter Card -->
                    <Border Grid.Column="0" Style="{StaticResource ModernCard}" Padding="20,16">
                        <StackPanel>
                            <TextBlock Style="{StaticResource CategoryHeader}" Margin="0,0,0,12"><Run Language="de-de" Text="Search with "/><Run Text="Filter"/></TextBlock>

                            <!-- Object Type Selection -->
                            <StackPanel Orientation="Horizontal" Margin="0,0,0,12">
                                <TextBlock Text="Object Type:" VerticalAlignment="Center" Width="90" FontWeight="Medium"/>
                                <RadioButton x:Name="RadioButtonUser" Content="User" IsChecked="True" Margin="12,0" VerticalAlignment="Center"/>
                                <RadioButton x:Name="RadioButtonGroup" Content="Group" Margin="12,0" VerticalAlignment="Center"/>
                                <RadioButton x:Name="RadioButtonComputer" Content="Computer" Margin="12,0" VerticalAlignment="Center"/>
                                <RadioButton x:Name="RadioButtonGroupMemberships" Content="Memberships" Margin="12,0" VerticalAlignment="Center"/>
                            </StackPanel>

                            <!-- Filter 1 -->
                            <Grid Margin="0,0,0,8">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="80"/>
                                    <ColumnDefinition Width="140"/>
                                    <ColumnDefinition Width="100"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="Filter 1:" VerticalAlignment="Center" FontWeight="Medium"/>
                                <ComboBox x:Name="ComboBoxFilterAttribute1" Grid.Column="1" Margin="8,0" VerticalAlignment="Center"/>
                                <ComboBox x:Name="ComboBoxFilterOperator1" Grid.Column="2" Margin="8,0" VerticalAlignment="Center">
                                    <ComboBoxItem Content="Contains" IsSelected="True"/>
                                    <ComboBoxItem Content="Equals"/>
                                    <ComboBoxItem Content="StartsWith"/>
                                    <ComboBoxItem Content="EndsWith"/>
                                    <ComboBoxItem Content="NotEqual"/>
                                </ComboBox>
                                <TextBox x:Name="TextBoxFilterValue1" Grid.Column="3" Margin="8,0,0,0" VerticalAlignment="Center" Padding="8,6"/>
                            </Grid>

                            <!-- Logic Selector -->
                            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                                <TextBlock Text="Logic:" VerticalAlignment="Center" Width="80" FontWeight="Medium"/>
                                <RadioButton x:Name="RadioButtonAnd" Content="AND" IsChecked="True" Margin="8,0" VerticalAlignment="Center"/>
                                <RadioButton x:Name="RadioButtonOr" Content="OR" Margin="12,0" VerticalAlignment="Center"/>
                                <CheckBox x:Name="CheckBoxUseSecondFilter" Content="Use second filter" VerticalAlignment="Center" Margin="20,0"/>
                            </StackPanel>

                            <!-- Filter 2 -->
                            <Grid x:Name="SecondFilterPanel" IsEnabled="False">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="80"/>
                                    <ColumnDefinition Width="140"/>
                                    <ColumnDefinition Width="100"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="Filter 2:" VerticalAlignment="Center" FontWeight="Medium"/>
                                <ComboBox x:Name="ComboBoxFilterAttribute2" Grid.Column="1" Margin="8,0" VerticalAlignment="Center"/>
                                <ComboBox x:Name="ComboBoxFilterOperator2" Grid.Column="2" Margin="8,0" VerticalAlignment="Center">
                                    <ComboBoxItem Content="Contains" IsSelected="True"/>
                                    <ComboBoxItem Content="Equals"/>
                                    <ComboBoxItem Content="StartsWith"/>
                                    <ComboBoxItem Content="EndsWith"/>
                                    <ComboBoxItem Content="NotEqual"/>
                                </ComboBox>
                                <TextBox x:Name="TextBoxFilterValue2" Grid.Column="3" Margin="8,0,0,0" VerticalAlignment="Center" Padding="8,6"/>
                            </Grid>
                        </StackPanel>
                    </Border>

                    <!-- Attributes Selection Card -->
                    <Border Grid.Column="2" Style="{StaticResource ModernCard}" Padding="20,16">
                        <StackPanel>
                            <Grid Margin="0,0,0,12">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Text="Export Attributes" Style="{StaticResource CategoryHeader}" Grid.Column="0"/>
                                <StackPanel Grid.Column="1" Orientation="Horizontal">
                                    <Button x:Name="ButtonSelectAllAttributes" Content="All" Style="{StaticResource ModernButton}" 
                                            Margin="0,0,4,0" Padding="6,4" FontSize="10" ToolTip="Select all attributes"/>
                                    <Button x:Name="ButtonSelectNoneAttributes" Content="None" Style="{StaticResource ModernButton}" 
                                            Padding="6,4" FontSize="10" ToolTip="Deselect all attributes"/>
                                </StackPanel>
                            </Grid>

                            <!-- Attribute Categories -->
                            <TabControl x:Name="TabControlAttributes" Height="140" BorderThickness="0" Background="Transparent">
                                <TabItem Header="Basic" FontSize="11">
                                    <ListBox x:Name="ListBoxBasicAttributes" SelectionMode="Multiple" BorderThickness="0" Background="Transparent">
                                        <ListBoxItem Content="DisplayName" IsSelected="True"/>
                                        <ListBoxItem Content="SamAccountName" IsSelected="True"/>
                                        <ListBoxItem Content="GivenName"/>
                                        <ListBoxItem Content="Surname"/>
                                        <ListBoxItem Content="mail"/>
                                        <ListBoxItem Content="Department"/>
                                        <ListBoxItem Content="Title"/>
                                        <ListBoxItem Content="Enabled" IsSelected="True"/>
                                    </ListBox>
                                </TabItem>
                                <TabItem Header="Security" FontSize="11">
                                    <ListBox x:Name="ListBoxSecurityAttributes" SelectionMode="Multiple" BorderThickness="0" Background="Transparent">
                                        <ListBoxItem Content="LastLogonTimestamp"/>
                                        <ListBoxItem Content="PasswordExpired"/>
                                        <ListBoxItem Content="PasswordLastSet"/>
                                        <ListBoxItem Content="AccountExpirationDate"/>
                                        <ListBoxItem Content="badPwdCount"/>
                                        <ListBoxItem Content="lockoutTime"/>
                                        <ListBoxItem Content="UserAccountControl"/>
                                        <ListBoxItem Content="memberOf"/>
                                    </ListBox>
                                </TabItem>
                                <TabItem Header="Extended" FontSize="11">
                                    <ListBox x:Name="ListBoxExtendedAttributes" SelectionMode="Multiple" BorderThickness="0" Background="Transparent">
                                        <ListBoxItem Content="whenCreated"/>
                                        <ListBoxItem Content="whenChanged"/>
                                        <ListBoxItem Content="Manager"/>
                                        <ListBoxItem Content="Company"/>
                                        <ListBoxItem Content="physicalDeliveryOfficeName"/>
                                        <ListBoxItem Content="telephoneNumber"/>
                                        <ListBoxItem Content="homeDirectory"/>
                                        <ListBoxItem Content="scriptPath"/>
                                    </ListBox>
                                </TabItem>
                            </TabControl>
                        </StackPanel>
                    </Border>

                    <!-- Action Buttons -->
                    <StackPanel Grid.Column="4" VerticalAlignment="Center" MinWidth="180" Height="120">
                        <Button x:Name="ButtonQueryAD" Content="SEARCH" Style="{StaticResource PrimaryButton}" 
                                Height="55" FontSize="15" Margin="0,0,0,12" Width="175"/>
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                            <Button x:Name="ButtonExportCSV" Content="CSV" Style="{StaticResource ModernButton}" 
                                    Width="83" Height="45" Margin="0,0,8,0" Background="#FFCCD8FF"/>
                            <Button x:Name="ButtonExportHTML" Content="HTML" Style="{StaticResource ModernButton}" 
                                    Width="83" Height="45" Background="#FFCCD8FF"/>
                        </StackPanel>
                    </StackPanel>

                </Grid>

                <!-- Results Section -->
                <Border Grid.Row="2" Style="{StaticResource ModernCard}" Padding="20,16" Margin="0,0,0,10">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <!-- Results Header -->
                        <Grid Grid.Row="0" Margin="0,0,0,16">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Grid.Column="0" Text="Query Results" Style="{StaticResource CategoryHeader}"/>
                            <StackPanel Grid.Column="1" Orientation="Horizontal">
                                <Button Content="Refresh" Style="{StaticResource ModernButton}" Margin="0,0,8,0" Padding="8,4"/>
                                <Button Content="Copy" Style="{StaticResource ModernButton}" Padding="8,4"/>
                            </StackPanel>
                        </Grid>

                        <!-- Enhanced DataGrid -->
                        <DataGrid Grid.Row="1" x:Name="DataGridResults" 
                                  AutoGenerateColumns="True" 
                                  IsReadOnly="True" 
                                  BorderThickness="1" 
                                  BorderBrush="#E1E5E9"
                                  Background="White" 
                                  GridLinesVisibility="Horizontal" 
                                  RowBackground="White" 
                                  AlternatingRowBackground="#FAFBFC"
                                  HeadersVisibility="Column"
                                  CanUserSortColumns="True"
                                  CanUserReorderColumns="True"
                                  CanUserResizeColumns="True">
                            <DataGrid.ColumnHeaderStyle>
                                <Style TargetType="DataGridColumnHeader">
                                    <Setter Property="Background" Value="#F7FAFC"/>
                                    <Setter Property="Foreground" Value="#2D3748"/>
                                    <Setter Property="FontWeight" Value="SemiBold"/>
                                    <Setter Property="BorderBrush" Value="#E1E5E9"/>
                                    <Setter Property="BorderThickness" Value="0,0,1,1"/>
                                    <Setter Property="Padding" Value="12,8"/>
                                </Style>
                            </DataGrid.ColumnHeaderStyle>
                            <DataGrid.CellStyle>
                                <Style TargetType="DataGridCell">
                                    <Setter Property="Padding" Value="12,6"/>
                                    <Setter Property="BorderThickness" Value="0"/>
                                    <Style.Triggers>
                                        <Trigger Property="IsSelected" Value="True">
                                            <Setter Property="Background" Value="#E3F2FD"/>
                                            <Setter Property="Foreground" Value="#1565C0"/>
                                        </Trigger>
                                    </Style.Triggers>
                                </Style>
                            </DataGrid.CellStyle>
                        </DataGrid>
                    </Grid>
                </Border>
            </Grid>
        </Grid>

        <!-- Enhanced Footer (Grid.Row="2") -->
        <Border Grid.Row="2" Background="White" BorderThickness="0,1,0,0" BorderBrush="#E1E5E9">
            <Grid Margin="24,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- Status and Info -->
                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <Ellipse x:Name="StatusIndicator" Width="8" Height="8" Fill="#10B981" Margin="0,0,8,0"/>
                    <TextBlock x:Name="TextBlockStatus" Text="Ready" FontWeight="Medium" Foreground="#374151" Margin="0,0,16,0"/>
                    <TextBlock x:Name="TextBlockSelectedRows" Text="0 selected" FontSize="12" Foreground="#6B7280" Margin="0,0,16,0"/>
                    <TextBlock x:Name="TextBlockLastUpdate" Text="Last update: Never" FontSize="12" Foreground="#6B7280"/>
                </StackPanel>

                <!-- Version Info -->
                <StackPanel Grid.Column="3" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="easyADReport v0.3.0" FontSize="12" Foreground="#9CA3AF" VerticalAlignment="Center"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Setze die Ausgabekodierung auf UTF-8, um Probleme mit Umlauten zu vermeiden
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)

# Assembly fÃ¼r WPF laden
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms # FÃ¼r SaveFileDialog

# --- Globale AD-Gruppennamen fÃ¼r Deutsch/Englisch KompatibilitÃ¤t ---
$Global:ADGroupNames = @{
    DomainAdmins = @("Domain Admins", "DomÃ¤nen-Admins")
    EnterpriseAdmins = @("Enterprise Admins", "Organisations-Admins")
    SchemaAdmins = @("Schema Admins", "Schema-Admins")
    Administrators = @("Administrators", "Administratoren")
    AccountOperators = @("Account Operators", "Konten-Operatoren")
    ServerOperators = @("Server Operators", "Server-Operatoren")
    BackupOperators = @("Backup Operators", "Sicherungs-Operatoren")
    PrintOperators = @("Print Operators", "Druck-Operatoren")
    Replicator = @("Replicator", "Replikations-Operator")
    RemoteDesktopUsers = @("Remote Desktop Users", "Remotedesktopbenutzer")
    PowerUsers = @("Power Users", "Hauptbenutzer")
    DomainControllers = @("Domain Controllers", "DomÃ¤nencontroller")
    EnterpriseDomainControllers = @("Enterprise Domain Controllers", "Organisations-DomÃ¤nencontroller")
}

# Hilfsfunktion zum Finden von AD-Gruppen in beiden Sprachen
Function Get-ADGroupByNames {
    param(
        [string[]]$GroupNames,
        [switch]$ReturnAll
    )
    
    $foundGroups = @()
    foreach ($name in $GroupNames) {
        try {
            $group = Get-ADGroup -Filter "Name -eq '$name'" -ErrorAction SilentlyContinue
            if ($group) {
                if ($ReturnAll) {
                    $foundGroups += $group
                } else {
                    return $group
                }
            }
        } catch {
            # Ignore errors for non-existent groups
        }
    }
    
    if ($ReturnAll) {
        return $foundGroups
    }
    return $null
}

# --- Log-Funktion fÃ¼r konsistente Fehlerausgabe ---
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
    
    # StandardmÃ¤ÃŸig sowohl GUI als auch Terminal, wenn nicht explizit angegeben
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

# --- Debug-Log-Funktion fÃ¼r konsistente Debug-Ausgabe ---
Function Write-DebugLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [string]$Category = 'Debug'
    )
    
    # Debug-Ausgabe nur wenn $DebugPreference gesetzt ist
    if ($DebugPreference -ne 'SilentlyContinue') {
        Write-Debug "[$Category] $Message"
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
        [Parameter(Mandatory=$false)]
        [string]$FilterOperator = "Contains",
        [Parameter(Mandatory=$false)]
        [string]$FilterAttribute2,
        [Parameter(Mandatory=$false)]
        [string]$FilterValue2,
        [Parameter(Mandatory=$false)]
        [string]$FilterOperator2 = "Contains",
        [Parameter(Mandatory=$false)]
        [string]$FilterLogic = "AND",
        [Parameter(Mandatory=$true)]
        [System.Collections.IList]$SelectedAttributes,
        [Parameter(Mandatory=$false)]
        [string]$CustomFilter,
        [Parameter(Mandatory=$false)]
        [string]$ObjectType = "User"
    )

    # ÃœberprÃ¼fen, ob das AD-Modul verfÃ¼gbar ist
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-ADReportLog -Message "Error: Active Directory module not found." -Type Error
        return $null
    }

    try {
        # Konvertiere SelectedAttributes zu String-Array
        $PropertiesToLoad = @(
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

        # Basis-Eigenschaften hinzufÃ¼gen
        if ('DistinguishedName' -notin $PropertiesToLoad) { $PropertiesToLoad += 'DistinguishedName' }
        if ('ObjectClass' -notin $PropertiesToLoad) { $PropertiesToLoad += 'ObjectClass' }
        $PropertiesToLoad = $PropertiesToLoad | Select-Object -Unique

        if ($PropertiesToLoad.Count -eq 0) {
            $PropertiesToLoad = @("DisplayName", "SamAccountName", "ObjectClass")
        }

        # Filter erstellen
        $Filter = "*"
        
        if ($CustomFilter) {
            $Filter = $CustomFilter
        } else {
            # Erstelle Filter basierend auf den Eingaben
            $FilterPart1 = ""
            $FilterPart2 = ""
            
            # Erster Filter
            if (-not [string]::IsNullOrWhiteSpace($FilterValue) -and -not [string]::IsNullOrWhiteSpace($FilterAttribute)) {
                $FilterPart1 = Build-FilterString -Attribute $FilterAttribute -Value $FilterValue -Operator $FilterOperator
            } elseif (-not [string]::IsNullOrWhiteSpace($FilterValue) -and [string]::IsNullOrWhiteSpace($FilterAttribute)) {
                $FilterPart1 = Build-FilterString -Attribute "DisplayName" -Value $FilterValue -Operator "Contains"
            }
            
            # Zweiter Filter (wenn aktiviert)
            if (-not [string]::IsNullOrWhiteSpace($FilterValue2) -and -not [string]::IsNullOrWhiteSpace($FilterAttribute2)) {
                $FilterPart2 = Build-FilterString -Attribute $FilterAttribute2 -Value $FilterValue2 -Operator $FilterOperator2
            }
            
            # Kombiniere Filter mit UND/ODER Logik
            if ($FilterPart1 -and $FilterPart2) {
                if ($FilterLogic -eq "AND") {
                    $Filter = "$FilterPart1 -and $FilterPart2"
                } else {
                    $Filter = "$FilterPart1 -or $FilterPart2"
                }
            } elseif ($FilterPart1) {
                $Filter = $FilterPart1
            } elseif ($FilterPart2) {
                $Filter = $FilterPart2
            }
        }

        Write-ADReportLog -Message "Executing AD query with filter: $Filter" -Type Info -Terminal

        # AD-Abfrage basierend auf Objekttyp
        $Results = $null
        switch ($ObjectType) {
            "User" {
                if ($Filter -and $Filter.Trim() -eq "LockedOut -eq `$true") {
                    $LockedOutAccounts = Search-ADAccount -LockedOut -UsersOnly -ErrorAction Stop
                    if ($LockedOutAccounts) {
                        $Results = foreach ($Account in $LockedOutAccounts) {
                            try {
                                Get-ADUser -Identity $Account.SamAccountName -Properties $PropertiesToLoad -ErrorAction SilentlyContinue
                            } catch {
                                Write-Warning "Could not get details for user $($Account.SamAccountName): $($_.Exception.Message)"
                                $null
                            }
                        }
                        $Results = $Results | Where-Object {$_ -ne $null}
                    }
                } else {
                    # Debug-Ausgabe fÃ¼r Filter
                    Write-DebugLog "Executing Get-ADUser with filter: $Filter" -Category "ADQuery"
                    $Results = @(Get-ADUser -Filter $Filter -Properties $PropertiesToLoad -ErrorAction Stop)
                }
            }
            "Group" {
                $Results = @(Get-ADGroup -Filter $Filter -Properties $PropertiesToLoad -ErrorAction Stop)
            }
            "Computer" {
                $Results = @(Get-ADComputer -Filter $Filter -Properties $PropertiesToLoad -ErrorAction Stop)
            }
        }

        if ($Results) {
            $Results = $Results | Select-Object -Property $PropertiesToLoad -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
            return $Results
        } else {
            Write-ADReportLog -Message "No objects found for the specified filter." -Type Warning
            return $null
        }
    } catch {
        $ErrorMessage = "Error in AD query: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# Hilfsfunktion zum Erstellen von Filter-Strings
Function Build-FilterString {
    param(
        [string]$Attribute,
        [string]$Value,
        [string]$Operator
    )
    
    switch ($Operator) {
        "Contains" { return "$Attribute -like '*$Value*'" }
        "Equals" { return "$Attribute -eq '$Value'" }
        "StartsWith" { return "$Attribute -like '$Value*'" }
        "EndsWith" { return "$Attribute -like '*$Value'" }
        "NotEqual" { return "$Attribute -ne '$Value'" }
        default { return "$Attribute -like '*$Value*'" }
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
        $User = Get-ADUser -Identity $SamAccountName -Properties SamAccountName, Name -ErrorAction Stop | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount' # HinzugefÃ¼gt: Name fÃ¼r UserDisplayName
        if (-not $User) {
            Write-ADReportLog -Message "Benutzer $SamAccountName nicht gefunden." -Type Warning
            return $null
        }
        
        $Groups = Get-ADPrincipalGroupMembership -Identity $User -ErrorAction Stop | 
                  Get-ADGroup -Properties Name, SamAccountName, Description, GroupCategory, GroupScope -ErrorAction SilentlyContinue # HinzugefÃ¼gt: SamAccountName fÃ¼r GroupSamAccountName

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
            Write-ADReportLog -Message "Keine Gruppenmitgliedschaften fÃ¼r Benutzer $SamAccountName gefunden." -Type Info
            return [System.Collections.ArrayList]::new() # Leeres Array zurÃ¼ckgeben, um Fehler zu vermeiden
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

    # ÃœberprÃ¼fen, ob das AD-Modul verfÃ¼gbar ist
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

        if ('ObjectClass' -notin $PropertiesToLoad) { # Wichtig fÃ¼r Visualisierung und Typbestimmung
            $PropertiesToLoad += 'ObjectClass'
        }

        if ($PropertiesToLoad.Count -eq 0) {
            $PropertiesToLoad = @("Name", "SamAccountName", "GroupCategory", "GroupScope", "ObjectClass") 
        }

        $FilterToUse = "*"
        if ($CustomFilter) {
            $FilterToUse = $CustomFilter
        }
        
        Write-Host "FÃ¼hre Get-ADGroup mit Filter '$FilterToUse' und Eigenschaften '$($PropertiesToLoad -join ', ')' aus"
        $Groups = Get-ADGroup -Filter $FilterToUse -Properties $PropertiesToLoad -ErrorAction Stop

        if ($Groups) {
            # Erstelle Array mit den bereinigten Attributnamen fÃ¼r Select-Object
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
            # Dies stellt sicher, dass wir eine IEnumerable-Sammlung zurÃ¼ckgeben
            $Output = $Groups | Select-Object $SelectAttributes -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
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

# --- Funktion zum Abrufen von AD-Computerdaten ---
Function Get-ADComputerReportData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [System.Collections.IList]$SelectedAttributes,
        [Parameter(Mandatory=$false)]
        [string]$CustomFilter
    )

    # ÃœberprÃ¼fen, ob das AD-Modul verfÃ¼gbar ist
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

        if ('ObjectClass' -notin $PropertiesToLoad) { # Wichtig fÃ¼r Visualisierung und Typbestimmung
            $PropertiesToLoad += 'ObjectClass'
        }

        if ($PropertiesToLoad.Count -eq 0) {
            $PropertiesToLoad = @("Name", "DNSHostName", "OperatingSystem", "Enabled", "ObjectClass") 
        }

        $FilterToUse = "*"
        if ($CustomFilter) {
            $FilterToUse = $CustomFilter
        }
        
        Write-Host "FÃ¼hre Get-ADComputer mit Filter '$FilterToUse' und Eigenschaften '$($PropertiesToLoad -join ', ')' aus"
        $Computers = Get-ADComputer -Filter $FilterToUse -Properties $PropertiesToLoad -ErrorAction Stop

        if ($Computers) {
            # Erstelle Array mit den bereinigten Attributnamen fÃ¼r Select-Object
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
            $Output = $Computers | Select-Object $SelectAttributes -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
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
        $AllUsers = Get-ADUser -Filter * -Properties $Properties | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
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
        
        # Definiere hochprivilegierte Gruppen aus zentraler Konfiguration
        $RiskyGroups = foreach ($groupType in $Global:ADGroupNames.GetEnumerator()) {
            $groupType.Value
        }
        
        $RiskyUsers = [System.Collections.Generic.List[PSObject]]::new()
        
        # Gruppenanalyse mit erweiterter Fehlerbehandlung
        foreach ($groupName in $RiskyGroups) {
            try {
                $group = Get-ADGroup -Filter "Name -eq '$groupName'" -ErrorAction Stop
                
                $members = Get-ADGroupMember -Identity $group.SamAccountName -ErrorAction SilentlyContinue |
                    Get-ADObject -Properties DisplayName, SamAccountName, ObjectClass -ErrorAction SilentlyContinue |
                    Where-Object { $_.objectClass -eq "user" }
                
                foreach ($member in $members) {
                    try {
                        $userDetails = Get-ADUser -Identity $member.SamAccountName -Properties DisplayName, Enabled, LastLogonDate, PasswordLastSet -ErrorAction Stop |
                            Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
                        
                        # Dynamische Risikobewertung
                        $riskLevel = switch -Wildcard ($group.Name) {
                            { $_ -in $Global:ADGroupNames['DomainAdmins'], $Global:ADGroupNames['EnterpriseAdmins'], $Global:ADGroupNames['SchemaAdmins'] } { 
                                "Kritisch"
                                break
                            }
                            { $_ -in $Global:ADGroupNames['Administrators'] } { 
                                "Hoch"
                                break
                            }
                            { $_ -in $Global:ADGroupNames['AccountOperators'], $Global:ADGroupNames['ServerOperators'], $Global:ADGroupNames['BackupOperators'] } { 
                                "Mittel"
                                break
                            }
                            default { "Niedrig" }
                        }
                        
                        # Empfehlungsgenerierung
                        $recommendation = switch ($true) {
                            (-not $userDetails.Enabled) { "Deaktiviertes Konto aus Gruppe entfernen" }
                            ($userDetails.LastLogonDate -lt (Get-Date).AddDays(-90)) { "Inaktives Konto Ã¼berprÃ¼fen" }
                            default { "Berechtigungen regelmÃ¤ÃŸig Ã¼berwachen" }
                        }
                        
                        # Objekterstellung mit Typisierung
                        $riskUser = [PSCustomObject]@{
                            DisplayName      = $userDetails.DisplayName
                            SamAccountName   = $userDetails.SamAccountName
                            Enabled          = [bool]$userDetails.Enabled
                            LastLogonDate    = $userDetails.LastLogonDate
                            PasswordLastSet  = $userDetails.PasswordLastSet
                            RisikoGruppe     = $group.Name
                            Risikostufe      = $riskLevel
                            Empfehlung       = $recommendation
                        }
                        
                        $RiskyUsers.Add($riskUser)
                    }
                    catch {
                        Write-ADReportLog -Message "Fehler bei Benutzer $($member.SamAccountName): $($_.Exception.Message)" -Type Warning
                    }
                }
            }
            catch {
                Write-ADReportLog -Message "Gruppenanalyse fehlgeschlagen fÃ¼r '$groupName': $($_.Exception.Message)" -Type Warning
            }
        }
        
        # Deduplizierung und Ausgabe
        $UniqueRiskyUsers = $RiskyUsers | Sort-Object SamAccountName -Unique
        
        Write-ADReportLog -Message "Gruppenanalyse abgeschlossen. Gefundene RisikofÃ¤lle: $($UniqueRiskyUsers.Count)" -Type Info -Terminal
        return $UniqueRiskyUsers
    }
    catch {
        $ErrorMessage = "Gesamtfehler Gruppenanalyse: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -Terminal
        return $null
    }
}

Function Get-PrivilegedAccounts {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analysiere Konten mit erhÃ¶hten Rechten..." -Type Info -Terminal
        
        # Eigenschaften fÃ¼r privilegierte Konten
        $Properties = @(
            "DisplayName", "SamAccountName", "Enabled", "AdminCount", 
            "LastLogonDate", "PasswordLastSet", "PasswordNeverExpires",
            "ServicePrincipalNames", "TrustedForDelegation", "TrustedToAuthForDelegation"
        )
        
        # Alle Benutzer mit AdminCount = 1 (historisch privilegiert)
        $AdminCountUsers = Get-ADUser -Filter "AdminCount -eq 1" -Properties $Properties | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
        
        # Service-Konten (Konten mit SPNs)
        $ServiceAccounts = Get-ADUser -Filter "ServicePrincipalNames -like '*'" -Properties $Properties | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
        
        # Konten mit Delegierungsrechten
        $DelegationAccounts = Get-ADUser -Filter "TrustedForDelegation -eq `$true -or TrustedToAuthForDelegation -eq `$true" -Properties $Properties | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
        
        # Alle privilegierten Konten zusammenfÃ¼hren
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
        
        Write-ADReportLog -Message "$($UniquePrivilegedAccounts.Count) Konten mit erhÃ¶hten Rechten gefunden." -Type Info -Terminal
        return $UniquePrivilegedAccounts
        
    } catch {
        $ErrorMessage = "Fehler beim Analysieren der privilegierten Konten: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Organisationsstruktur-Reports ---
Function Get-DepartmentStatistics {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing department statistics..." -Type Info -Terminal
        
        # Lade alle Benutzer mit Abteilungsinformationen
        $Users = Get-ADUser -Filter * -Properties Department, Enabled, LastLogonDate, whenCreated, PasswordLastSet, PasswordNeverExpires, LockedOut, Title -ErrorAction Stop
        
        # Gruppiere nach Abteilung
        $DepartmentGroups = $Users | Group-Object Department
        
        $DepartmentStats = foreach ($dept in $DepartmentGroups) {
            $deptName = if ([string]::IsNullOrWhiteSpace($dept.Name)) { "(No Department)" } else { $dept.Name }
            $deptUsers = $dept.Group
            
            # Statistiken berechnen
            $enabledCount = ($deptUsers | Where-Object { $_.Enabled -eq $true }).Count
            $disabledCount = ($deptUsers | Where-Object { $_.Enabled -eq $false }).Count
            $lockedCount = ($deptUsers | Where-Object { $_.LockedOut -eq $true }).Count
            $neverExpireCount = ($deptUsers | Where-Object { $_.PasswordNeverExpires -eq $true }).Count
            $inactiveCount = ($deptUsers | Where-Object { $_.LastLogonDate -and $_.LastLogonDate -lt (Get-Date).AddDays(-90) }).Count
            
            # Durchschnittliches Kontoalter
            $accountAgesList = New-Object System.Collections.ArrayList
            foreach ($user in $deptUsers) {
                if ($user.whenCreated) {
                    try {
                        $age = (New-TimeSpan -Start $user.whenCreated -End (Get-Date)).Days
                        [void]$accountAgesList.Add($age)
                    } catch {
                        # Skip if date calculation fails
                    }
                }
            }
            $avgAccountAge = if ($accountAgesList.Count -gt 0) { 
                [math]::Round(($accountAgesList | Measure-Object -Average).Average, 1)
            } else { 0 }
            
            # Titel-Statistik
            $titlesList = New-Object System.Collections.ArrayList
            foreach ($user in $deptUsers) {
                if ($user.Title) {
                    if ($user.Title -is [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
                        if ($user.Title.Count -gt 0) {
                            [void]$titlesList.Add($user.Title[0].ToString())
                        }
                    } else {
                        [void]$titlesList.Add($user.Title.ToString())
                    }
                }
            }
            $uniqueTitles = @($titlesList | Select-Object -Unique).Count
            
            # Sicherheits-Score berechnen
            $totalUsers = $deptUsers.Count
            $securityScore = 100
            if ($totalUsers -gt 0) {
                $issueCount = $disabledCount + $lockedCount + $neverExpireCount + $inactiveCount
                $securityScore = [math]::Round((100 - ($issueCount / $totalUsers * 100)), 1)
            }
            
            [PSCustomObject]@{
                Department = $deptName
                TotalUsers = $totalUsers
                EnabledUsers = $enabledCount
                DisabledUsers = $disabledCount
                LockedUsers = $lockedCount
                InactiveUsers = $inactiveCount
                PasswordNeverExpires = $neverExpireCount
                AvgAccountAgeDays = $avgAccountAge
                UniqueTitles = $uniqueTitles
                SecurityScore = $securityScore
            }
        }
        
        Write-ADReportLog -Message "Department statistics analysis completed. $($DepartmentStats.Count) departments found." -Type Info -Terminal
        return $DepartmentStats | Sort-Object TotalUsers -Descending
        
    } catch {
        $ErrorMessage = "Error analyzing department statistics: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-DepartmentSecurityRisks {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing security risks by department..." -Type Info -Terminal
        
        # Lade alle Benutzer mit erweiterten Eigenschaften
        $Users = Get-ADUser -Filter * -Properties Department, Enabled, LastLogonDate, PasswordLastSet, PasswordNeverExpires, 
                                                  LockedOut, AdminCount, ServicePrincipalNames, DoesNotRequirePreAuth,
                                                  TrustedForDelegation, AllowReversiblePasswordEncryption -ErrorAction Stop
        
        # Gruppiere nach Abteilung
        $DepartmentGroups = $Users | Group-Object Department
        
        $DepartmentRisks = foreach ($dept in $DepartmentGroups) {
            $deptName = if ([string]::IsNullOrWhiteSpace($dept.Name)) { "(No Department)" } else { $dept.Name }
            $deptUsers = $dept.Group
            
            # Risiko-Metriken
            $adminUsers = @($deptUsers | Where-Object { $_.AdminCount -eq 1 }).Count
            $serviceAccounts = @($deptUsers | Where-Object { 
                $_.ServicePrincipalNames -and 
                ($_.ServicePrincipalNames.Count -gt 0 -or $_.ServicePrincipalNames -ne $null)
            }).Count
            $kerberoastable = @($deptUsers | Where-Object { 
                $_.ServicePrincipalNames -and 
                ($_.ServicePrincipalNames.Count -gt 0 -or $_.ServicePrincipalNames -ne $null) -and 
                $_.Enabled -eq $true
            }).Count
            $asrepRoastable = @($deptUsers | Where-Object { $_.DoesNotRequirePreAuth -eq $true }).Count
            $delegationEnabled = @($deptUsers | Where-Object { $_.TrustedForDelegation -eq $true }).Count
            $reversiblePwd = @($deptUsers | Where-Object { $_.AllowReversiblePasswordEncryption -eq $true }).Count
            $neverExpire = @($deptUsers | Where-Object { $_.PasswordNeverExpires -eq $true }).Count
            $oldPasswords = @($deptUsers | Where-Object { $_.PasswordLastSet -and $_.PasswordLastSet -lt (Get-Date).AddDays(-180) }).Count
            
            # Risiko-Score berechnen
            $riskScore = 0
            $riskScore += $adminUsers * 3
            $riskScore += $kerberoastable * 2
            $riskScore += $asrepRoastable * 3
            $riskScore += $delegationEnabled * 2
            $riskScore += $reversiblePwd * 4
            $riskScore += [int]($neverExpire * 0.5)
            $riskScore += [int]($oldPasswords * 0.3)
            
            $riskLevel = switch ($riskScore) {
                {$_ -ge 20} { "Critical" }
                {$_ -ge 10} { "High" }
                {$_ -ge 5} { "Medium" }
                {$_ -gt 0} { "Low" }
                default { "Minimal" }
            }
            
            [PSCustomObject]@{
                Department = $deptName
                TotalUsers = $deptUsers.Count
                AdminUsers = $adminUsers
                ServiceAccounts = $serviceAccounts
                Kerberoastable = $kerberoastable
                ASREPRoastable = $asrepRoastable
                DelegationEnabled = $delegationEnabled
                ReversiblePasswords = $reversiblePwd
                PasswordNeverExpires = $neverExpire
                OldPasswords = $oldPasswords
                RiskScore = $riskScore
                RiskLevel = $riskLevel
            }
        }
        
        Write-ADReportLog -Message "Department security risk analysis completed." -Type Info -Terminal
        return $DepartmentRisks | Sort-Object RiskScore -Descending
        
    } catch {
        $ErrorMessage = "Error analyzing department security risks: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Erweiterte Kerberos-Sicherheitsanalysen ---
Function Get-KerberoastableAccounts {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing Kerberoastable accounts (accounts with SPNs)..." -Type Info -Terminal
        
        # Lade alle Benutzer mit SPNs
        $SPNUsers = Get-ADUser -Filter "ServicePrincipalNames -like '*'" -Properties ServicePrincipalNames, DisplayName, 
                                                                                     SamAccountName, Enabled, PasswordLastSet, 
                                                                                     PasswordNeverExpires, LastLogonDate, AdminCount,
                                                                                     Description, whenCreated -ErrorAction Stop
        
        $KerberoastableAccounts = foreach ($user in $SPNUsers) {
            # Analysiere SPNs
            $spns = @()
            if ($user.ServicePrincipalNames) {
                if ($user.ServicePrincipalNames -is [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
                    $spns = @($user.ServicePrincipalNames)
                } else {
                    $spns = @($user.ServicePrincipalNames)
                }
            }
            $spnList = $spns -join "; "
            $spnCount = $spns.Count
            
            # Risiko-Bewertung
            $riskFactors = @()
            $riskLevel = 0
            
            if ($user.AdminCount -eq 1) {
                $riskFactors += "Privileged Account"
                $riskLevel += 3
            }
            
            if ($user.PasswordNeverExpires) {
                $riskFactors += "Password Never Expires"
                $riskLevel += 2
            }
            
            if ($user.PasswordLastSet -and $user.PasswordLastSet -lt (Get-Date).AddDays(-365)) {
                $riskFactors += "Old Password (>1 year)"
                $riskLevel += 1
            }
            
            if ($spnCount -gt 5) {
                $riskFactors += "Multiple SPNs ($spnCount)"
                $riskLevel += 1
            }
            
            $overallRisk = switch ($riskLevel) {
                {$_ -ge 5} { "Critical" }
                {$_ -ge 3} { "High" }
                {$_ -ge 2} { "Medium" }
                default { "Low" }
            }
            
            [PSCustomObject]@{
                DisplayName = $user.DisplayName
                SamAccountName = $user.SamAccountName
                Enabled = $user.Enabled
                SPNCount = $spnCount
                ServicePrincipalNames = $spnList
                PasswordLastSet = $user.PasswordLastSet
                PasswordNeverExpires = $user.PasswordNeverExpires
                LastLogonDate = $user.LastLogonDate
                AdminAccount = ($user.AdminCount -eq 1)
                AccountAge = if ($user.whenCreated) { [math]::Round((New-TimeSpan -Start $user.whenCreated -End (Get-Date)).TotalDays) } else { "Unknown" }
                RiskFactors = $riskFactors -join "; "
                RiskLevel = $overallRisk
                Remediation = "Review SPNs, enforce strong passwords, consider managed service accounts"
            }
        }
        
        $KerberoastableAccountsArray = @($KerberoastableAccounts)
        
        Write-ADReportLog -Message "Kerberoastable accounts analysis completed. $($KerberoastableAccountsArray.Count) accounts found." -Type Info -Terminal
        return $KerberoastableAccountsArray | Sort-Object RiskLevel, AdminAccount -Descending
        
    } catch {
        $ErrorMessage = "Error analyzing Kerberoastable accounts: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-ASREPRoastableAccounts {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing ASREPRoastable accounts (pre-authentication disabled)..." -Type Info -Terminal
        
        # Lade alle Benutzer mit deaktivierter Kerberos-Vorauthentifizierung
        $ASREPUsers = Get-ADUser -Filter "DoesNotRequirePreAuth -eq `$true" -Properties DoesNotRequirePreAuth, DisplayName,
                                                                                        SamAccountName, Enabled, PasswordLastSet,
                                                                                        PasswordNeverExpires, LastLogonDate, AdminCount,
                                                                                        Description, whenCreated, Department -ErrorAction Stop
        
        $ASREPRoastableAccounts = foreach ($user in $ASREPUsers) {
            # Risiko-Bewertung
            $riskFactors = @()
            $riskLevel = 5  # Basis-Risiko fÃ¼r ASREP Roastable ist hoch
            
            if ($user.AdminCount -eq 1) {
                $riskFactors += "Privileged Account"
                $riskLevel += 3
            }
            
            if ($user.Enabled -eq $true) {
                $riskFactors += "Account Enabled"
                $riskLevel += 1
            }
            
            if ($user.PasswordNeverExpires) {
                $riskFactors += "Password Never Expires"
                $riskLevel += 2
            }
            
            if ($user.LastLogonDate -and $user.LastLogonDate -gt (Get-Date).AddDays(-30)) {
                $riskFactors += "Recently Active"
                $riskLevel += 1
            }
            
            $overallRisk = switch ($riskLevel) {
                {$_ -ge 8} { "Critical" }
                {$_ -ge 6} { "High" }
                {$_ -ge 4} { "Medium" }
                default { "Low" }
            }
            
            [PSCustomObject]@{
                DisplayName = $user.DisplayName
                SamAccountName = $user.SamAccountName
                Department = $user.Department
                Enabled = $user.Enabled
                DoesNotRequirePreAuth = $true
                PasswordLastSet = $user.PasswordLastSet
                PasswordNeverExpires = $user.PasswordNeverExpires
                LastLogonDate = $user.LastLogonDate
                AdminAccount = ($user.AdminCount -eq 1)
                AccountAge = if ($user.whenCreated) { [math]::Round((New-TimeSpan -Start $user.whenCreated -End (Get-Date)).TotalDays) } else { "Unknown" }
                RiskFactors = $riskFactors -join "; "
                RiskLevel = $overallRisk
                Remediation = "Enable Kerberos pre-authentication immediately"
                Description = $user.Description
            }
        }
        
        Write-ADReportLog -Message "ASREPRoastable accounts analysis completed. $($ASREPRoastableAccounts.Count) accounts found." -Type Info -Terminal
        return $ASREPRoastableAccounts | Sort-Object RiskLevel, AdminAccount -Descending
        
    } catch {
        $ErrorMessage = "Error analyzing ASREPRoastable accounts: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-DelegationAnalysis {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing delegation settings..." -Type Info -Terminal
        
        # Lade alle Objekte mit Delegierung
        $DelegatedObjects = @()
        
        # Unconstrained Delegation (sehr gefÃ¤hrlich)
        $UnconstrainedDelegation = Get-ADObject -Filter "TrustedForDelegation -eq `$true -and TrustedToAuthForDelegation -eq `$false" `
                                                -Properties Name, ObjectClass, DistinguishedName, whenCreated, whenChanged -ErrorAction SilentlyContinue
        
        foreach ($obj in $UnconstrainedDelegation) {
            $DelegatedObjects += [PSCustomObject]@{
                Name = $obj.Name
                ObjectType = $obj.ObjectClass
                DelegationType = "Unconstrained Delegation"
                DistinguishedName = $obj.DistinguishedName
                Risk = "Critical"
                Created = $obj.whenCreated
                LastModified = $obj.whenChanged
                Remediation = "Remove unconstrained delegation or convert to constrained delegation"
            }
        }
        
        # Constrained Delegation
        $ConstrainedDelegation = Get-ADObject -Filter "TrustedToAuthForDelegation -eq `$true" `
                                             -Properties Name, ObjectClass, DistinguishedName, msDS-AllowedToDelegateTo, whenCreated, whenChanged -ErrorAction SilentlyContinue
        
        foreach ($obj in $ConstrainedDelegation) {
            $allowedServices = if ($obj.'msDS-AllowedToDelegateTo') { $obj.'msDS-AllowedToDelegateTo' -join "; " } else { "None specified" }
            
            $DelegatedObjects += [PSCustomObject]@{
                Name = $obj.Name
                ObjectType = $obj.ObjectClass
                DelegationType = "Constrained Delegation"
                AllowedServices = $allowedServices
                DistinguishedName = $obj.DistinguishedName
                Risk = "Medium"
                Created = $obj.whenCreated
                LastModified = $obj.whenChanged
                Remediation = "Review allowed services and minimize delegation scope"
            }
        }
        
        # Resource-based Constrained Delegation
        $ResourceBasedDelegation = Get-ADObject -Filter "msDS-AllowedToActOnBehalfOfOtherIdentity -like '*'" `
                                               -Properties Name, ObjectClass, DistinguishedName, msDS-AllowedToActOnBehalfOfOtherIdentity, whenCreated, whenChanged -ErrorAction SilentlyContinue
        
        foreach ($obj in $ResourceBasedDelegation) {
            $DelegatedObjects += [PSCustomObject]@{
                Name = $obj.Name
                ObjectType = $obj.ObjectClass
                DelegationType = "Resource-based Constrained Delegation"
                DistinguishedName = $obj.DistinguishedName
                Risk = "High"
                Created = $obj.whenCreated
                LastModified = $obj.whenChanged
                Remediation = "Audit resource-based delegation permissions"
            }
        }
        
        Write-ADReportLog -Message "Delegation analysis completed. $($DelegatedObjects.Count) delegated objects found." -Type Info -Terminal
        return $DelegatedObjects | Sort-Object Risk, DelegationType
        
    } catch {
        $ErrorMessage = "Error analyzing delegation settings: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Erweiterte Privilegien-Eskalation Analyse ---
Function Get-DCSyncRights {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing DCSync rights..." -Type Info -Terminal
        
        $Domain = Get-ADDomain
        $DomainDN = $Domain.DistinguishedName
        $DCSyncRights = @()
        
        # Get ACL for domain object
        try {
            $DomainACL = Get-Acl "AD:\$DomainDN" -ErrorAction Stop
        } catch {
            Write-ADReportLog -Message "Could not access domain ACL. Requires appropriate permissions." -Type Warning -Terminal
            return @([PSCustomObject]@{
                Identity = "N/A"
                ObjectType = "Error"
                Permissions = "Could not read ACL"
                Risk = "Unknown"
                Remediation = "Run with Domain Admin privileges"
            })
        }
        
        # DCSync requires DS-Replication-Get-Changes and DS-Replication-Get-Changes-All
        $ReplicationGUIDs = @{
            "DS-Replication-Get-Changes" = "1131f6aa-9c07-11d1-f79f-00c04fc2dcd2"
            "DS-Replication-Get-Changes-All" = "1131f6ad-9c07-11d1-f79f-00c04fc2dcd2"
            "DS-Replication-Get-Changes-In-Filtered-Set" = "89e95b76-444d-4c62-991a-0facbeda640c"
        }
        
        # Analyze each ACE
        foreach ($ace in $DomainACL.Access) {
            if ($ace.AccessControlType -eq "Allow" -and $ace.ObjectType) {
                $aceObjectType = $ace.ObjectType.ToString()
                
                foreach ($guid in $ReplicationGUIDs.GetEnumerator()) {
                    if ($aceObjectType -eq $guid.Value) {
                        # Resolve the identity
                        $identity = $ace.IdentityReference.Value
                        try {
                            $adObject = Get-ADObject -Filter "SamAccountName -eq '$($identity.Split('\')[1])'" -Properties ObjectClass, AdminCount -ErrorAction SilentlyContinue
                            $objectType = if ($adObject) { $adObject.ObjectClass } else { "Unknown" }
                            $isAdmin = if ($adObject -and $adObject.AdminCount -eq 1) { $true } else { $false }
                        } catch {
                            $objectType = "Unknown"
                            $isAdmin = $false
                        }
                        
                        # Check if identity is a known privileged group
                        $isExpectedGroup = $false
                        $allExpectedGroups = @()
                        $allExpectedGroups += $Global:ADGroupNames.DomainControllers
                        $allExpectedGroups += $Global:ADGroupNames.EnterpriseDomainControllers  
                        $allExpectedGroups += $Global:ADGroupNames.Administrators
                        
                        foreach ($expectedGroup in $allExpectedGroups) {
                            if ($identity -like "*$expectedGroup*") {
                                $isExpectedGroup = $true
                                break
                            }
                        }
                        
                        $DCSyncRights += [PSCustomObject]@{
                            Identity = $identity
                            Permission = $guid.Key
                            ObjectType = $objectType
                            IsPrivileged = $isAdmin
                            Risk = if (-not $isExpectedGroup) { "High" } else { "Expected" }
                            Remediation = if (-not $isExpectedGroup) { 
                                "Review and remove DCSync rights from this identity" 
                            } else { 
                                "Standard DCSync permission - monitor for changes" 
                            }
                        }
                    }
                }
            }
        }
        
        # Group by identity to show complete DCSync capability
        $DCSyncCapable = $DCSyncRights | Group-Object Identity | Where-Object {
            ($_.Group.Permission -contains "DS-Replication-Get-Changes") -and 
            ($_.Group.Permission -contains "DS-Replication-Get-Changes-All")
        }
        
        $FinalReport = foreach ($capable in $DCSyncCapable) {
            $permissions = $capable.Group.Permission -join ", "
            $firstEntry = $capable.Group[0]
            
            [PSCustomObject]@{
                Identity = $capable.Name
                ObjectType = $firstEntry.ObjectType
                IsPrivileged = $firstEntry.IsPrivileged
                Permissions = $permissions
                HasFullDCSync = $true
                Risk = $firstEntry.Risk
                Remediation = $firstEntry.Remediation
            }
        }
        
        if ($FinalReport.Count -eq 0) {
            Write-ADReportLog -Message "No explicit DCSync rights found. This might be normal if only default permissions exist." -Type Info -Terminal
        } else {
            Write-ADReportLog -Message "DCSync rights analysis completed. $($FinalReport.Count) identities with DCSync capabilities found." -Type Info -Terminal
        }
        
        return $FinalReport
        
    } catch {
        $ErrorMessage = "Error analyzing DCSync rights: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-SchemaAdminPaths {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing Schema Admin access paths..." -Type Info -Terminal
        
        $SchemaAdminPaths = @()
        
        # Get Schema Admins group members - try both English and German names
        $SchemaAdmins = $null
        $SchemaAdminsGroup = Get-ADGroupByNames -GroupNames $Global:ADGroupNames.SchemaAdmins
        
        if ($SchemaAdminsGroup) {
            Write-ADReportLog -Message "Found Schema Admins group as: $($SchemaAdminsGroup.Name)" -Type Info -Terminal
            $SchemaAdmins = Get-ADGroupMember -Identity $SchemaAdminsGroup -Recursive -ErrorAction SilentlyContinue
        }
        
        if (-not $SchemaAdmins) {
            Write-ADReportLog -Message "Schema Admins group not found. This might be normal if the group doesn't exist." -Type Info -Terminal
            return @([PSCustomObject]@{
                Name = "N/A"
                SamAccountName = "N/A"
                ObjectType = "Information"
                Status = "Schema Admins group not found"
                Details = "The Schema Admins group might not exist or has a different name"
                Remediation = "Verify Schema Admins group existence"
            })
        }
        
        foreach ($admin in $SchemaAdmins) {
            $userDetails = Get-ADUser -Identity $admin.SamAccountName -Properties Enabled, LastLogonDate, PasswordLastSet, AdminCount, whenCreated -ErrorAction SilentlyContinue
            
            if ($userDetails) {
                $SchemaAdminPaths += [PSCustomObject]@{
                    Name = $admin.Name
                    SamAccountName = $admin.SamAccountName
                    ObjectType = $admin.ObjectClass
                    Enabled = $userDetails.Enabled
                    LastLogonDate = $userDetails.LastLogonDate
                    PasswordLastSet = $userDetails.PasswordLastSet
                    AccountAge = if ($userDetails.whenCreated) { [math]::Round((New-TimeSpan -Start $userDetails.whenCreated -End (Get-Date)).TotalDays) } else { "Unknown" }
                    Risk = if ($userDetails.Enabled -and $userDetails.LastLogonDate -gt (Get-Date).AddDays(-90)) { "Active Schema Admin" } else { "Inactive Schema Admin" }
                    Remediation = "Schema Admins should be empty except during schema modifications"
                }
            }
        }
        
        # Check for users who can add themselves to Schema Admins
        $EnterpriseAdmins = $null
        $EnterpriseAdminsGroup = Get-ADGroupByNames -GroupNames $Global:ADGroupNames.EnterpriseAdmins
        
        if ($EnterpriseAdminsGroup) {
            $EnterpriseAdmins = Get-ADGroupMember -Identity $EnterpriseAdminsGroup -Recursive -ErrorAction SilentlyContinue
        }
        
        if ($EnterpriseAdmins) {
            foreach ($ea in $EnterpriseAdmins) {
                # Check if not already in Schema Admins (if Schema Admins exists)
                $isAlreadySchemaAdmin = $false
                if ($SchemaAdmins) {
                    $isAlreadySchemaAdmin = $SchemaAdmins.SamAccountName -contains $ea.SamAccountName
                }
                
                if (-not $isAlreadySchemaAdmin) {
                    $userDetails = Get-ADUser -Identity $ea.SamAccountName -Properties Enabled, LastLogonDate -ErrorAction SilentlyContinue
                    
                    if ($userDetails) {
                        $SchemaAdminPaths += [PSCustomObject]@{
                            Name = $ea.Name
                            SamAccountName = $ea.SamAccountName
                            ObjectType = "Potential Path"
                            Enabled = $userDetails.Enabled
                            LastLogonDate = $userDetails.LastLogonDate
                            Risk = "Can elevate to Schema Admin"
                            Remediation = "Monitor Enterprise Admin membership"
                        }
                    }
                }
            }
        }
        
        Write-ADReportLog -Message "Schema Admin paths analysis completed. $($SchemaAdminPaths.Count) paths found." -Type Info -Terminal
        return $SchemaAdminPaths
        
    } catch {
        $ErrorMessage = "Error analyzing Schema Admin paths: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-CertificateSecurityAnalysis {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing Certificate Services security..." -Type Info -Terminal
        
        $CertAnalysis = @()
        
        # Check if AD CS is installed
        try {
            $Domain = Get-ADDomain
            $ConfigDN = "CN=Configuration,$($Domain.DistinguishedName)"
            $CertConfig = Get-ADObject -Filter "objectClass -eq 'pKIEnrollmentService'" -SearchBase $ConfigDN -Properties * -ErrorAction SilentlyContinue
        } catch {
            $CertConfig = $null
        }
        
        if (-not $CertConfig) {
            Write-ADReportLog -Message "No Certificate Services found in this domain." -Type Info -Terminal
            return @([PSCustomObject]@{
                Component = "AD CS Status"
                Finding = "Not Installed"
                Risk = "N/A"
                Details = "Active Directory Certificate Services not detected in this domain"
            })
        }
        
        # Analyze certificate templates
        $CertTemplates = Get-ADObject -Filter "objectClass -eq 'pKICertificateTemplate'" -Properties * -SearchBase "CN=Certificate Templates,CN=Public Key Services,CN=Services,CN=Configuration,$((Get-ADDomain).DistinguishedName)" -ErrorAction SilentlyContinue
        
        foreach ($template in $CertTemplates) {
            # Check for dangerous configurations
            $risks = @()
            $riskLevel = "Low"
            
            # Check if template allows requestor to specify SAN
            if ($template.'msPKI-Certificate-Name-Flag' -band 1) {
                $risks += "Allows SAN specification"
                $riskLevel = "High"
            }
            
            # Check for client authentication
            if ($template.'msPKI-Certificate-Application-Policy' -contains "1.3.6.1.5.5.7.3.2") {
                $risks += "Allows client authentication"
                if ($riskLevel -ne "High") { $riskLevel = "Medium" }
            }
            
            # Check enrollment permissions
            $templateACL = Get-Acl "AD:$($template.DistinguishedName)" -ErrorAction SilentlyContinue
            $enrollmentRights = $templateACL.Access | Where-Object { $_.ActiveDirectoryRights -match "ExtendedRight" }
            
            if ($enrollmentRights) {
                $enrollers = $enrollmentRights.IdentityReference.Value -join ", "
                if ($enrollers -match "Authenticated Users|Domain Users") {
                    $risks += "Wide enrollment permissions"
                    $riskLevel = "Critical"
                }
            }
            
            if ($risks.Count -gt 0) {
                $CertAnalysis += [PSCustomObject]@{
                    TemplateName = $template.Name
                    DisplayName = $template.DisplayName
                    Risks = $risks -join "; "
                    RiskLevel = $riskLevel
                    EnrollmentRights = $enrollers
                    Remediation = "Review template permissions and configuration"
                }
            }
        }
        
        # Check for web enrollment
        $WebEnrollment = Get-ADObject -Filter "objectClass -eq 'certificationAuthority'" -Properties * -ErrorAction SilentlyContinue
        
        if ($WebEnrollment) {
            $CertAnalysis += [PSCustomObject]@{
                Component = "Web Enrollment"
                Status = "Detected"
                Risk = "Medium"
                Details = "Web enrollment increases attack surface"
                Remediation = "Ensure web enrollment is properly secured with HTTPS and authentication"
            }
        }
        
        Write-ADReportLog -Message "Certificate security analysis completed. $($CertAnalysis.Count) findings." -Type Info -Terminal
        return $CertAnalysis
        
    } catch {
        $ErrorMessage = "Error analyzing Certificate Services: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Performance & Health Monitoring ---
Function Get-SYSVOLHealthCheck {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Performing SYSVOL health check..." -Type Info -Terminal
        
        $SYSVOLHealth = @()
        $DomainControllers = Get-ADDomainController -Filter * -ErrorAction Stop
        
        foreach ($DC in $DomainControllers) {
            try {
                # Test SYSVOL share accessibility
                $SYSVOLPath = "\\$($DC.HostName)\SYSVOL"
                $SYSVOLAccessible = Test-Path $SYSVOLPath -ErrorAction SilentlyContinue
                
                # Check DFSR or FRS
                $ReplicationMethod = "Unknown"
                try {
                    $DFSRCheck = Get-WmiObject -ComputerName $DC.HostName -Namespace "root\microsoftdfs" -Class "dfsrreplicatedfolderinfo" -ErrorAction SilentlyContinue
                    if ($DFSRCheck) {
                        $ReplicationMethod = "DFSR"
                    } else {
                        $FRSCheck = Get-WmiObject -ComputerName $DC.HostName -Class "Win32_Service" -Filter "Name='NtFrs'" -ErrorAction SilentlyContinue
                        if ($FRSCheck -and $FRSCheck.State -eq "Running") {
                            $ReplicationMethod = "FRS (Legacy)"
                        }
                    }
                } catch {
                    $ReplicationMethod = "Unable to determine"
                }
                
                # Check for orphaned GPO folders
                $orphanedGPOs = 0
                if ($SYSVOLAccessible) {
                    try {
                        $GPOFolders = Get-ChildItem "$SYSVOLPath\$($DC.Domain)\Policies" -Directory -ErrorAction SilentlyContinue
                        $ADGPOs = Get-GPO -All -Domain $DC.Domain -ErrorAction SilentlyContinue
                        $ADGPOIds = $ADGPOs | ForEach-Object { "{$($_.Id)}" }
                        
                        foreach ($folder in $GPOFolders) {
                            if ($folder.Name -ne "PolicyDefinitions" -and $ADGPOIds -notcontains $folder.Name) {
                                $orphanedGPOs++
                            }
                        }
                    } catch {
                        $orphanedGPOs = "Unable to check"
                    }
                }
                
                $SYSVOLHealth += [PSCustomObject]@{
                    DomainController = $DC.Name
                    SYSVOLAccessible = $SYSVOLAccessible
                    ReplicationMethod = $ReplicationMethod
                    OrphanedGPOFolders = $orphanedGPOs
                    NetlogonShare = Test-Path "\\$($DC.HostName)\NETLOGON" -ErrorAction SilentlyContinue
                    Status = if ($SYSVOLAccessible -and $ReplicationMethod -ne "Unknown") { "Healthy" } else { "Issues Detected" }
                    Remediation = if ($ReplicationMethod -eq "FRS (Legacy)") { "Migrate from FRS to DFSR" } 
                                 elseif ($orphanedGPOs -gt 0) { "Clean up orphaned GPO folders" }
                                 else { "None required" }
                }
                
            } catch {
                $SYSVOLHealth += [PSCustomObject]@{
                    DomainController = $DC.Name
                    SYSVOLAccessible = $false
                    Status = "Error"
                    Error = $_.Exception.Message
                }
            }
        }
        
        Write-ADReportLog -Message "SYSVOL health check completed for $($SYSVOLHealth.Count) domain controllers." -Type Info -Terminal
        return $SYSVOLHealth
        
    } catch {
        $ErrorMessage = "Error performing SYSVOL health check: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-DNSHealthAnalysis {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing DNS health..." -Type Info -Terminal
        
        $DNSHealth = @()
        $Domain = Get-ADDomain
        
        # Check DNS zones - only if DNS cmdlets are available
        $DNSZones = $null
        try {
            if (Get-Command Get-DnsServerZone -ErrorAction SilentlyContinue) {
                $DNSZones = Get-DnsServerZone -ComputerName $Domain.PDCEmulator -ErrorAction SilentlyContinue
            } else {
                Write-ADReportLog -Message "DNS Server cmdlets not available. Skipping DNS zone analysis." -Type Info -Terminal
            }
        } catch {
            Write-ADReportLog -Message "Could not query DNS zones: $($_.Exception.Message)" -Type Warning -Terminal
        }
        
        if ($DNSZones) {
            foreach ($zone in $DNSZones) {
            # Check zone health
            $zoneHealth = [PSCustomObject]@{
                ZoneName = $zone.ZoneName
                ZoneType = $zone.ZoneType
                IsDsIntegrated = $zone.IsDsIntegrated
                IsReverseLookup = $zone.IsReverseLookupZone
                SecureSecondaries = $zone.SecureSecondaries
                DynamicUpdate = $zone.DynamicUpdate
                Status = "Unknown"
                Issues = @()
            }
            
            # Check for security issues
            if ($zone.DynamicUpdate -eq "NonsecureAndSecure") {
                $zoneHealth.Issues += "Allows non-secure dynamic updates"
                $zoneHealth.Status = "Security Risk"
            }
            
            # Check aging/scavenging
            if ($zone.AgingEnabled) {
                $zoneHealth.ScavengingEnabled = $true
                $zoneHealth.NoRefreshInterval = $zone.NoRefreshInterval
                $zoneHealth.RefreshInterval = $zone.RefreshInterval
            } else {
                $zoneHealth.Issues += "Scavenging not enabled"
                if ($zoneHealth.Status -eq "Unknown") { $zoneHealth.Status = "Warning" }
            }
            
            if ($zoneHealth.Status -eq "Unknown") { $zoneHealth.Status = "Healthy" }
            
                $DNSHealth += $zoneHealth
            }
        }
        
        # Check for stale DNS records
        $StaleRecordCheck = [PSCustomObject]@{
            Component = "Stale Record Detection"
            Description = "Records older than 30 days"
            Status = "Manual check required"
            Remediation = "Enable scavenging or manually clean up stale records"
        }
        $DNSHealth += $StaleRecordCheck
        
        Write-ADReportLog -Message "DNS health analysis completed. $($DNSHealth.Count) items analyzed." -Type Info -Terminal
        return $DNSHealth
        
    } catch {
        $ErrorMessage = "Error analyzing DNS health: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Backup & Recovery Readiness ---
Function Get-BackupReadinessStatus {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Checking AD backup readiness..." -Type Info -Terminal
        
        $BackupStatus = @()
        $Domain = Get-ADDomain
        
        # Check System State backup on PDC
        $PDC = $Domain.PDCEmulator
        
        # Tombstone lifetime
        $ConfigNC = "CN=Configuration,$($Domain.DistinguishedName)"
        $TombstoneObject = Get-ADObject -Identity "CN=Directory Service,CN=Windows NT,CN=Services,$ConfigNC" -Properties tombstoneLifetime -ErrorAction SilentlyContinue
        $TombstoneLifetime = if ($TombstoneObject.tombstoneLifetime) { $TombstoneObject.tombstoneLifetime } else { 180 }
        
        $BackupStatus += [PSCustomObject]@{
            Component = "Tombstone Lifetime"
            Value = "$TombstoneLifetime days"
            Status = if ($TombstoneLifetime -ge 180) { "Good" } else { "Warning" }
            Details = "Backups older than tombstone lifetime cannot be restored"
            Remediation = if ($TombstoneLifetime -lt 180) { "Consider increasing tombstone lifetime" } else { "None required" }
        }
        
        # Check Deleted Objects container size
        try {
            $DeletedObjects = Get-ADObject -Filter * -IncludeDeletedObjects -SearchBase "CN=Deleted Objects,$($Domain.DistinguishedName)" -ErrorAction SilentlyContinue
            $DeletedCount = if ($DeletedObjects) { @($DeletedObjects).Count } else { 0 }
            
            $BackupStatus += [PSCustomObject]@{
                Component = "Deleted Objects"
                Value = "$DeletedCount objects"
                Status = if ($DeletedCount -gt 10000) { "Warning" } else { "Good" }
                Details = "Large number of deleted objects can impact restore performance"
                Remediation = if ($DeletedCount -gt 10000) { "Consider garbage collection" } else { "None required" }
            }
        } catch {
            $BackupStatus += [PSCustomObject]@{
                Component = "Deleted Objects"
                Value = "Unable to access"
                Status = "Unknown"
                Details = "Requires elevated permissions"
            }
        }
        
        # DSRM password age check
        $DomainControllers = Get-ADDomainController -Filter *
        
        foreach ($DC in $DomainControllers) {
            $BackupStatus += [PSCustomObject]@{
                Component = "DSRM Password"
                DomainController = $DC.Name
                Status = "Manual Check Required"
                Details = "DSRM password age cannot be checked remotely"
                Remediation = "Ensure DSRM passwords are documented and regularly updated"
            }
        }
        
        # Backup GPO existence
        $BackupGPOs = Get-GPO -All | Where-Object { $_.DisplayName -like "*Backup*" -or $_.DisplayName -like "*Restore*" }
        
        $BackupStatus += [PSCustomObject]@{
            Component = "Backup Procedures"
            Value = if ($BackupGPOs) { "$($BackupGPOs.Count) backup-related GPOs found" } else { "No backup GPOs found" }
            Status = if ($BackupGPOs) { "Good" } else { "Warning" }
            Details = "GPOs can help standardize backup procedures"
            Remediation = if (-not $BackupGPOs) { "Consider creating backup procedure GPOs" } else { "Review existing backup GPOs" }
        }
        
        Write-ADReportLog -Message "Backup readiness check completed." -Type Info -Terminal
        return $BackupStatus
        
    } catch {
        $ErrorMessage = "Error checking backup readiness: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Schema & Trusts Analysis ---
Function Get-SchemaAnalysis {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing AD schema..." -Type Info -Terminal
        
        $SchemaAnalysis = @()
        $RootDSE = Get-ADRootDSE
        $SchemaNC = $RootDSE.schemaNamingContext
        
        # Get schema version
        $Schema = Get-ADObject -Identity $SchemaNC -Properties objectVersion
        $SchemaVersion = switch ($Schema.objectVersion) {
            87 { "Windows Server 2016" }
            88 { "Windows Server 2019" }
            89 { "Windows Server 2022" }
            default { "Version $($Schema.objectVersion)" }
        }
        
        $SchemaAnalysis += [PSCustomObject]@{
            Component = "Schema Version"
            Value = $SchemaVersion
            VersionNumber = $Schema.objectVersion
            Status = "Information"
        }
        
        # Count custom schema extensions
        $AllSchemaObjects = Get-ADObject -Filter * -SearchBase $SchemaNC -Properties whenCreated, adminDescription
        $CustomSchema = $AllSchemaObjects | Where-Object { 
            $_.whenCreated -and 
            $_.adminDescription -notlike "Microsoft*" -and 
            $_.Name -notlike "ms-*" 
        }
        
        $SchemaAnalysis += [PSCustomObject]@{
            Component = "Custom Extensions"
            Value = "$($CustomSchema.Count) custom schema objects"
            Status = if ($CustomSchema.Count -gt 50) { "Review Needed" } else { "Normal" }
            Details = "High number of custom extensions may indicate complex environment"
        }
        
        # Recent schema changes
        $RecentChanges = $AllSchemaObjects | Where-Object { 
            $_.whenCreated -gt (Get-Date).AddDays(-90) 
        }
        
        if ($RecentChanges) {
            foreach ($change in $RecentChanges) {
                $SchemaAnalysis += [PSCustomObject]@{
                    Component = "Recent Change"
                    ObjectName = $change.Name
                    Created = $change.whenCreated
                    Description = $change.adminDescription
                    Status = "Monitor"
                }
            }
        }
        
        Write-ADReportLog -Message "Schema analysis completed." -Type Info -Terminal
        return $SchemaAnalysis
        
    } catch {
        $ErrorMessage = "Error analyzing schema: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-TrustRelationshipAnalysis {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing trust relationships..." -Type Info -Terminal
        
        $TrustAnalysis = @()
        
        # Get all trusts
        $Trusts = Get-ADTrust -Filter * -Properties * -ErrorAction Stop
        
        foreach ($trust in $Trusts) {
            # Analyze trust properties
            $trustHealth = "Unknown"
            $issues = @()
            
            # Check trust type and transitivity
            if ($trust.TrustType -eq "External" -and $trust.Transitivity -eq "Transitive") {
                $issues += "External trust is transitive"
                $trustHealth = "Security Risk"
            }
            
            # Check SID filtering
            if ($trust.SIDFilteringForestAware -eq $false -or $trust.SIDFilteringQuarantined -eq $false) {
                $issues += "SID filtering may be disabled"
                if ($trustHealth -ne "Security Risk") { $trustHealth = "Warning" }
            }
            
            # Check selective authentication
            if ($trust.SelectiveAuthentication -eq $false -and $trust.TrustType -eq "External") {
                $issues += "Selective authentication not enabled"
                if ($trustHealth -eq "Unknown") { $trustHealth = "Review Needed" }
            }
            
            if ($trustHealth -eq "Unknown" -and $issues.Count -eq 0) { $trustHealth = "Healthy" }
            
            $TrustAnalysis += [PSCustomObject]@{
                TrustPartner = $trust.Target
                Direction = $trust.Direction
                TrustType = $trust.TrustType
                Transitivity = if ($trust.Transitivity) { "Transitive" } else { "Non-Transitive" }
                SIDFiltering = if ($trust.SIDFilteringQuarantined) { "Enabled" } else { "Check Required" }
                SelectiveAuth = if ($trust.SelectiveAuthentication) { "Enabled" } else { "Disabled" }
                Created = $trust.Created
                Status = $trustHealth
                Issues = $issues -join "; "
                Remediation = if ($issues.Count -gt 0) { "Review trust configuration for security best practices" } else { "None required" }
            }
        }
        
        # Forest trust insights
        $Forest = Get-ADForest
        if ($Forest.Domains.Count -gt 1) {
            $TrustAnalysis += [PSCustomObject]@{
                Component = "Forest Structure"
                Domains = $Forest.Domains.Count
                GlobalCatalogs = $Forest.GlobalCatalogs.Count
                Sites = $Forest.Sites.Count
                Status = "Multi-Domain Forest"
                Details = "Complex forest structure requires careful trust management"
            }
        }
        
        Write-ADReportLog -Message "Trust relationship analysis completed. $($Trusts.Count) trusts analyzed." -Type Info -Terminal
        return $TrustAnalysis
        
    } catch {
        $ErrorMessage = "Error analyzing trust relationships: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-QuotasAndLimits {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing AD quotas and limits..." -Type Info -Terminal
        
        $QuotaAnalysis = @()
        $Domain = Get-ADDomain
        
        # RID Pool check
        $DomainControllers = Get-ADDomainController -Filter *
        foreach ($DC in $DomainControllers) {
            try {
                $RIDInfo = Get-ADObject -Identity "CN=RID Manager$,CN=System,$($Domain.DistinguishedName)" -Properties rIDAvailablePool -Server $DC.HostName -ErrorAction SilentlyContinue
                
                if ($RIDInfo -and $RIDInfo.rIDAvailablePool) {
                    # RID Pool Berechnung
                    try {
                        # RID Pool value is a large integer property
                        if ($RIDInfo.rIDAvailablePool -is [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
                            $ridPoolValue = [int64]$RIDInfo.rIDAvailablePool[0]
                        } else {
                            $ridPoolValue = [int64]$RIDInfo.rIDAvailablePool
                        }
                        
                        # Calculate using int64 to avoid overflow
                        [int64]$pow32 = [math]::Pow(2, 32)
                        [int64]$pow30 = [math]::Pow(2, 30)
                        
                        [int64]$totalSIDS = [math]::Floor($ridPoolValue / $pow32)
                        [int64]$totalRIDS = $totalSIDS * $pow30
                        [int64]$currentRIDPoolCount = $ridPoolValue % $pow32
                        
                        if ($totalRIDS -gt 0) {
                            $percentUsed = [math]::Round((($totalRIDS - $currentRIDPoolCount) / $totalRIDS * 100), 2)
                        } else {
                            $percentUsed = 0
                        }
                    } catch {
                        Write-Warning "Error calculating RID pool for $($DC.Name): $($_.Exception.Message)"
                        continue
                    }
                    
                    $QuotaAnalysis += [PSCustomObject]@{
                        Component = "RID Pool"
                        DomainController = $DC.Name
                        TotalRIDs = $totalRIDS
                        RemainingRIDs = $currentRIDPoolCount
                        PercentUsed = $percentUsed
                        Status = if ($percentUsed -gt 80) { "Critical" } elseif ($percentUsed -gt 60) { "Warning" } else { "Healthy" }
                        Remediation = if ($percentUsed -gt 80) { "RID pool depletion risk - plan for RID recovery" } else { "Monitor RID usage" }
                    }
                }
            } catch {
                Write-Warning "Could not get RID pool information from $($DC.Name)"
            }
        }
        
        # LDAP Query Policy
        $DefaultPolicy = Get-ADObject -Identity "CN=Default Query Policy,CN=Query-Policies,CN=Directory Service,CN=Windows NT,CN=Services,CN=Configuration,$($Domain.DistinguishedName)" -Properties * -ErrorAction SilentlyContinue
        
        if ($DefaultPolicy) {
            $QuotaAnalysis += [PSCustomObject]@{
                Component = "LDAP Query Policy"
                MaxPageSize = $DefaultPolicy.lDAPAdminLimits | Where-Object { $_ -like "MaxPageSize=*" } | ForEach-Object { $_.Split('=')[1] }
                MaxQueryDuration = $DefaultPolicy.lDAPAdminLimits | Where-Object { $_ -like "MaxQueryDuration=*" } | ForEach-Object { $_.Split('=')[1] }
                Status = "Configured"
                Details = "Default LDAP query limits in effect"
            }
        }
        
        # Token size estimation
        $LargeTokenUsers = Get-ADUser -Filter * -Properties MemberOf | Where-Object { $_.MemberOf.Count -gt 100 }
        
        if ($LargeTokenUsers) {
            $QuotaAnalysis += [PSCustomObject]@{
                Component = "Token Size Risk"
                UsersAtRisk = $LargeTokenUsers.Count
                Status = if ($LargeTokenUsers.Count -gt 50) { "Warning" } else { "Monitor" }
                Details = "Users with >100 group memberships may experience token size issues"
                Remediation = "Review group membership strategy for highly nested users"
            }
        }
        
        Write-ADReportLog -Message "Quotas and limits analysis completed." -Type Info -Terminal
        return $QuotaAnalysis
        
    } catch {
        $ErrorMessage = "Error analyzing quotas and limits: $($_.Exception.Message)"
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
        
        # DomÃ¤nen-spezifische FSMO-Rollen
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
        
        # AD-Health Sammlung - Ã¤hnlich wie dxdiag
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
            # PrÃ¼fe ob es eine Collection ist und hole den ersten Wert
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
            
            # Konvertiere zu Array und zÃ¤hle dann
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
        
        Write-ADReportLog -Message "AD-Health Diagnose abgeschlossen: $($FinalReport.Count) Checks durchgefÃ¼hrt." -Type Info -Terminal
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
            Partition = "N/A"
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

# --- Funktion zum Abrufen von OU-Hierarchie-Berichten ---
Function Get-ADOUHierarchyReport {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Generating OU Hierarchy Report..." -Type Info -Terminal
    try {
        $AllOUs = Get-ADOrganizationalUnit -Filter * -Properties DistinguishedName, Name, CanonicalName, whenCreated, whenChanged, Description, ProtectedFromAccidentalDeletion -ErrorAction SilentlyContinue
        if (-not $AllOUs) {
            Write-ADReportLog -Message "No Organizational Units found in the domain." -Type Warning -Terminal
            return $null
        }

        $OUReport = foreach ($OU in $AllOUs) {
            $ParentPath = "Domain Root"
            try {
                # Versuche, den Parent DN zu bekommen und daraus den Namen abzuleiten
                $parentDN = $OU.DistinguishedName.Substring($OU.DistinguishedName.IndexOf(',') + 1)
                if ($parentDN -and $parentDN -ne $OU.DistinguishedName) {
                    # PrÃ¼fe ob es eine Domain-Root OU ist (direkt unter DC=...)
                    if ($parentDN -match '^DC=') {
                        # FÃ¼r Root-OUs zeige "Domain Root" oder den Domain Namen
                        $domain = Get-ADDomain -ErrorAction SilentlyContinue
                        if ($domain) {
                            $ParentPath = "Domain Root ($($domain.NetBIOSName))"
                        } else {
                            $ParentPath = "Domain Root"
                        }
                    } else {
                        # Versuche Parent OU zu finden
                        $parentOUObj = Get-ADOrganizationalUnit -Identity $parentDN -Properties Name -ErrorAction SilentlyContinue
                        if ($parentOUObj) {
                            $ParentPath = $parentOUObj.Name
                        } else {
                            # MÃ¶glicherweise ein Container, kein OU
                            try {
                                $parentObj = Get-ADObject -Identity $parentDN -Properties Name -ErrorAction SilentlyContinue
                                if ($parentObj) {
                                    $ParentPath = $parentObj.Name
                                } else {
                                    $ParentPath = $parentDN
                                }
                            } catch {
                                $ParentPath = $parentDN
                            }
                        }
                    }
                }
            } catch {
                # Fehler beim Ermitteln des Parent, Standardwert bleibt
                Write-DebugLog "Could not determine parent for $($OU.Name). Error: $($_.Exception.Message)" -Category "OUHierarchy"
            }
            
            $DirectChildrenCount = 0
            try {
                 $DirectChildrenCount = (Get-ADObject -Filter * -SearchBase $OU.DistinguishedName -SearchScope OneLevel -ErrorAction SilentlyContinue).Count
            } catch {
                Write-ADReportLog -Message "Could not determine children count for $($OU.Name). Error: $($_.Exception.Message)" -Type Warning
            }

            [PSCustomObject]@{
                Name                        = $OU.Name
                DistinguishedName           = $OU.DistinguishedName
                ParentContainer             = $ParentPath
                Description                 = $OU.Description
                ProtectedFromDeletion       = $OU.ProtectedFromAccidentalDeletion
                Created                     = $OU.whenCreated
                Modified                    = $OU.whenChanged
                CanonicalName               = $OU.CanonicalName
            }
        }
        Write-ADReportLog -Message "Successfully generated OU Hierarchy Report for $($OUReport.Count) OUs." -Type Info -Terminal
        return $OUReport | Sort-Object DistinguishedName
    }
    catch {
        $ErrorMessage = "Error generating OU Hierarchy Report: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -Terminal
        return $null
    }
}

# --- Funktion zum Abrufen von AD Sites und Subnets ---
Function Get-ADSitesAndSubnetsReport {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Gathering AD Sites and Subnets information..." -Type Info -Terminal
    Initialize-ResultCounters

    try {
        $Report = @()

        # Sites abrufen
        $Sites = Get-ADReplicationSite -Filter * -Properties Description, Options, InterSiteTopologyGenerator -ErrorAction SilentlyContinue
        if ($Sites) {
            foreach ($Site in $Sites) {
                # Server fÃ¼r diesen Site separat abrufen
                $ServersCount = 0
                try {
                    $Servers = Get-ADDomainController -Filter { Site -eq $Site.Name } -ErrorAction SilentlyContinue
                    $ServersCount = if ($Servers) { @($Servers).Count } else { 0 }
                } catch {
                    Write-ADReportLog -Message "Could not get servers for site $($Site.Name): $($_.Exception.Message)" -Type Warning
                }

                # Site Links separat abrufen
                $SiteLinksText = ""
                try {
                    $SiteLinks = Get-ADReplicationSiteLink -Filter "(Sites -eq '$($Site.DistinguishedName)')" -ErrorAction SilentlyContinue
                    $SiteLinksText = if ($SiteLinks) { ($SiteLinks | ForEach-Object { $_.Name }) -join ", " }
                } catch {
                    Write-ADReportLog -Message "Could not get site links for site $($Site.Name): $($_.Exception.Message)" -Type Warning
                }

                $Report += [PSCustomObject]@{
                    Type = "Site"
                    Name = $Site.Name
                    DistinguishedName = $Site.DistinguishedName
                    Description = $Site.Description
                    ServersInSiteCount = $ServersCount
                    InterSiteTopologyGenerator = $Site.InterSiteTopologyGenerator
                    Options = $Site.Options
                    SiteLinks = $SiteLinksText
                    Location = $null # FÃ¼r Konsistenz
                    AssociatedSite = $null # FÃ¼r Konsistenz mit Subnets
                }
            }
            Write-ADReportLog -Message "Successfully retrieved $($Sites.Count) AD Replication Sites." -Type Info -Terminal
        } else {
            Write-ADReportLog -Message "No AD Replication Sites found." -Type Warning -Terminal
        }

        # Subnets abrufen
        $Subnets = Get-ADReplicationSubnet -Filter * -Properties Description, Location, Site -ErrorAction SilentlyContinue
        if ($Subnets) {
            foreach ($Subnet in $Subnets) {
                $Report += [PSCustomObject]@{
                    Type = "Subnet"
                    Name = $Subnet.Name
                    DistinguishedName = $Subnet.DistinguishedName
                    Description = $Subnet.Description
                    ServersInSiteCount = $null 
                    InterSiteTopologyGenerator = $null
                    Options = $null
                    SiteLinks = $null # FÃ¼r Konsistenz
                    Location = $Subnet.Location
                    AssociatedSite = try { (Get-ADReplicationSite -Identity $Subnet.Site -ErrorAction SilentlyContinue).Name } catch { $Subnet.Site }
                }
            }
            Write-ADReportLog -Message "Successfully retrieved $($Subnets.Count) AD Replication Subnets." -Type Info -Terminal
        } else {
            Write-ADReportLog -Message "No AD Replication Subnets found." -Type Warning -Terminal
        }

        if ($Report.Count -gt 0) {
            Write-ADReportLog -Message "Successfully generated Sites and Subnets Report for $($Report.Count) entries." -Type Info
            return $Report | Sort-Object Type, Name
        } else {
            # Wenn keine Daten gefunden wurden, einen Platzhalter-Eintrag als Array zurÃ¼ckgeben
            Write-ADReportLog -Message "No data found for Sites and Subnets Report." -Type Info
            return @(
                [PSCustomObject]@{
                    Type = "Information"
                    Name = "No Data"
                    DistinguishedName = "N/A"
                    Description = "No Sites or Subnets found in Active Directory"
                    ServersInSiteCount = 0
                    InterSiteTopologyGenerator = "N/A"
                    Options = "N/A"
                    SiteLinks = "N/A"
                    Location = "N/A"
                    AssociatedSite = "N/A"
                }
            )
        }
    }
    catch {
        $ErrorMessage = "Error generating Sites and Subnets Report: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -Terminal
        return @(
            [PSCustomObject]@{
                Type = "Error"
                Name = "System Error"
                DistinguishedName = "N/A"
                Description = $ErrorMessage
                ServersInSiteCount = 0
                InterSiteTopologyGenerator = "N/A"
                Options = "N/A"
                SiteLinks = "N/A"
                Location = "N/A"
                AssociatedSite = "N/A"
            }
        )
    }
}

# Funktion zum Aktualisieren der ErgebniszÃ¤hler im Header
Function Initialize-ResultCounters {
    # GesamtergebniszÃ¤hler auf 0 setzen
    if ($null -ne $Global:TotalResultCountText) {
        $Global:TotalResultCountText.Text = "0"
    }
    
    # Sicherstellen, dass alle ZÃ¤hler zurÃ¼ckgesetzt werden
    if ($null -ne $Global:UserCountText) {
        $Global:UserCountText.Text = "0"
    }
    
    if ($null -ne $Global:ComputerCountText) {
        $Global:ComputerCountText.Text = "0"
    }
    
    if ($null -ne $Global:GroupCountText) {
        $Global:GroupCountText.Text = "0"
    }
    
    # Status zurÃ¼cksetzen
    if ($null -ne $Global:TextBlockStatus) {
        $Global:TextBlockStatus.Text = "Ready for query..."
    }
    
    # DataGrid leeren
    if ($null -ne $Global:DataGridResults) {
        $Global:DataGridResults.ItemsSource = $null
    }
}

# Funktion zum Aktualisieren der ErgebniszÃ¤hler im Header
Function Update-ResultCounters {
    param (
        [Parameter(Mandatory=$true)]
        [System.Collections.IList]$Results
    )

    if ($null -eq $Results) {
        if ($null -ne $Global:TotalResultCountText) { $Global:TotalResultCountText.Text = "0" }
        if ($null -ne $Global:UserCountText) { $Global:UserCountText.Text = "0" }
        if ($null -ne $Global:ComputerCountText) { $Global:ComputerCountText.Text = "0" }
        if ($null -ne $Global:GroupCountText) { $Global:GroupCountText.Text = "0" }
        return
    }

    $totalCount = $Results.Count
    $userCount = 0
    $computerCount = 0
    $groupCount = 0

    if ($totalCount -gt 0) {
        foreach ($item in $Results) {
            if ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [array] -and $item.ObjectClass -contains "user") {
                $userCount++
            } elseif ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [string] -and $item.ObjectClass -eq "user") {
                $userCount++
            } elseif ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [array] -and $item.ObjectClass -contains "computer") {
                $computerCount++
            } elseif ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [string] -and $item.ObjectClass -eq "computer") {
                $computerCount++
            } elseif ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [array] -and $item.ObjectClass -contains "group") {
                $groupCount++
            } elseif ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [string] -and $item.ObjectClass -eq "group") {
                $groupCount++
            } elseif ($item.PSObject.Properties["ObjectCategory"] -and $item.ObjectCategory -like "*CN=Person,CN=Schema,CN=Configuration,*") {
                $userCount++
            } elseif ($item.PSObject.Properties["ObjectCategory"] -and $item.ObjectCategory -like "*CN=Computer,CN=Schema,CN=Configuration,*") {
                $computerCount++
            } elseif ($item.PSObject.Properties["ObjectCategory"] -and $item.ObjectCategory -like "*CN=Group,CN=Schema,CN=Configuration,*") {
                $groupCount++
            }
        }
    }

    if ($null -ne $Global:TotalResultCountText) { $Global:TotalResultCountText.Text = $totalCount.ToString() }
    if ($null -ne $Global:UserCountText) { $Global:UserCountText.Text = $userCount.ToString() }
    if ($null -ne $Global:ComputerCountText) { $Global:ComputerCountText.Text = $computerCount.ToString() }
    if ($null -ne $Global:GroupCountText) { $Global:GroupCountText.Text = $groupCount.ToString() }
}

# Funktion zur universellen Aktualisierung der Ergebnisanzeige und DataGrid
Function Update-ADReportResults {
    param (
        [Parameter(Mandatory=$false)]
        $Results = $null,
        
        [Parameter(Mandatory=$false)]
        [string]$StatusMessage
    )

    # Stelle sicher, dass Results immer ein Array ist, auch bei NULL
    if ($null -eq $Results) {
        $Results = @()
    } elseif ($Results -isnot [array]) {
        $Results = @($Results)
    }

    if ($null -ne $Global:DataGridResults) {
        $Global:DataGridResults.ItemsSource = $Results
    }

    if (-not [string]::IsNullOrWhiteSpace($StatusMessage) -and $null -ne $Global:TextBlockStatus) {
        $Global:TextBlockStatus.Text = $StatusMessage
    }
    
    Update-ResultCounters -Results $Results
    Update-ResultVisualization -Results $Results
}

# Hilfsfunktion zum sicheren HinzufÃ¼gen von Event-Handlern
Function Add-SafeEventHandler {
    param(
        [Parameter(Mandatory=$true)]
        $Button,
        [Parameter(Mandatory=$true)]
        [scriptblock]$Handler
    )
    
    if ($null -ne $Button) {
        try {
            $Button.add_Click($Handler)
        } catch {
            Write-ADReportLog -Message "Warnung: Konnte Event-Handler fÃ¼r Button nicht hinzufÃ¼gen: $($_.Exception.Message)" -Type Warning -Terminal
        }
   }
}

# --- Globale Variablen fÃ¼r UI Elemente --- 
Function Initialize-ADReportForm {
    param($XAMLContent)
    # ÃœberprÃ¼fen, ob das Window-Objekt bereits existiert und zurÃ¼cksetzen
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
    $Global:ComboBoxFilterAttribute1 = $Window.FindName("ComboBoxFilterAttribute1")
    $Global:ComboBoxFilterOperator1 = $Window.FindName("ComboBoxFilterOperator1")
    $Global:TextBoxFilterValue1 = $Window.FindName("TextBoxFilterValue1")
    $Global:ComboBoxFilterAttribute2 = $Window.FindName("ComboBoxFilterAttribute2")  
    $Global:ComboBoxFilterOperator2 = $Window.FindName("ComboBoxFilterOperator2")
    $Global:TextBoxFilterValue2 = $Window.FindName("TextBoxFilterValue2")
    $Global:RadioButtonAnd = $Window.FindName("RadioButtonAnd")
    $Global:RadioButtonOr = $Window.FindName("RadioButtonOr")
    $Global:CheckBoxUseSecondFilter = $Window.FindName("CheckBoxUseSecondFilter")
    $Global:SecondFilterPanel = $Window.FindName("SecondFilterPanel")

    # Bestehende UI-Elemente
    $Global:RadioButtonUser = $Window.FindName("RadioButtonUser")
    $Global:RadioButtonGroup = $Window.FindName("RadioButtonGroup")
    $Global:RadioButtonComputer = $Window.FindName("RadioButtonComputer")
    $Global:RadioButtonGroupMemberships = $Window.FindName("RadioButtonGroupMemberships")
    $Global:ListBoxSelectAttributes = $Window.FindName("ListBoxSelectAttributes")
    $Global:ButtonQueryAD = $Window.FindName("ButtonQueryAD")
    $Global:DataGridResults = $Window.FindName("DataGridResults")
    $Global:TextBlockStatus = $Window.FindName("TextBlockStatus")
    $Global:TotalResultCountText = $Window.FindName("TotalResultCountText")
    $Global:ButtonExportCSV = $Window.FindName("ButtonExportCSV")
    $Global:ButtonExportHTML = $Window.FindName("ButtonExportHTML")

    # Quick Report Buttons
    $Global:ButtonQuickAllUsers = $Window.FindName("ButtonQuickAllUsers")
    $Global:ButtonQuickDisabledUsers = $Window.FindName("ButtonQuickDisabledUsers")
    $Global:ButtonQuickLockedUsers = $Window.FindName("ButtonQuickLockedUsers")
    $Global:ButtonQuickNeverExpire = $Window.FindName("ButtonQuickNeverExpire")
    $Global:ButtonQuickInactiveUsers = $Window.FindName("ButtonQuickInactiveUsers")
    $Global:ButtonQuickAdminUsers = $Window.FindName("ButtonQuickAdminUsers")
    $Global:ButtonQuickRecentlyCreatedUsers = $Window.FindName("ButtonQuickRecentlyCreatedUsers")
    $Global:ButtonQuickPasswordExpiringSoon = $Window.FindName("ButtonQuickPasswordExpiringSoon")
    $Global:ButtonQuickExpiredPasswords = $Window.FindName("ButtonQuickExpiredPasswords")
    $Global:ButtonQuickNeverLoggedOn = $Window.FindName("ButtonQuickNeverLoggedOn")
    $Global:ButtonQuickRecentlyDeletedUsers = $Window.FindName("ButtonQuickRecentlyDeletedUsers")
    $Global:ButtonQuickRecentlyModifiedUsers = $Window.FindName("ButtonQuickRecentlyModifiedUsers")
    $Global:ButtonQuickInactiveUsersXDays = $Window.FindName("ButtonQuickInactiveUsersXDays")
    $Global:ButtonQuickUsersWithoutManager = $Window.FindName("ButtonQuickUsersWithoutManager")
    $Global:ButtonQuickUsersMissingRequiredAttributes = $Window.FindName("ButtonQuickUsersMissingRequiredAttributes")
    $Global:ButtonQuickUsersDuplicateLogonNames = $Window.FindName("ButtonQuickUsersDuplicateLogonNames")
    $Global:ButtonQuickOrphanedSIDsUsers = $Window.FindName("ButtonQuickOrphanedSIDsUsers")
    
    # Gruppe und Computer Buttons
    $Global:ButtonQuickGroups = $Window.FindName("ButtonQuickGroups")
    $Global:ButtonQuickSecurityGroups = $Window.FindName("ButtonQuickSecurityGroups")
    $Global:ButtonQuickDistributionGroups = $Window.FindName("ButtonQuickDistributionGroups")
    $Global:ButtonQuickComputers = $Window.FindName("ButtonQuickComputers")
    $Global:ButtonQuickInactiveComputers = $Window.FindName("ButtonQuickInactiveComputers")
    
    # Security Audit Buttons
    $Global:ButtonQuickWeakPasswordPolicy = $Window.FindName("ButtonQuickWeakPasswordPolicy")
    $Global:ButtonQuickRiskyGroupMemberships = $Window.FindName("ButtonQuickRiskyGroupMemberships")
    $Global:ButtonQuickPrivilegedAccounts = $Window.FindName("ButtonQuickPrivilegedAccounts")
    
    # AD-Health Buttons
    $Global:ButtonQuickFSMORoles = $Window.FindName("ButtonQuickFSMORoles")
    $Global:ButtonQuickDCStatus = $Window.FindName("ButtonQuickDCStatus")
    $Global:ButtonQuickReplicationStatus = $Window.FindName("ButtonQuickReplicationStatus")
    $Global:ButtonQuickOUHierarchy = $Window.FindName("ButtonQuickOUHierarchy")
    $Global:ButtonQuickSitesSubnets = $Window.FindName("ButtonQuickSitesSubnets")
    
    # Neue Buttons fÃ¼r erweiterte Reports
    $Global:ButtonQuickDepartmentStats = $Window.FindName("ButtonQuickDepartmentStats")
    $Global:ButtonQuickDepartmentSecurity = $Window.FindName("ButtonQuickDepartmentSecurity")
    $Global:ButtonQuickKerberoastable = $Window.FindName("ButtonQuickKerberoastable")
    $Global:ButtonQuickASREPRoastable = $Window.FindName("ButtonQuickASREPRoastable")
    $Global:ButtonQuickDelegation = $Window.FindName("ButtonQuickDelegation")
    $Global:ButtonQuickDCSyncRights = $Window.FindName("ButtonQuickDCSyncRights")
    $Global:ButtonQuickSchemaAdmins = $Window.FindName("ButtonQuickSchemaAdmins")
    $Global:ButtonQuickCertificateAnalysis = $Window.FindName("ButtonQuickCertificateAnalysis")
    $Global:ButtonQuickSYSVOLHealth = $Window.FindName("ButtonQuickSYSVOLHealth")
    $Global:ButtonQuickDNSHealth = $Window.FindName("ButtonQuickDNSHealth")
    $Global:ButtonQuickBackupStatus = $Window.FindName("ButtonQuickBackupStatus")
    $Global:ButtonQuickSchemaAnalysis = $Window.FindName("ButtonQuickSchemaAnalysis")
    $Global:ButtonQuickTrustRelationships = $Window.FindName("ButtonQuickTrustRelationships")
    $Global:ButtonQuickQuotasLimits = $Window.FindName("ButtonQuickQuotasLimits")

    # Neue User Report Buttons aus Roadmap
    $Global:ButtonQuickStalePasswords = $Window.FindName("ButtonQuickStalePasswords")
    $Global:ButtonQuickNeverChangingPasswords = $Window.FindName("ButtonQuickNeverChangingPasswords")
    $Global:ButtonQuickExpiringAccounts = $Window.FindName("ButtonQuickExpiringAccounts")
    $Global:ButtonQuickReversibleEncryption = $Window.FindName("ButtonQuickReversibleEncryption")
    $Global:ButtonQuickKerberosDES = $Window.FindName("ButtonQuickKerberosDES")
    $Global:ButtonQuickUsersWithSPN = $Window.FindName("ButtonQuickUsersWithSPN")
    $Global:ButtonQuickGuestAccountStatus = $Window.FindName("ButtonQuickGuestAccountStatus")
    $Global:ButtonQuickUsersByDepartment = $Window.FindName("ButtonQuickUsersByDepartment")
    $Global:ButtonQuickUsersByManager = $Window.FindName("ButtonQuickUsersByManager")
    $Global:ButtonQuickRemoteAccessUsers = $Window.FindName("ButtonQuickRemoteAccessUsers")
    $Global:ButtonQuickMobileDeviceUsers = $Window.FindName("ButtonQuickMobileDeviceUsers")
    
    # Neue Group Report Buttons aus Roadmap
    $Global:ButtonQuickEmptyGroups = $Window.FindName("ButtonQuickEmptyGroups")
    $Global:ButtonQuickNestedGroups = $Window.FindName("ButtonQuickNestedGroups")
    $Global:ButtonQuickCircularGroups = $Window.FindName("ButtonQuickCircularGroups")
    $Global:ButtonQuickGroupsByTypeScope = $Window.FindName("ButtonQuickGroupsByTypeScope")
    $Global:ButtonQuickDynamicDistGroups = $Window.FindName("ButtonQuickDynamicDistGroups")
    $Global:ButtonQuickMailEnabledGroups = $Window.FindName("ButtonQuickMailEnabledGroups")
    $Global:ButtonQuickGroupsWithoutOwners = $Window.FindName("ButtonQuickGroupsWithoutOwners")
    $Global:ButtonQuickLargeGroups = $Window.FindName("ButtonQuickLargeGroups")
    $Global:ButtonQuickRecentlyModifiedGroups = $Window.FindName("ButtonQuickRecentlyModifiedGroups")
    
    # Neue Computer Report Buttons aus Roadmap
    $Global:ButtonQuickOSSummary = $Window.FindName("ButtonQuickOSSummary")
    $Global:ButtonQuickComputersByOSVersion = $Window.FindName("ButtonQuickComputersByOSVersion")
    $Global:ButtonQuickBitLockerStatus = $Window.FindName("ButtonQuickBitLockerStatus")
    $Global:ButtonQuickTPMStatus = $Window.FindName("ButtonQuickTPMStatus")
    $Global:ButtonQuickStaleComputerPasswords = $Window.FindName("ButtonQuickStaleComputerPasswords")
    $Global:ButtonQuickComputersNeverLoggedOn = $Window.FindName("ButtonQuickComputersNeverLoggedOn")
    $Global:ButtonQuickDuplicateComputerNames = $Window.FindName("ButtonQuickDuplicateComputerNames")
    $Global:ButtonQuickComputersByLocation = $Window.FindName("ButtonQuickComputersByLocation")
    $Global:ButtonQuickVirtualVsPhysical = $Window.FindName("ButtonQuickVirtualVsPhysical")
    $Global:ButtonQuickDomainControllerHealth = $Window.FindName("ButtonQuickDomainControllerHealth")
    
    # Neue Security Audit Buttons aus Roadmap
    $Global:ButtonQuickHoneyTokens = $Window.FindName("ButtonQuickHoneyTokens")
    $Global:ButtonQuickPrivilegeEscalation = $Window.FindName("ButtonQuickPrivilegeEscalation")
    $Global:ButtonQuickExposedCredentials = $Window.FindName("ButtonQuickExposedCredentials")
    $Global:ButtonQuickSuspiciousLogons = $Window.FindName("ButtonQuickSuspiciousLogons")
    $Global:ButtonQuickForeignSecurityPrincipals = $Window.FindName("ButtonQuickForeignSecurityPrincipals")
    $Global:ButtonQuickSIDHistoryAbuse = $Window.FindName("ButtonQuickSIDHistoryAbuse")
    
    # Service Account Buttons
    $Global:ButtonQuickServiceAccountsOverview = $Window.FindName("ButtonQuickServiceAccountsOverview")
    $Global:ButtonQuickManagedServiceAccounts = $Window.FindName("ButtonQuickManagedServiceAccounts")
    $Global:ButtonQuickServiceAccountsSPN = $Window.FindName("ButtonQuickServiceAccountsSPN")
    $Global:ButtonQuickHighPrivServiceAccounts = $Window.FindName("ButtonQuickHighPrivServiceAccounts")
    $Global:ButtonQuickServiceAccountPasswordAge = $Window.FindName("ButtonQuickServiceAccountPasswordAge")
    $Global:ButtonQuickUnusedServiceAccounts = $Window.FindName("ButtonQuickUnusedServiceAccounts")
    
    # GPO & Policy Buttons
    $Global:ButtonQuickGPOOverview = $Window.FindName("ButtonQuickGPOOverview")
    $Global:ButtonQuickUnlinkedGPOs = $Window.FindName("ButtonQuickUnlinkedGPOs")
    $Global:ButtonQuickEmptyGPOs = $Window.FindName("ButtonQuickEmptyGPOs")
    $Global:ButtonQuickGPOPermissions = $Window.FindName("ButtonQuickGPOPermissions")
    $Global:ButtonQuickPasswordPolicySummary = $Window.FindName("ButtonQuickPasswordPolicySummary")
    $Global:ButtonQuickAccountLockoutPolicies = $Window.FindName("ButtonQuickAccountLockoutPolicies")
    $Global:ButtonQuickFineGrainedPasswordPolicies = $Window.FindName("ButtonQuickFineGrainedPasswordPolicies")
    
    # Compliance & Audit Buttons
    $Global:ButtonQuickGDPRCompliance = $Window.FindName("ButtonQuickGDPRCompliance")
    $Global:ButtonQuickISO27001Readiness = $Window.FindName("ButtonQuickISO27001Readiness")
    $Global:ButtonQuickNISTFramework = $Window.FindName("ButtonQuickNISTFramework")
    $Global:ButtonQuickPCIDSSRequirements = $Window.FindName("ButtonQuickPCIDSSRequirements")
    $Global:ButtonQuickHIPAACompliance = $Window.FindName("ButtonQuickHIPAACompliance")
    $Global:ButtonQuickCustomComplianceTemplates = $Window.FindName("ButtonQuickCustomComplianceTemplates")
    
    # Advanced Permissions Buttons
    $Global:ButtonQuickACLAnalysis = $Window.FindName("ButtonQuickACLAnalysis")
    $Global:ButtonQuickInheritanceBreaks = $Window.FindName("ButtonQuickInheritanceBreaks")
    $Global:ButtonQuickAdminSDHolderObjects = $Window.FindName("ButtonQuickAdminSDHolderObjects")
    $Global:ButtonQuickAdvancedDelegation = $Window.FindName("ButtonQuickAdvancedDelegation")
    $Global:ButtonQuickSchemaPermissions = $Window.FindName("ButtonQuickSchemaPermissions")
    
    # Attribute selection buttons
    $Global:ButtonSelectAllAttributes = $Window.FindName("ButtonSelectAllAttributes")
    $Global:ButtonSelectNoneAttributes = $Window.FindName("ButtonSelectNoneAttributes")
    $Global:TabControlAttributes = $Window.FindName("TabControlAttributes")
    $Global:ListBoxBasicAttributes = $Window.FindName("ListBoxBasicAttributes")
    $Global:ListBoxSecurityAttributes = $Window.FindName("ListBoxSecurityAttributes")
    $Global:ListBoxExtendedAttributes = $Window.FindName("ListBoxExtendedAttributes")
    
    # Help and About buttons
    $Global:ButtonHelp = $Window.FindName("ButtonHelp")
    $Global:ButtonAbout = $Window.FindName("ButtonAbout")
    
    # Footer elements
    $Global:StatusIndicator = $Window.FindName("StatusIndicator")
    $Global:TextBlockSelectedRows = $Window.FindName("TextBlockSelectedRows")
    $Global:TextBlockLastUpdate = $Window.FindName("TextBlockLastUpdate")

    # Event Handler fÃ¼r erweiterte Filter
    $Global:CheckBoxUseSecondFilter.add_Checked({
        $Global:SecondFilterPanel.IsEnabled = $true
    })
    
    $Global:CheckBoxUseSecondFilter.add_Unchecked({
        $Global:SecondFilterPanel.IsEnabled = $false
    })

    # Helper function to get all selected attributes from all ListBoxes
    function Get-AllSelectedAttributes {
        $selectedAttributes = @()
        
        if ($Global:ListBoxBasicAttributes) {
            foreach ($item in $Global:ListBoxBasicAttributes.SelectedItems) {
                $selectedAttributes += $item.Content
            }
        }
        if ($Global:ListBoxSecurityAttributes) {
            foreach ($item in $Global:ListBoxSecurityAttributes.SelectedItems) {
                $selectedAttributes += $item.Content
            }
        }
        if ($Global:ListBoxExtendedAttributes) {
            foreach ($item in $Global:ListBoxExtendedAttributes.SelectedItems) {
                $selectedAttributes += $item.Content
            }
        }
        
        return $selectedAttributes | Select-Object -Unique
    }

    # Helper function to select specific attributes in the tabbed ListBoxes
    function Select-AttributesInListBoxes {
        param (
            [string[]]$Attributes
        )
        
        # Clear all selections
        if ($Global:ListBoxBasicAttributes) {
            foreach ($item in $Global:ListBoxBasicAttributes.Items) {
                $item.IsSelected = $false
            }
        }
        if ($Global:ListBoxSecurityAttributes) {
            foreach ($item in $Global:ListBoxSecurityAttributes.Items) {
                $item.IsSelected = $false
            }
        }
        if ($Global:ListBoxExtendedAttributes) {
            foreach ($item in $Global:ListBoxExtendedAttributes.Items) {
                $item.IsSelected = $false
            }
        }
        
        # Select requested attributes
        foreach ($attr in $Attributes) {
            # Check in Basic Attributes
            if ($Global:ListBoxBasicAttributes) {
                foreach ($item in $Global:ListBoxBasicAttributes.Items) {
                    if ($item.Content -eq $attr) {
                        $item.IsSelected = $true
                        break
                    }
                }
            }
            # Check in Security Attributes
            if ($Global:ListBoxSecurityAttributes) {
                foreach ($item in $Global:ListBoxSecurityAttributes.Items) {
                    if ($item.Content -eq $attr) {
                        $item.IsSelected = $true
                        break
                    }
                }
            }
            # Check in Extended Attributes
            if ($Global:ListBoxExtendedAttributes) {
                foreach ($item in $Global:ListBoxExtendedAttributes.Items) {
                    if ($item.Content -eq $attr) {
                        $item.IsSelected = $true
                        break
                    }
                }
            }
        }
    }

    # Helper function to populate attributes in the tabbed ListBoxes
    function Set-AttributesForObjectType {
        param (
            [string]$ObjectType
        )
        
        # Clear all ListBoxes
        if ($Global:ListBoxBasicAttributes) { $Global:ListBoxBasicAttributes.Items.Clear() }
        if ($Global:ListBoxSecurityAttributes) { $Global:ListBoxSecurityAttributes.Items.Clear() }
        if ($Global:ListBoxExtendedAttributes) { $Global:ListBoxExtendedAttributes.Items.Clear() }
        
        switch ($ObjectType) {
            "User" {
                # Basic Attributes
                @("DisplayName", "SamAccountName", "GivenName", "Surname", "mail", "Department", "Title", "Enabled") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    if ($_ -in @("DisplayName", "SamAccountName", "Enabled")) { $item.IsSelected = $true }
                    $Global:ListBoxBasicAttributes.Items.Add($item)
                }
                
                # Security Attributes
                @("LastLogonTimestamp", "PasswordExpired", "PasswordLastSet", "AccountExpirationDate", "badPwdCount", "lockoutTime", "UserAccountControl", "memberOf") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    $Global:ListBoxSecurityAttributes.Items.Add($item)
                }
                
                # Extended Attributes
                @("whenCreated", "whenChanged", "Manager", "Company", "physicalDeliveryOfficeName", "telephoneNumber", "homeDirectory", "scriptPath") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    $Global:ListBoxExtendedAttributes.Items.Add($item)
                }
            }
            "Group" {
                # Basic Attributes
                @("Name", "SamAccountName", "Description", "GroupCategory", "GroupScope", "mail") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    if ($_ -in @("Name", "SamAccountName")) { $item.IsSelected = $true }
                    $Global:ListBoxBasicAttributes.Items.Add($item)
                }
                
                # Security Attributes
                @("ManagedBy", "info", "memberOf", "members") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    $Global:ListBoxSecurityAttributes.Items.Add($item)
                }
                
                # Extended Attributes
                @("whenCreated", "whenChanged", "distinguishedName", "objectGUID") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    $Global:ListBoxExtendedAttributes.Items.Add($item)
                }
            }
            "Computer" {
                # Basic Attributes
                @("Name", "DNSHostName", "OperatingSystem", "OperatingSystemVersion", "Enabled", "IPv4Address") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    if ($_ -in @("Name", "OperatingSystem", "Enabled")) { $item.IsSelected = $true }
                    $Global:ListBoxBasicAttributes.Items.Add($item)
                }
                
                # Security Attributes
                @("LastLogonDate", "PasswordLastSet", "userAccountControl") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    $Global:ListBoxSecurityAttributes.Items.Add($item)
                }
                
                # Extended Attributes
                @("whenCreated", "Description", "Location", "ManagedBy", "servicePrincipalName") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    $Global:ListBoxExtendedAttributes.Items.Add($item)
                }
            }
        }
    }

    # RadioButton Event Handler
    $RadioButtonUser.add_Checked({
        Write-ADReportLog -Message "Object type changed to User" -Type Info -Terminal
        
        # Filter-Attribute fÃ¼r Benutzer
        $UserFilterAttributes = @("DisplayName", "SamAccountName", "GivenName", "Surname", "mail", "Department", "Title", "EmployeeID", "UserPrincipalName")
        $Global:ComboBoxFilterAttribute1.Items.Clear()
        $Global:ComboBoxFilterAttribute2.Items.Clear()
        $UserFilterAttributes | ForEach-Object { 
            $Global:ComboBoxFilterAttribute1.Items.Add($_)
            $Global:ComboBoxFilterAttribute2.Items.Add($_)
        }
        if ($Global:ComboBoxFilterAttribute1.Items.Count -gt 0) { 
            $Global:ComboBoxFilterAttribute1.SelectedIndex = 0 
            $Global:ComboBoxFilterAttribute2.SelectedIndex = 1
        }
        
        # Populate attributes in tabbed ListBoxes
        Set-AttributesForObjectType -ObjectType "User"
        
        $Global:TextBlockStatus.Text = "Ready for user query"
    })

    $RadioButtonGroup.add_Checked({
        Write-ADReportLog -Message "Object type changed to Group" -Type Info -Terminal
        
        # Filter-Attribute fÃ¼r Gruppen
        $GroupFilterAttributes = @("Name", "SamAccountName", "Description", "GroupCategory", "GroupScope")
        $Global:ComboBoxFilterAttribute1.Items.Clear()
        $Global:ComboBoxFilterAttribute2.Items.Clear()
        $GroupFilterAttributes | ForEach-Object { 
            $Global:ComboBoxFilterAttribute1.Items.Add($_)
            $Global:ComboBoxFilterAttribute2.Items.Add($_)
        }
        if ($Global:ComboBoxFilterAttribute1.Items.Count -gt 0) { 
            $Global:ComboBoxFilterAttribute1.SelectedIndex = 0 
            $Global:ComboBoxFilterAttribute2.SelectedIndex = 1
        }
        
        # Populate attributes in tabbed ListBoxes
        Set-AttributesForObjectType -ObjectType "Group"
        
        $Global:TextBlockStatus.Text = "Ready for group query"
    })

    $RadioButtonComputer.add_Checked({
        Write-ADReportLog -Message "Object type changed to Computer" -Type Info -Terminal
        
        # Filter-Attribute fÃ¼r Computer
        $ComputerFilterAttributes = @("Name", "DNSHostName", "OperatingSystem", "Description")
        $Global:ComboBoxFilterAttribute1.Items.Clear()
        $Global:ComboBoxFilterAttribute2.Items.Clear()
        $ComputerFilterAttributes | ForEach-Object { 
            $Global:ComboBoxFilterAttribute1.Items.Add($_)
            $Global:ComboBoxFilterAttribute2.Items.Add($_)
        }
        if ($Global:ComboBoxFilterAttribute1.Items.Count -gt 0) { 
            $Global:ComboBoxFilterAttribute1.SelectedIndex = 0 
            $Global:ComboBoxFilterAttribute2.SelectedIndex = 1
        }
        
        # Populate attributes in tabbed ListBoxes
        Set-AttributesForObjectType -ObjectType "Computer"
        
        $Global:TextBlockStatus.Text = "Ready for computer query"
    })

    $RadioButtonGroupMemberships.add_Checked({
        Write-ADReportLog -Message "Object type changed to GroupMemberships" -Type Info -Terminal
        # Clear all ListBoxes for Group Memberships (no attribute selection needed)
        if ($Global:ListBoxBasicAttributes) { $Global:ListBoxBasicAttributes.Items.Clear() }
        if ($Global:ListBoxSecurityAttributes) { $Global:ListBoxSecurityAttributes.Items.Clear() }
        if ($Global:ListBoxExtendedAttributes) { $Global:ListBoxExtendedAttributes.Items.Clear() }
        Write-ADReportLog -Message "Attribute selection disabled for GroupMemberships query." -Type Info
    })

    # Event Handler fÃ¼r ButtonQueryAD anpassen, um Objekttyp zu berÃ¼cksichtigen
    $ButtonQueryAD.add_Click({
        Write-ADReportLog -Message "Executing query..." -Type Info
        try {
            # Hole Filter-Werte
            $SelectedFilterAttribute = if ($Global:ComboBoxFilterAttribute1.SelectedItem) { $Global:ComboBoxFilterAttribute1.SelectedItem.ToString() } else { "" }
            $FilterValue = $Global:TextBoxFilterValue1.Text
            $FilterOperator = if ($Global:ComboBoxFilterOperator1.SelectedItem) { $Global:ComboBoxFilterOperator1.SelectedItem.Content.ToString() } else { "Contains" }
            
            # Zweiter Filter (wenn aktiviert)
            $UseSecondFilter = $Global:CheckBoxUseSecondFilter.IsChecked
            $SelectedFilterAttribute2 = ""
            $FilterValue2 = ""
            $FilterOperator2 = "Contains"
            $FilterLogic = if ($Global:RadioButtonAnd.IsChecked) { "AND" } else { "OR" }
            
            if ($UseSecondFilter) {
                $SelectedFilterAttribute2 = if ($Global:ComboBoxFilterAttribute2.SelectedItem) { $Global:ComboBoxFilterAttribute2.SelectedItem.ToString() } else { "" }
                $FilterValue2 = $Global:TextBoxFilterValue2.Text
                $FilterOperator2 = if ($Global:ComboBoxFilterOperator2.SelectedItem) { $Global:ComboBoxFilterOperator2.SelectedItem.Content.ToString() } else { "Contains" }
            }
            
            # Get selected attributes from all three ListBoxes
            $SelectedAttributes = Get-AllSelectedAttributes
            Write-Host "DEBUG: Selected attributes: $($SelectedAttributes -join '; ')"
            $isUserSearch = $Global:RadioButtonUser.IsChecked

            if ($SelectedAttributes.Count -eq 0 -and $Global:RadioButtonGroupMemberships.IsChecked -eq $false) {
                [System.Windows.MessageBox]::Show("Please select at least one attribute for export.", "Warnung", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
                return
            }

            # Bestimme den aktuell ausgewÃ¤hlten Objekttyp
            $ObjectType = if ($Global:RadioButtonUser.IsChecked) { "User" } 
                        elseif ($Global:RadioButtonGroup.IsChecked) { "Group" } 
                        elseif ($Global:RadioButtonGroupMemberships.IsChecked) { "GroupMemberships" } 
                        else { "Computer" }
            
            # AD-Abfrage basierend auf Objekttyp durchfÃ¼hren
            $ReportData = $null
            switch ($ObjectType) {
                "User" {
                    $ReportData = Get-ADReportData -FilterAttribute $SelectedFilterAttribute -FilterValue $FilterValue -FilterOperator $FilterOperator `
                                                  -FilterAttribute2 $SelectedFilterAttribute2 -FilterValue2 $FilterValue2 -FilterOperator2 $FilterOperator2 `
                                                  -FilterLogic $FilterLogic -SelectedAttributes $SelectedAttributes -ObjectType "User"
                }
                "Group" {
                    $ReportData = Get-ADReportData -FilterAttribute $SelectedFilterAttribute -FilterValue $FilterValue -FilterOperator $FilterOperator `
                                                  -FilterAttribute2 $SelectedFilterAttribute2 -FilterValue2 $FilterValue2 -FilterOperator2 $FilterOperator2 `
                                                  -FilterLogic $FilterLogic -SelectedAttributes $SelectedAttributes -ObjectType "Group"
                }
                "Computer" {
                    $ReportData = Get-ADReportData -FilterAttribute $SelectedFilterAttribute -FilterValue $FilterValue -FilterOperator $FilterOperator `
                                                  -FilterAttribute2 $SelectedFilterAttribute2 -FilterValue2 $FilterValue2 -FilterOperator2 $FilterOperator2 `
                                                  -FilterLogic $FilterLogic -SelectedAttributes $SelectedAttributes -ObjectType "Computer"
                }
                "GroupMemberships" {
                    if (-not ([string]::IsNullOrWhiteSpace($FilterValue)) -and -not ([string]::IsNullOrWhiteSpace($SelectedFilterAttribute))) {
                        $ReportData = Get-ADGroupMembershipsReportData -FilterAttribute $SelectedFilterAttribute -FilterValue $FilterValue
                    } else {
                        Write-ADReportLog -Message "Filter attribute or value is empty for GroupMemberships query. Please specify a filter." -Type Warning
                        [System.Windows.Forms.MessageBox]::Show("Bitte geben Sie einen Filter (Attribut und Wert) fÃ¼r die Mitgliedschaftsabfrage an.", "Hinweis", "OK", "Information")
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
                    
                    # ZÃ¤hle die Anzahl der Ergebnisse
                    $Count = $ReportCollection.Count
                    Write-ADReportLog -Message "Query completed. $Count result(s) found for $ObjectType." -Type Info
                    
                    # ErgebniszÃ¤hler im Header aktualisieren
                    Update-ResultCounters -Results $ReportCollection
                    
                    if ($isUserSearch -and $ReportCollection.Count -eq 1 -and $ReportCollection[0].PSObject.Properties['SamAccountName']) {
                        $userSamAccountName = $ReportCollection[0].SamAccountName
                        
                        # Check if the "Mitgliedschaften" RadioButton is checked
                        if ($Global:RadioButtonGroupMemberships.IsChecked -eq $true) {
                            Write-ADReportLog -Message "Rufe Gruppenmitgliedschaften fÃ¼r Benutzer $($userSamAccountName) ab (RadioButton 'Mitgliedschaften' ist aktiv)..." -Type Info
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
                                }
                            } else {
                                # RadioButton "Mitgliedschaften" is NOT checked. Display user data as usual.
                                Write-ADReportLog -Message "RadioButton 'Mitgliedschaften' ist nicht aktiv. Zeige Benutzerdaten fÃ¼r $userSamAccountName." -Type Info
                                $Global:DataGridResults.ItemsSource = $ReportCollection
                                $Global:TextBlockStatus.Text = "Benutzer $userSamAccountName gefunden. Mitgliedschaften nicht abgefragt."
                                Update-ResultCounters -Results $ReportCollection
                                Update-ResultVisualization -Results $ReportCollection
                            }
                        } else {
                            # RadioButton "Mitgliedschaften" is NOT checked. Display user data as usual.
                            Write-ADReportLog -Message "RadioButton 'Mitgliedschaften' ist nicht aktiv. Zeige Benutzerdaten fÃ¼r $userSamAccountName." -Type Info
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
                    Update-ResultCounters -Results @() # Leeres Array fÃ¼r die ZÃ¤hler
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
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""
                    
                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "GivenName", "Surname", "mail", "Enabled", "LastLogonDate", "whenCreated", "LockedOut")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ReportData = Get-ADReportData -CustomFilter "*" -SelectedAttributes $QuickReportAttributes
                    Update-ADReportResults -Results $ReportData -StatusMessage "All users loaded."
                    $Global:TextBlockStatus.Text = "All users loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "All users loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading all users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickDisabledUsers.add_Click({
                Write-ADReportLog -Message "Loading disabled users..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "Enabled", "LastLogonDate")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes
                    
                    $ReportData = Get-ADReportData -CustomFilter "Enabled -eq `$false" -SelectedAttributes $QuickReportAttributes
                    Update-ADReportResults -Results $ReportData -StatusMessage "Disabled users loaded."
                    $Global:TextBlockStatus.Text = "Disabled users loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "Disabled users loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading disabled users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickLockedUsers.add_Click({
                Write-ADReportLog -Message "Loading locked out users..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "LockedOut", "LastLogonDate", "BadLogonCount")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Verwende Search-ADAccount statt Get-ADUser fÃ¼r gesperrte Konten (basierend auf Microsoft-Dokumentation)
                    $LockedOutAccounts = Search-ADAccount -LockedOut -UsersOnly -ErrorAction SilentlyContinue
                    
                    if ($LockedOutAccounts) {
                        # Hole detaillierte Informationen fÃ¼r jeden gesperrten Benutzer
                        $DetailedLockedAccounts = @()
                        foreach ($account in $LockedOutAccounts) {
                            $userDetails = Get-ADUser -Identity $account.SamAccountName -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                            if ($userDetails) {
                                $DetailedLockedAccounts += $userDetails
                            }
                        }
                        
                        if ($DetailedLockedAccounts.Count -gt 0) {
                            Update-ADReportResults -Results $DetailedLockedAccounts -StatusMessage "Locked out users loaded."
                            $Global:TextBlockStatus.Text = "Locked out users loaded. $($DetailedLockedAccounts.Count) record(s) found."
                            Write-ADReportLog -Message "Locked out users loaded. $($DetailedLockedAccounts.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "No locked out users found."
                            $Global:TextBlockStatus.Text = "No locked out users found."
                            Write-ADReportLog -Message "No locked out users found." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "No locked out users found."
                        $Global:TextBlockStatus.Text = "No locked out users found."
                        Write-ADReportLog -Message "No locked out users found." -Type Info
                    }
                } catch {
                    $ErrorMessage = "Error loading locked out users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickNeverExpire.add_Click({
                Write-ADReportLog -Message "Loading users with password never expires..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "PasswordNeverExpires", "Enabled")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ReportData = Get-ADReportData -CustomFilter "PasswordNeverExpires -eq `$true" -SelectedAttributes $QuickReportAttributes
                    Update-ADReportResults -Results $ReportData -StatusMessage "Users with password never expires loaded."
                    $Global:TextBlockStatus.Text = "Users with password never expires loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "Users with password never expires loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading users with password never expires: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickInactiveUsers.add_Click({
                Write-ADReportLog -Message "Loading inactive users (90 days)..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "LastLogonDate", "Enabled")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Verwende FileTime-Format fÃ¼r AD-Datumsvergleiche
                    $inactiveThreshold = (Get-Date).AddDays(-90).ToFileTime()
                    # Alternative: Lade alle Benutzer und filtere mit Where-Object
                    $AllUsers = Get-ADUser -Filter * -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                    
                    if ($AllUsers) {
                        $InactiveDate = (Get-Date).AddDays(-90)
                        $InactiveUsers = $AllUsers | Where-Object { 
                            $_.LastLogonDate -and $_.LastLogonDate -lt $InactiveDate 
                        }
                        
                        if ($InactiveUsers -and $InactiveUsers.Count -gt 0) {
                            Update-ADReportResults -Results $InactiveUsers -StatusMessage "Inactive users loaded."
                            $Global:TextBlockStatus.Text = "Inactive users loaded. $($InactiveUsers.Count) record(s) found."
                            Write-ADReportLog -Message "Inactive users loaded. $($InactiveUsers.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "No inactive users found."
                            $Global:TextBlockStatus.Text = "No users inactive for more than 90 days found."
                            Write-ADReportLog -Message "No inactive users found." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "Failed to query users."
                        $Global:TextBlockStatus.Text = "Failed to query users."
                        Write-ADReportLog -Message "Failed to query users." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading inactive users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickAdminUsers.add_Click({
                Write-ADReportLog -Message "Loading admin users..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    # Verbesserte Methode zum Finden von Admin-Benutzern
                    # Zuerst alle Benutzer laden
                    $AllUsers = Get-ADUser -Filter * -Properties DisplayName, SamAccountName, Enabled, LastLogonDate | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
                    
                    # Bekannte Admin-Gruppenbezeichnungen (deutsch und englisch)
                    $AdminGroups = @()
                    $AdminGroups += $Global:ADGroupNames.DomainAdmins
                    $AdminGroups += $Global:ADGroupNames.EnterpriseAdmins
                    $AdminGroups += $Global:ADGroupNames.Administrators
                    $AdminGroups += $Global:ADGroupNames.SchemaAdmins
                    $AdminGroups += @("AAD DC Administrators", "Azure AD-DC-Administratoren")
                    $AdminGroups += $Global:ADGroupNames.ServerOperators
                    $AdminGroups += $Global:ADGroupNames.AccountOperators
                    $AdminGroups += $Global:ADGroupNames.BackupOperators
                    
                    $AdminUsers = @()
                    # PrÃ¼fe Benutzer auf Admin-Rechte - erst SIDs der Admin-Gruppen ermitteln
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
                        # FÃ¼r jeden Benutzer die Gruppenmitgliedschaften prÃ¼fen
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
                    $ErrorMessage = "Error loading admin users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickRecentlyCreatedUsers.add_Click({
                Write-ADReportLog -Message "Loading recently created users (30 days)..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "whenCreated", "Enabled", "mail")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Lade alle Benutzer und filtere mit Where-Object fÃ¼r Datumsvergleiche
                    $AllUsers = Get-ADUser -Filter * -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                    
                    if ($AllUsers) {
                        $CreatedDate = (Get-Date).AddDays(-30)
                        $RecentUsers = $AllUsers | Where-Object { 
                            $_.whenCreated -and $_.whenCreated -gt $CreatedDate 
                        }
                        
                        if ($RecentUsers -and $RecentUsers.Count -gt 0) {
                            Update-ADReportResults -Results $RecentUsers -StatusMessage "Recently created users loaded."
                            $Global:TextBlockStatus.Text = "Recently created users loaded. $($RecentUsers.Count) record(s) found."
                            Write-ADReportLog -Message "Recently created users loaded. $($RecentUsers.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "No recently created users found."
                            $Global:TextBlockStatus.Text = "No users created in the last 30 days."
                            Write-ADReportLog -Message "No recently created users found." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "Failed to query users."
                        $Global:TextBlockStatus.Text = "Failed to query users."
                        Write-ADReportLog -Message "Failed to query users." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading recently created users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickPasswordExpiringSoon.add_Click({
                Write-ADReportLog -Message "Loading users with password expiring soon (7 days)..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "PasswordLastSet", "Enabled", "PasswordNeverExpires", "AccountExpirationDate")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Hole die Domain Password Policy
                    $DomainPasswordPolicy = Get-ADDefaultDomainPasswordPolicy
                    $MaxPasswordAge = $DomainPasswordPolicy.MaxPasswordAge.Days
                    
                    # Lade aktivierte Benutzer deren Passwort ablaufen kann
                    $AllActiveUsers = Get-ADUser -Filter "PasswordNeverExpires -eq `$false -and Enabled -eq `$true" -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                    
                    if ($AllActiveUsers) {
                        $ExpiryThreshold = (Get-Date).AddDays(-($MaxPasswordAge - 7))
                        $UsersPasswordExpiring = $AllActiveUsers | Where-Object { 
                            $_.PasswordLastSet -and $_.PasswordLastSet -lt $ExpiryThreshold 
                        }
                        
                        if ($UsersPasswordExpiring -and $UsersPasswordExpiring.Count -gt 0) {
                            Update-ADReportResults -Results $UsersPasswordExpiring -StatusMessage "Users with password expiring soon loaded."
                            $Global:TextBlockStatus.Text = "Users with password expiring soon loaded. $($UsersPasswordExpiring.Count) record(s) found."
                            Write-ADReportLog -Message "Users with password expiring soon loaded. $($UsersPasswordExpiring.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "No users with password expiring soon found."
                            $Global:TextBlockStatus.Text = "No users with passwords expiring in the next 7 days."
                            Write-ADReportLog -Message "No users with password expiring soon found." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "Failed to query users."
                        $Global:TextBlockStatus.Text = "Failed to query users."
                        Write-ADReportLog -Message "Failed to query users." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading users with password expiring soon: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickGroups.add_Click({
                Write-ADReportLog -Message "Loading all groups..." -Type Info
                try {
                    $Global:RadioButtonGroup.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("Name", "SamAccountName", "Description", "GroupCategory", "GroupScope", "whenCreated")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ReportData = Get-ADReportData -CustomFilter "*" -SelectedAttributes $QuickReportAttributes -ObjectType "Group"
                    Update-ADReportResults -Results $ReportData -StatusMessage "All groups loaded."
                    $Global:TextBlockStatus.Text = "All groups loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "All groups loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading all groups: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickSecurityGroups.add_Click({
                Write-ADReportLog -Message "Loading security groups..." -Type Info
                try {
                    $Global:RadioButtonGroup.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("Name", "SamAccountName", "Description", "GroupCategory", "GroupScope")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ReportData = Get-ADReportData -CustomFilter "GroupCategory -eq 'Security'" -SelectedAttributes $QuickReportAttributes -ObjectType "Group"
                    Update-ADReportResults -Results $ReportData -StatusMessage "Security groups loaded."
                    $Global:TextBlockStatus.Text = "Security groups loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "Security groups loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading security groups: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickComputers.add_Click({
                Write-ADReportLog -Message "Loading all computers..." -Type Info
                try {
                    $Global:RadioButtonComputer.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("Name", "DNSHostName", "OperatingSystem", "Enabled", "LastLogonDate", "whenCreated")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ReportData = Get-ADReportData -CustomFilter "*" -SelectedAttributes $QuickReportAttributes -ObjectType "Computer"
                    Update-ADReportResults -Results $ReportData -StatusMessage "All computers loaded."
                    $Global:TextBlockStatus.Text = "All computers loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "All computers loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading all computers: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickInactiveComputers.add_Click({
                Write-ADReportLog -Message "Loading inactive computers (90 days)..." -Type Info
                try {
                    $Global:RadioButtonComputer.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("Name", "DNSHostName", "LastLogonDate", "Enabled", "OperatingSystem")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Lade alle Computer und filtere mit Where-Object fÃ¼r Datumsvergleiche
                    $AllComputers = Get-ADComputer -Filter * -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                    
                    if ($AllComputers) {
                        $InactiveDate = (Get-Date).AddDays(-90)
                        $InactiveComputers = $AllComputers | Where-Object { 
                            $_.LastLogonDate -and $_.LastLogonDate -lt $InactiveDate 
                        }
                        
                        if ($InactiveComputers -and $InactiveComputers.Count -gt 0) {
                            Update-ADReportResults -Results $InactiveComputers -StatusMessage "Inactive computers loaded."
                            $Global:TextBlockStatus.Text = "Inactive computers loaded. $($InactiveComputers.Count) record(s) found."
                            Write-ADReportLog -Message "Inactive computers loaded. $($InactiveComputers.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "No inactive computers found."
                            $Global:TextBlockStatus.Text = "No computers inactive for more than 90 days found."
                            Write-ADReportLog -Message "No inactive computers found." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "Failed to query computers."
                        $Global:TextBlockStatus.Text = "Failed to query computers."
                        Write-ADReportLog -Message "Failed to query computers." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading inactive computers: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickExpiredPasswords.add_Click({
                Write-ADReportLog -Message "Loading users with expired passwords..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "PasswordLastSet", "Enabled", "PasswordNeverExpires")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Hole die Domain Password Policy
                    $DomainPasswordPolicy = Get-ADDefaultDomainPasswordPolicy
                    $MaxPasswordAge = $DomainPasswordPolicy.MaxPasswordAge.Days
                    
                    # Lade aktivierte Benutzer deren Passwort ablaufen kann
                    $AllActiveUsers = Get-ADUser -Filter "PasswordNeverExpires -eq `$false -and Enabled -eq `$true" -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                    
                    if ($AllActiveUsers) {
                        $ExpiredDate = (Get-Date).AddDays(-$MaxPasswordAge)
                        $UsersPasswordExpired = $AllActiveUsers | Where-Object { 
                            $_.PasswordLastSet -and $_.PasswordLastSet -lt $ExpiredDate 
                        }
                        
                        if ($UsersPasswordExpired -and $UsersPasswordExpired.Count -gt 0) {
                            Update-ADReportResults -Results $UsersPasswordExpired -StatusMessage "Users with expired passwords loaded."
                            $Global:TextBlockStatus.Text = "Users with expired passwords loaded. $($UsersPasswordExpired.Count) record(s) found."
                            Write-ADReportLog -Message "Users with expired passwords loaded. $($UsersPasswordExpired.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "No users with expired passwords found."
                            $Global:TextBlockStatus.Text = "No users with expired passwords found."
                            Write-ADReportLog -Message "No users with expired passwords found." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "Failed to query users."
                        $Global:TextBlockStatus.Text = "Failed to query users."
                        Write-ADReportLog -Message "Failed to query users." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading users with expired passwords: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickNeverLoggedOn.add_Click({
                Write-ADReportLog -Message "Loading users who never logged on..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "LastLogonDate", "whenCreated", "Enabled")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ReportData = Get-ADReportData -CustomFilter "LastLogonDate -notlike '*'" -SelectedAttributes $QuickReportAttributes -ObjectType "User"
                    Update-ADReportResults -Results $ReportData -StatusMessage "Users who never logged on loaded."
                    $Global:TextBlockStatus.Text = "Users who never logged on loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "Users who never logged on loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading users who never logged on: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickRecentlyDeletedUsers.add_Click({
                Write-ADReportLog -Message "Loading recently deleted users..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "whenDeleted", "isDeleted", "lastKnownParent", "objectClass")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # FÃ¼r gelÃ¶schte Objekte mÃ¼ssen wir den Deleted Objects Container abfragen
                    try {
                        $deletedAfter = (Get-Date).AddDays(-30)
                        $Domain = Get-ADDomain
                        $DeletedObjectsContainer = "CN=Deleted Objects,$($Domain.DistinguishedName)"
                        
                        # Verwende Get-ADObject mit IncludeDeletedObjects
                        $DeletedUsers = Get-ADObject -Filter {(ObjectClass -eq "user") -and (whenDeleted -gt $deletedAfter)} `
                                                    -IncludeDeletedObjects `
                                                    -SearchBase $DeletedObjectsContainer `
                                                    -Properties DisplayName, SamAccountName, whenDeleted, isDeleted, lastKnownParent `
                                                    -ErrorAction SilentlyContinue | 
                                       Select-Object DisplayName, SamAccountName, whenDeleted, isDeleted, lastKnownParent, ObjectClass

                        if ($DeletedUsers) {
                            $Global:DataGridResults.ItemsSource = $DeletedUsers
                            Write-ADReportLog -Message "Recently deleted users loaded. $($DeletedUsers.Count) result(s) found." -Type Info
                            Update-ResultCounters -Results $DeletedUsers
                        } else {
                            $Global:DataGridResults.ItemsSource = $null
                            Write-ADReportLog -Message "No recently deleted users found." -Type Info
                            $Global:TextBlockStatus.Text = "No recently deleted users found in the last 30 days."
                        }
                    } catch {
                        # Fallback wenn keine Berechtigung fÃ¼r Deleted Objects
                        Write-ADReportLog -Message "Cannot access deleted objects. Requires appropriate permissions." -Type Warning
                        $Global:DataGridResults.ItemsSource = $null
                        $Global:TextBlockStatus.Text = "Access to deleted objects denied. Administrator permissions required."
                    }
                } catch {
                    $ErrorMessage = "Error loading recently deleted users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickRecentlyModifiedUsers.add_Click({
                Write-ADReportLog -Message "Loading recently modified users..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "whenChanged", "whenCreated", "Enabled", "modifyTimeStamp")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Lade alle Benutzer und filtere mit Where-Object fÃ¼r Datumsvergleiche
                    $AllUsers = Get-ADUser -Filter * -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                    
                    if ($AllUsers) {
                        $ModifiedDate = (Get-Date).AddDays(-7)
                        $ModifiedUsers = $AllUsers | Where-Object { 
                            $_.whenChanged -and $_.whenChanged -gt $ModifiedDate 
                        }
                        
                        if ($ModifiedUsers -and $ModifiedUsers.Count -gt 0) {
                            Update-ADReportResults -Results $ModifiedUsers -StatusMessage "Recently modified users loaded."
                            $Global:TextBlockStatus.Text = "Recently modified users loaded. $($ModifiedUsers.Count) record(s) found."
                            Write-ADReportLog -Message "Recently modified users loaded. $($ModifiedUsers.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "No recently modified users found."
                            $Global:TextBlockStatus.Text = "No users modified in the last 7 days."
                            Write-ADReportLog -Message "No recently modified users found." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "Failed to query users."
                        $Global:TextBlockStatus.Text = "Failed to query users."
                        Write-ADReportLog -Message "Failed to query users." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading recently modified users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickInactiveUsersXDays.add_Click({
                Write-ADReportLog -Message "Loading inactive users (custom days)..." -Type Info
                try {
                    # Eingabedialog fÃ¼r Anzahl der Tage
                    Add-Type -AssemblyName Microsoft.VisualBasic
                    $days = [Microsoft.VisualBasic.Interaction]::InputBox('Enter number of days for inactivity check:', 'Inactive Users', '60')
                    
                    if ([string]::IsNullOrWhiteSpace($days)) {
                        Write-ADReportLog -Message "User cancelled inactive users dialog." -Type Info
                        return
                    }
                    
                    if (-not ($days -match '^\d+$') -or [int]$days -lt 1) {
                        [System.Windows.MessageBox]::Show("Please enter a valid number of days (minimum 1).", "Invalid Input", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                    
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "LastLogonDate", "Enabled", "whenCreated")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Lade alle Benutzer und filtere mit Where-Object fÃ¼r Datumsvergleiche
                    $AllUsers = Get-ADUser -Filter * -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                    
                    if ($AllUsers) {
                        $InactiveDate = (Get-Date).AddDays(-[int]$days)
                        $InactiveUsers = $AllUsers | Where-Object { 
                            $_.LastLogonDate -and $_.LastLogonDate -lt $InactiveDate 
                        }
                        
                        if ($InactiveUsers -and $InactiveUsers.Count -gt 0) {
                            Update-ADReportResults -Results $InactiveUsers -StatusMessage "Inactive users (>$days days) loaded."
                            $Global:TextBlockStatus.Text = "Inactive users (>$days days) loaded. $($InactiveUsers.Count) record(s) found."
                            Write-ADReportLog -Message "Inactive users (>$days days) loaded. $($InactiveUsers.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "No inactive users found."
                            $Global:TextBlockStatus.Text = "No users inactive for more than $days days found."
                            Write-ADReportLog -Message "No inactive users found." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "Failed to query users."
                        $Global:TextBlockStatus.Text = "Failed to query users."
                        Write-ADReportLog -Message "Failed to query users." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading inactive users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickUsersWithoutManager.add_Click({
                Write-ADReportLog -Message "Loading users without manager..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "Department", "Title", "Enabled", "manager")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Manager-Attribut unterstÃ¼tzt nur -eq und -ne Operatoren
                    # Lade alle aktivierten Benutzer und filtere dann nach fehlenden Managern
                    $AllEnabledUsers = Get-ADUser -Filter "Enabled -eq `$true" -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                    
                    if ($AllEnabledUsers) {
                        # Filtere Benutzer ohne Manager
                        $UsersWithoutManager = $AllEnabledUsers | Where-Object { [string]::IsNullOrWhiteSpace($_.manager) }
                        
                        if ($UsersWithoutManager.Count -gt 0) {
                            Update-ADReportResults -Results $UsersWithoutManager -StatusMessage "Users without manager loaded."
                            $Global:TextBlockStatus.Text = "Users without manager loaded. $($UsersWithoutManager.Count) record(s) found."
                            Write-ADReportLog -Message "Users without manager loaded. $($UsersWithoutManager.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "All users have a manager assigned."
                            $Global:TextBlockStatus.Text = "All users have a manager assigned."
                            Write-ADReportLog -Message "All users have a manager assigned." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "No enabled users found."
                        $Global:TextBlockStatus.Text = "No enabled users found."
                        Write-ADReportLog -Message "No enabled users found." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading users without manager: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickUsersMissingRequiredAttributes.add_Click({
                Write-ADReportLog -Message "Loading users missing required attributes..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    # Lade alle Benutzer mit wichtigen Attributen
                    $AllUsers = Get-ADUser -Filter "Enabled -eq `$true" -Properties DisplayName, SamAccountName, mail, telephoneNumber, Department, Title, manager, Enabled
                    
                    # Filtere Benutzer mit fehlenden Attributen
                    $UsersWithMissingAttributes = @()
                    foreach ($user in $AllUsers) {
                        $missingAttributes = @()
                        
                        if ([string]::IsNullOrWhiteSpace($user.DisplayName)) { $missingAttributes += "DisplayName" }
                        if ([string]::IsNullOrWhiteSpace($user.mail)) { $missingAttributes += "Email" }
                        if ([string]::IsNullOrWhiteSpace($user.telephoneNumber)) { $missingAttributes += "Phone" }
                        if ([string]::IsNullOrWhiteSpace($user.Department)) { $missingAttributes += "Department" }
                        if ([string]::IsNullOrWhiteSpace($user.Title)) { $missingAttributes += "Title" }
                        if ([string]::IsNullOrWhiteSpace($user.manager)) { $missingAttributes += "Manager" }
                        
                        if ($missingAttributes.Count -gt 0) {
                            $UsersWithMissingAttributes += [PSCustomObject]@{
                                DisplayName = $user.DisplayName
                                SamAccountName = $user.SamAccountName
                                mail = $user.mail
                                Department = $user.Department
                                Title = $user.Title
                                Enabled = $user.Enabled
                                MissingAttributes = $missingAttributes -join ", "
                                MissingCount = $missingAttributes.Count
                            }
                        }
                    }
                    
                    if ($UsersWithMissingAttributes.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $UsersWithMissingAttributes | Sort-Object MissingCount -Descending
                        Write-ADReportLog -Message "Users missing required attributes loaded. $($UsersWithMissingAttributes.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $UsersWithMissingAttributes
                        $Global:TextBlockStatus.Text = "Users missing required attributes loaded. $($UsersWithMissingAttributes.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No users with missing required attributes found." -Type Info
                        $Global:TextBlockStatus.Text = "All active users have required attributes filled."
                    }
                } catch {
                    $ErrorMessage = "Error loading users missing required attributes: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickUsersDuplicateLogonNames.add_Click({
                Write-ADReportLog -Message "Loading users with duplicate logon names..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    # Lade alle Benutzer
                    $AllUsers = Get-ADUser -Filter * -Properties DisplayName, SamAccountName, UserPrincipalName, Enabled, whenCreated
                    
                    # Gruppiere nach SamAccountName um Duplikate zu finden
                    $DuplicateSamAccounts = $AllUsers | Group-Object SamAccountName | Where-Object { $_.Count -gt 1 }
                    
                    # Gruppiere auch nach UserPrincipalName
                    $DuplicateUPNs = $AllUsers | Where-Object { $_.UserPrincipalName } | Group-Object UserPrincipalName | Where-Object { $_.Count -gt 1 }
                    
                    $DuplicateUsers = @()
                    
                    # Verarbeite SamAccountName Duplikate
                    foreach ($group in $DuplicateSamAccounts) {
                        foreach ($user in $group.Group) {
                            $DuplicateUsers += [PSCustomObject]@{
                                DisplayName = $user.DisplayName
                                SamAccountName = $user.SamAccountName
                                UserPrincipalName = $user.UserPrincipalName
                                Enabled = $user.Enabled
                                whenCreated = $user.whenCreated
                                DuplicateType = "SamAccountName"
                                DuplicateCount = $group.Count
                            }
                        }
                    }
                    
                    # Verarbeite UPN Duplikate
                    foreach ($group in $DuplicateUPNs) {
                        foreach ($user in $group.Group) {
                            # PrÃ¼fe ob dieser User nicht schon als SamAccountName Duplikat erfasst wurde
                            if (-not ($DuplicateUsers | Where-Object { $_.SamAccountName -eq $user.SamAccountName })) {
                                $DuplicateUsers += [PSCustomObject]@{
                                    DisplayName = $user.DisplayName
                                    SamAccountName = $user.SamAccountName
                                    UserPrincipalName = $user.UserPrincipalName
                                    Enabled = $user.Enabled
                                    whenCreated = $user.whenCreated
                                    DuplicateType = "UserPrincipalName"
                                    DuplicateCount = $group.Count
                                }
                            }
                        }
                    }
                    
                    if ($DuplicateUsers.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $DuplicateUsers | Sort-Object SamAccountName
                        Write-ADReportLog -Message "Users with duplicate logon names loaded. $($DuplicateUsers.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $DuplicateUsers
                        $Global:TextBlockStatus.Text = "Users with duplicate logon names loaded. $($DuplicateUsers.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No users with duplicate logon names found." -Type Info
                        $Global:TextBlockStatus.Text = "No duplicate logon names found."
                    }
                } catch {
                    $ErrorMessage = "Error loading users with duplicate logon names: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickOrphanedSIDsUsers.add_Click({
                Write-ADReportLog -Message "Loading orphaned SIDs (Foreign Security Principals)..." -Type Info
                try {
                    # Lade Foreign Security Principals
                    $Domain = Get-ADDomain
                    $ForeignSecurityPrincipals = Get-ADObject -Filter * -SearchBase "CN=ForeignSecurityPrincipals,$($Domain.DistinguishedName)" -Properties Name, ObjectSID, whenCreated, whenChanged
                    
                    $OrphanedSIDs = @()
                    
                    foreach ($fsp in $ForeignSecurityPrincipals) {
                        # Versuche den SID aufzulÃ¶sen
                        $resolved = $false
                        $resolvedName = "Unknown"
                        $sidString = $fsp.Name
                        
                        try {
                            # Versuche SID zu einem Namen aufzulÃ¶sen
                            $sid = New-Object System.Security.Principal.SecurityIdentifier($sidString)
                            $resolvedName = $sid.Translate([System.Security.Principal.NTAccount]).Value
                            $resolved = $true
                        } catch {
                            # SID konnte nicht aufgelÃ¶st werden - wahrscheinlich verwaist
                            $resolved = $false
                        }
                        
                        if (-not $resolved) {
                            $OrphanedSIDs += [PSCustomObject]@{
                                Name = $fsp.Name
                                DistinguishedName = $fsp.DistinguishedName
                                ObjectSID = $sidString
                                Status = "Orphaned"
                                ResolvedName = "Cannot resolve"
                                whenCreated = $fsp.whenCreated
                                whenChanged = $fsp.whenChanged
                            }
                        } else {
                            # Optional: Auch aufgelÃ¶ste SIDs anzeigen
                            $OrphanedSIDs += [PSCustomObject]@{
                                Name = $fsp.Name
                                DistinguishedName = $fsp.DistinguishedName
                                ObjectSID = $sidString
                                Status = "Active"
                                ResolvedName = $resolvedName
                                whenCreated = $fsp.whenCreated
                                whenChanged = $fsp.whenChanged
                            }
                        }
                    }
                    
                    # Filtere nur die verwaisten SIDs
                    $OrphanedOnly = $OrphanedSIDs | Where-Object { $_.Status -eq "Orphaned" }
                    
                    if ($OrphanedOnly.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $OrphanedOnly
                        Write-ADReportLog -Message "Orphaned SIDs loaded. $($OrphanedOnly.Count) orphaned out of $($OrphanedSIDs.Count) total FSPs found." -Type Info
                        Update-ResultCounters -Results $OrphanedOnly
                        $Global:TextBlockStatus.Text = "Orphaned SIDs loaded. $($OrphanedOnly.Count) orphaned out of $($OrphanedSIDs.Count) total FSPs."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No orphaned SIDs found. All $($OrphanedSIDs.Count) FSPs can be resolved." -Type Info
                        $Global:TextBlockStatus.Text = "No orphaned SIDs found. All Foreign Security Principals are valid."
                    }
                } catch {
                    $ErrorMessage = "Error loading orphaned SIDs: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickReversibleEncryption.add_Click({
                Write-ADReportLog -Message "Loading users with reversible encryption..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $ReversibleUsers = Get-ReversibleEncryptionUsers
                    if ($ReversibleUsers -and $ReversibleUsers.Count -gt 0) {
                        Update-ADReportResults -Results $ReversibleUsers
                        Write-ADReportLog -Message "Users with reversible encryption loaded. $($ReversibleUsers.Count) result(s) found." -Type Info
                        $Global:TextBlockStatus.Text = "Users with reversible encryption loaded. $($ReversibleUsers.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No users with reversible encryption found." -Type Info
                        $Global:TextBlockStatus.Text = "No users with reversible encryption found."
                    }
                } catch {
                    $ErrorMessage = "Error loading users with reversible encryption: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickKerberosDES.add_Click({
                Write-ADReportLog -Message "Loading users with Kerberos DES encryption..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $DESUsers = Get-KerberosDESUsers
                    if ($DESUsers -and $DESUsers.Count -gt 0) {
                        Update-ADReportResults -Results $DESUsers
                        Write-ADReportLog -Message "Users with Kerberos DES loaded. $($DESUsers.Count) result(s) found." -Type Info
                        $Global:TextBlockStatus.Text = "Users with Kerberos DES loaded. $($DESUsers.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No users with Kerberos DES found." -Type Info
                        $Global:TextBlockStatus.Text = "No users with Kerberos DES encryption found."
                    }
                } catch {
                    $ErrorMessage = "Error loading users with Kerberos DES: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickUsersWithSPN.add_Click({
                Write-ADReportLog -Message "Loading users with Service Principal Names..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $SPNUsers = Get-UsersWithSPN
                    if ($SPNUsers -and $SPNUsers.Count -gt 0) {
                        Update-ADReportResults -Results $SPNUsers
                        Write-ADReportLog -Message "Users with SPN loaded. $($SPNUsers.Count) result(s) found." -Type Info
                        $Global:TextBlockStatus.Text = "Users with SPN loaded. $($SPNUsers.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No users with SPN found." -Type Info
                        $Global:TextBlockStatus.Text = "No users with Service Principal Names found."
                    }
                } catch {
                    $ErrorMessage = "Error loading users with SPN: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickGuestAccountStatus.add_Click({
                Write-ADReportLog -Message "Analyzing Guest account status..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $GuestStatus = Get-GuestAccountStatus
                    if ($GuestStatus -and $GuestStatus.Count -gt 0) {
                        Update-ADReportResults -Results $GuestStatus
                        Write-ADReportLog -Message "Guest account analysis completed. $($GuestStatus.Count) result(s) found." -Type Info
                        $Global:TextBlockStatus.Text = "Guest account analysis completed. $($GuestStatus.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No Guest account information found." -Type Info
                        $Global:TextBlockStatus.Text = "No Guest account found."
                    }
                } catch {
                    $ErrorMessage = "Error analyzing Guest account: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickUsersByDepartment.add_Click({
                Write-ADReportLog -Message "Loading users by department..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $UsersByDept = Get-UsersByDepartment
                    if ($UsersByDept -and $UsersByDept.Count -gt 0) {
                        Update-ADReportResults -Results $UsersByDept
                        Write-ADReportLog -Message "Users by department loaded. $($UsersByDept.Count) result(s) found." -Type Info
                        $Global:TextBlockStatus.Text = "Users by department loaded. $($UsersByDept.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No department data found." -Type Info
                        $Global:TextBlockStatus.Text = "No users with department information found."
                    }
                } catch {
                    $ErrorMessage = "Error loading users by department: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickUsersByManager.add_Click({
                Write-ADReportLog -Message "Loading users by manager..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $UsersByMgr = Get-UsersByManager
                    if ($UsersByMgr -and $UsersByMgr.Count -gt 0) {
                        Update-ADReportResults -Results $UsersByMgr
                        Write-ADReportLog -Message "Users by manager loaded. $($UsersByMgr.Count) result(s) found." -Type Info
                        $Global:TextBlockStatus.Text = "Users by manager loaded. $($UsersByMgr.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No manager data found." -Type Info
                        $Global:TextBlockStatus.Text = "No users with manager information found."
                    }
                } catch {
                    $ErrorMessage = "Error loading users by manager: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickRemoteAccessUsers.add_Click({
                Write-ADReportLog -Message "Loading remote access users..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $RemoteUsers = Get-RemoteAccessUsers
                    if ($RemoteUsers -and $RemoteUsers.Count -gt 0) {
                        Update-ADReportResults -Results $RemoteUsers
                        Write-ADReportLog -Message "Remote access users loaded. $($RemoteUsers.Count) result(s) found." -Type Info
                        $Global:TextBlockStatus.Text = "Remote access users loaded. $($RemoteUsers.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No remote access users found." -Type Info
                        $Global:TextBlockStatus.Text = "No remote access users found."
                    }
                } catch {
                    $ErrorMessage = "Error loading remote access users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickDistributionGroups.add_Click({
                Write-ADReportLog -Message "Loading distribution groups..." -Type Info
                try {
                    $Global:RadioButtonGroup.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("Name", "SamAccountName", "Description", "GroupCategory", "GroupScope", "mail")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ReportData = Get-ADReportData -CustomFilter "GroupCategory -eq 'Distribution'" -SelectedAttributes $QuickReportAttributes -ObjectType "Group"
                    Update-ADReportResults -Results $ReportData -StatusMessage "Distribution groups loaded."
                    $Global:TextBlockStatus.Text = "Distribution groups loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "Distribution groups loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading distribution groups: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickWeakPasswordPolicy.add_Click({
                Write-ADReportLog -Message "Loading users with weak password policies..." -Type Info
                try {
                    $WeakPasswordUsers = Get-WeakPasswordPolicyUsers
                    if ($WeakPasswordUsers -and $WeakPasswordUsers.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $WeakPasswordUsers
                        Write-ADReportLog -Message "Users with weak password policies loaded. $($WeakPasswordUsers.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $WeakPasswordUsers
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No users with weak password policies found." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading users with weak password policies: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    $Global:DataGridResults.ItemsSource = $null
                }
            })

            $ButtonQuickRiskyGroupMemberships.add_Click({
                Write-ADReportLog -Message "Loading risky group memberships..." -Type Info
                try {
                    $RiskyMemberships = Get-RiskyGroupMemberships
                    if ($RiskyMemberships -and $RiskyMemberships.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $RiskyMemberships
                        Write-ADReportLog -Message "Risky group memberships loaded. $($RiskyMemberships.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $RiskyMemberships
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No risky group memberships found." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading risky group memberships: $($_.Exception.Message)"
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

            $ButtonQuickOUHierarchy.add_Click({
                Write-ADReportLog -Message "Loading OU hierarchy..." -Type Info
                try {
                    $OUHierarchy = Get-ADOUHierarchyReport
                    if ($OUHierarchy -and $OUHierarchy.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $OUHierarchy
                        Write-ADReportLog -Message "OU hierarchy loaded. $($OUHierarchy.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $OUHierarchy
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No OU hierarchy information found." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading OU hierarchy: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    $Global:DataGridResults.ItemsSource = $null
                }
            })

            $ButtonQuickSitesSubnets.add_Click({
                Write-ADReportLog -Message "Loading AD sites and subnets..." -Type Info
                try {
                    $SitesSubnets = @(Get-ADSitesAndSubnetsReport)
                    if ($SitesSubnets -and $SitesSubnets.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $SitesSubnets
                        Write-ADReportLog -Message "AD sites and subnets loaded. $($SitesSubnets.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $SitesSubnets
                        $Global:TextBlockStatus.Text = "AD sites and subnets loaded. $($SitesSubnets.Count) result(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No AD sites and subnets information found." -Type Warning
                        $Global:TextBlockStatus.Text = "No AD sites and subnets information found."
                    }
                } catch {
                    $ErrorMessage = "Error loading AD sites and subnets: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    $Global:DataGridResults.ItemsSource = $null
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
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

    # --- Event Handler fÃ¼r neue erweiterte Reports ---
    
    # Organisationsstruktur-Reports
    if ($null -ne $Global:ButtonQuickDepartmentStats) {
        $ButtonQuickDepartmentStats.add_Click({
        Write-ADReportLog -Message "Loading department statistics..." -Type Info
        try {
            $DeptStats = Get-DepartmentStatistics
            if ($DeptStats -and $DeptStats.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $DeptStats
                Write-ADReportLog -Message "Department statistics loaded. $($DeptStats.Count) department(s) analyzed." -Type Info
                Update-ResultCounters -Results $DeptStats
                $Global:TextBlockStatus.Text = "Department statistics loaded. $($DeptStats.Count) department(s) analyzed."
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No department statistics found." -Type Warning
                $Global:TextBlockStatus.Text = "No department statistics found."
            }
        } catch {
            $ErrorMessage = "Error loading department statistics: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
        }
    })
    }

    if ($null -ne $Global:ButtonQuickDepartmentSecurity) {
        $ButtonQuickDepartmentSecurity.add_Click({
        Write-ADReportLog -Message "Loading department security analysis..." -Type Info
        try {
            $DeptSecurity = Get-DepartmentSecurityRisks
            if ($DeptSecurity -and $DeptSecurity.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $DeptSecurity
                Write-ADReportLog -Message "Department security analysis loaded. $($DeptSecurity.Count) department(s) analyzed." -Type Info
                Update-ResultCounters -Results $DeptSecurity
                $Global:TextBlockStatus.Text = "Department security analysis loaded. $($DeptSecurity.Count) department(s) analyzed."
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No department security data found." -Type Warning
                $Global:TextBlockStatus.Text = "No department security data found."
            }
        } catch {
            $ErrorMessage = "Error loading department security analysis: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
        }
    })
    }

    # Kerberos Security
    if ($null -ne $Global:ButtonQuickKerberoastable) {
        $ButtonQuickKerberoastable.add_Click({
        Write-ADReportLog -Message "Loading Kerberoastable accounts..." -Type Info
        try {
            $Kerberoastable = Get-KerberoastableAccounts
            if ($Kerberoastable -and $Kerberoastable.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $Kerberoastable
                Write-ADReportLog -Message "Kerberoastable accounts loaded. $($Kerberoastable.Count) account(s) found." -Type Info
                Update-ResultCounters -Results $Kerberoastable
                $Global:TextBlockStatus.Text = "Kerberoastable accounts loaded. $($Kerberoastable.Count) account(s) with SPNs found."
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No Kerberoastable accounts found." -Type Info
                $Global:TextBlockStatus.Text = "No Kerberoastable accounts (users with SPNs) found."
            }
        } catch {
            $ErrorMessage = "Error loading Kerberoastable accounts: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
        }
    })
    }

    if ($null -ne $Global:ButtonQuickASREPRoastable) {
        $ButtonQuickASREPRoastable.add_Click({
        Write-ADReportLog -Message "Loading ASREPRoastable accounts..." -Type Info
        try {
            $ASREPRoastable = Get-ASREPRoastableAccounts
            if ($ASREPRoastable -and $ASREPRoastable.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $ASREPRoastable
                Write-ADReportLog -Message "ASREPRoastable accounts loaded. $($ASREPRoastable.Count) account(s) found." -Type Info
                Update-ResultCounters -Results $ASREPRoastable
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No ASREPRoastable accounts found." -Type Info
            }
        } catch {
            $ErrorMessage = "Error loading ASREPRoastable accounts: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
    }

    if ($null -ne $Global:ButtonQuickDelegation) {
        $ButtonQuickDelegation.add_Click({
        Write-ADReportLog -Message "Loading delegation analysis..." -Type Info
        try {
            $Delegation = Get-DelegationAnalysis
            if ($Delegation -and $Delegation.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $Delegation
                Write-ADReportLog -Message "Delegation analysis loaded. $($Delegation.Count) delegated object(s) found." -Type Info
                Update-ResultCounters -Results $Delegation
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No delegation settings found." -Type Info
            }
        } catch {
            $ErrorMessage = "Error loading delegation analysis: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
    }

    # Advanced Security
    if ($null -ne $Global:ButtonQuickDCSyncRights) {
        $ButtonQuickDCSyncRights.add_Click({
        Write-ADReportLog -Message "Loading DCSync rights analysis..." -Type Info
        try {
            $DCSyncRights = Get-DCSyncRights
            if ($DCSyncRights -and $DCSyncRights.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $DCSyncRights
                Write-ADReportLog -Message "DCSync rights analysis loaded. $($DCSyncRights.Count) identities with DCSync found." -Type Info
                Update-ResultCounters -Results $DCSyncRights
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No DCSync rights found." -Type Info
            }
        } catch {
            $ErrorMessage = "Error loading DCSync rights: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
    } else {
        Write-ADReportLog -Message "ButtonQuickDCSyncRights nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
    }

    # Advanced Security
    try {
        if ($null -ne $Global:ButtonQuickSchemaAdmins) {
            $ButtonQuickSchemaAdmins.add_Click({
        Write-ADReportLog -Message "Loading Schema Admin paths..." -Type Info
        try {
            $SchemaAdmins = Get-SchemaAdminPaths
            if ($SchemaAdmins -and $SchemaAdmins.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $SchemaAdmins
                Write-ADReportLog -Message "Schema Admin paths loaded. $($SchemaAdmins.Count) path(s) found." -Type Info
                Update-ResultCounters -Results $SchemaAdmins
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No Schema Admin paths found." -Type Info
            }
        } catch {
            $ErrorMessage = "Error loading Schema Admin paths: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickSchemaAdmins nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickSchemaAdmins: $($_.Exception.Message)" -Type Warning -Terminal
    }

    try {
        if ($null -ne $Global:ButtonQuickCertificateAnalysis) {
            $ButtonQuickCertificateAnalysis.add_Click({
        Write-ADReportLog -Message "Loading certificate security analysis..." -Type Info
        try {
            $CertAnalysis = Get-CertificateSecurityAnalysis
            if ($CertAnalysis -and $CertAnalysis.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $CertAnalysis
                Write-ADReportLog -Message "Certificate security analysis loaded. $($CertAnalysis.Count) finding(s)." -Type Info
                Update-ResultCounters -Results $CertAnalysis
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No certificate security findings." -Type Info
            }
        } catch {
            $ErrorMessage = "Error loading certificate analysis: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickCertificateAnalysis nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickCertificateAnalysis: $($_.Exception.Message)" -Type Warning -Terminal
    }

    # Advanced Monitoring
    try {
        if ($null -ne $Global:ButtonQuickSYSVOLHealth) {
            $ButtonQuickSYSVOLHealth.add_Click({
        Write-ADReportLog -Message "Loading SYSVOL health check..." -Type Info
        try {
            $SYSVOLHealth = @(Get-SYSVOLHealthCheck)
            if ($SYSVOLHealth -and $SYSVOLHealth.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $SYSVOLHealth
                Write-ADReportLog -Message "SYSVOL health check loaded. $($SYSVOLHealth.Count) DC(s) checked." -Type Info
                Update-ResultCounters -Results $SYSVOLHealth
                $Global:TextBlockStatus.Text = "SYSVOL health check loaded. $($SYSVOLHealth.Count) domain controller(s) checked."
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No SYSVOL health data found." -Type Warning
                $Global:TextBlockStatus.Text = "No SYSVOL health data found."
            }
        } catch {
            $ErrorMessage = "Error loading SYSVOL health: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickSYSVOLHealth nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickSYSVOLHealth: $($_.Exception.Message)" -Type Warning -Terminal
    }

    try {
        if ($null -ne $Global:ButtonQuickDNSHealth) {
            $ButtonQuickDNSHealth.add_Click({
        Write-ADReportLog -Message "Loading DNS health analysis..." -Type Info
        try {
            $DNSHealth = Get-DNSHealthAnalysis
            if ($DNSHealth -and $DNSHealth.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $DNSHealth
                Write-ADReportLog -Message "DNS health analysis loaded. $($DNSHealth.Count) item(s) analyzed." -Type Info
                Update-ResultCounters -Results $DNSHealth
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No DNS health data found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading DNS health: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickDNSHealth nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickDNSHealth: $($_.Exception.Message)" -Type Warning -Terminal
    }

    try {
        if ($null -ne $Global:ButtonQuickBackupStatus) {
            $ButtonQuickBackupStatus.add_Click({
        Write-ADReportLog -Message "Loading backup readiness status..." -Type Info
        try {
            $BackupStatus = Get-BackupReadinessStatus
            if ($BackupStatus -and $BackupStatus.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $BackupStatus
                Write-ADReportLog -Message "Backup readiness status loaded. $($BackupStatus.Count) check(s) performed." -Type Info
                Update-ResultCounters -Results $BackupStatus
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No backup status data found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading backup status: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickBackupStatus nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickBackupStatus: $($_.Exception.Message)" -Type Warning -Terminal
    }

    # Schema & Trusts
    try {
        if ($null -ne $Global:ButtonQuickSchemaAnalysis) {
            $ButtonQuickSchemaAnalysis.add_Click({
        Write-ADReportLog -Message "Loading schema analysis..." -Type Info
        try {
            $SchemaAnalysis = Get-SchemaAnalysis
            if ($SchemaAnalysis -and $SchemaAnalysis.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $SchemaAnalysis
                Write-ADReportLog -Message "Schema analysis loaded. $($SchemaAnalysis.Count) item(s) found." -Type Info
                Update-ResultCounters -Results $SchemaAnalysis
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No schema data found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading schema analysis: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickSchemaAnalysis nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickSchemaAnalysis: $($_.Exception.Message)" -Type Warning -Terminal
    }

    try {
        if ($null -ne $Global:ButtonQuickTrustRelationships) {
            $ButtonQuickTrustRelationships.add_Click({
        Write-ADReportLog -Message "Loading trust relationship analysis..." -Type Info
        try {
            $TrustAnalysis = Get-TrustRelationshipAnalysis
            if ($TrustAnalysis -and $TrustAnalysis.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $TrustAnalysis
                Write-ADReportLog -Message "Trust relationship analysis loaded. $($TrustAnalysis.Count) trust(s) analyzed." -Type Info
                Update-ResultCounters -Results $TrustAnalysis
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No trust relationships found." -Type Info
            }
        } catch {
            $ErrorMessage = "Error loading trust analysis: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickTrustRelationships nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickTrustRelationships: $($_.Exception.Message)" -Type Warning -Terminal
    }

    try {
        if ($null -ne $Global:ButtonQuickQuotasLimits) {
            $ButtonQuickQuotasLimits.add_Click({
        Write-ADReportLog -Message "Loading quotas and limits analysis..." -Type Info
        try {
            $QuotasLimits = Get-QuotasAndLimits
            if ($QuotasLimits -and $QuotasLimits.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $QuotasLimits
                Write-ADReportLog -Message "Quotas and limits analysis loaded. $($QuotasLimits.Count) item(s) analyzed." -Type Info
                Update-ResultCounters -Results $QuotasLimits
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No quota data found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading quotas and limits: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickQuotasLimits nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickQuotasLimits: $($_.Exception.Message)" -Type Warning -Terminal
    }

    $ButtonExportCSV.add_Click({
        Write-ADReportLog -Message "Preparing CSV export..." -Type Info
        if ($null -eq $Global:DataGridResults.ItemsSource -or $Global:DataGridResults.Items.Count -eq 0) {
            Write-ADReportLog -Message "No data available for export." -Type Warning
            [System.Windows.Forms.MessageBox]::Show("Es sind keine Daten zum Exportieren vorhanden.", "Hinweis", "OK", "Information") | Out-Null
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

    # --- Event Handler fÃ¼r Attribute Selection Buttons ---
    if ($null -ne $Global:ButtonSelectAllAttributes) {
        $ButtonSelectAllAttributes.add_Click({
            Write-ADReportLog -Message "Selecting all attributes..." -Type Info
            try {
                # Select all items in all three ListBoxes
                if ($Global:ListBoxBasicAttributes) {
                    foreach ($item in $Global:ListBoxBasicAttributes.Items) {
                        $item.IsSelected = $true
                    }
                }
                if ($Global:ListBoxSecurityAttributes) {
                    foreach ($item in $Global:ListBoxSecurityAttributes.Items) {
                        $item.IsSelected = $true
                    }
                }
                if ($Global:ListBoxExtendedAttributes) {
                    foreach ($item in $Global:ListBoxExtendedAttributes.Items) {
                        $item.IsSelected = $true
                    }
                }
                Write-ADReportLog -Message "All attributes selected." -Type Info
            } catch {
                Write-ADReportLog -Message "Error selecting all attributes: $($_.Exception.Message)" -Type Error
            }
        })
    }

    if ($null -ne $Global:ButtonSelectNoneAttributes) {
        $ButtonSelectNoneAttributes.add_Click({
            Write-ADReportLog -Message "Deselecting all attributes..." -Type Info
            try {
                # Deselect all items in all three ListBoxes
                if ($Global:ListBoxBasicAttributes) {
                    foreach ($item in $Global:ListBoxBasicAttributes.Items) {
                        $item.IsSelected = $false
                    }
                }
                if ($Global:ListBoxSecurityAttributes) {
                    foreach ($item in $Global:ListBoxSecurityAttributes.Items) {
                        $item.IsSelected = $false
                    }
                }
                if ($Global:ListBoxExtendedAttributes) {
                    foreach ($item in $Global:ListBoxExtendedAttributes.Items) {
                        $item.IsSelected = $false
                    }
                }
                Write-ADReportLog -Message "All attributes deselected." -Type Info
            } catch {
                Write-ADReportLog -Message "Error deselecting all attributes: $($_.Exception.Message)" -Type Error
            }
        })
    }

    # --- Event Handler fÃ¼r Help and About Buttons ---
    if ($null -ne $Global:ButtonHelp) {
        $ButtonHelp.add_Click({
            Write-ADReportLog -Message "Showing help dialog..." -Type Info
            [System.Windows.MessageBox]::Show(
                "easyADReport Help:`n`n" +
                "1. Select Object Type: Choose between Users, Groups, Computers, or Group Memberships`n`n" +
                "2. Set Filters: Configure search filters using attribute, operator, and value`n`n" +
                "3. Select Attributes: Choose which attributes to include in the report`n`n" +
                "4. Quick Reports: Use predefined reports for common AD queries`n`n" +
                "5. Export Options: Export results to CSV or HTML format`n`n" +
                "For more information, visit the documentation.",
                "Help",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            ) | Out-Null
        })
    }

    if ($null -ne $Global:ButtonAbout) {
        $ButtonAbout.add_Click({
            Write-ADReportLog -Message "Showing about dialog..." -Type Info
            [System.Windows.MessageBox]::Show(
                "easyADReport v0.3.3`n`n" +
                "A comprehensive Active Directory reporting tool`n`n" +
                "Features:`n" +
                "â€¢ Advanced filtering and search capabilities`n" +
                "â€¢ Multiple export formats (CSV, HTML)`n" +
                "â€¢ Quick report templates for common scenarios`n" +
                "â€¢ Security and compliance reports`n" +
                "â€¢ Multi-language support`n`n" +
                "Â© 2024 Your Organization",
                "About easyADReport",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            ) | Out-Null
        })
    }

    # --- Event Handler fÃ¼r neue Quick Report Buttons ---
    
    # Event Handler fÃ¼r ButtonQuickStalePasswords
    if ($null -ne $Global:ButtonQuickStalePasswords) {
        $ButtonQuickStalePasswords.add_Click({
            Write-ADReportLog -Message "Loading users with stale passwords..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                $Global:TextBoxFilterValue1.Text = ""
                $Global:TextBoxFilterValue2.Text = ""

                $StalePasswords = Get-StalePasswords -Days 90
                if ($StalePasswords -and $StalePasswords.Count -gt 0) {
                    Update-ADReportResults -Results $StalePasswords
                    Write-ADReportLog -Message "Stale passwords loaded. $($StalePasswords.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $StalePasswords
                    $Global:TextBlockStatus.Text = "Stale passwords loaded. $($StalePasswords.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No users with stale passwords found." -Type Info
                    $Global:TextBlockStatus.Text = "No users with stale passwords found."
                }
            } catch {
                $ErrorMessage = "Error loading stale passwords: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler fÃ¼r ButtonQuickNeverChangingPasswords
    if ($null -ne $Global:ButtonQuickNeverChangingPasswords) {
        $ButtonQuickNeverChangingPasswords.add_Click({
            Write-ADReportLog -Message "Loading users with never changing passwords..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                $Global:TextBoxFilterValue1.Text = ""
                $Global:TextBoxFilterValue2.Text = ""

                $NeverChangingPasswords = Get-NeverChangingPasswords -Days 365
                if ($NeverChangingPasswords -and $NeverChangingPasswords.Count -gt 0) {
                    Update-ADReportResults -Results $NeverChangingPasswords
                    Write-ADReportLog -Message "Never changing passwords loaded. $($NeverChangingPasswords.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $NeverChangingPasswords
                    $Global:TextBlockStatus.Text = "Never changing passwords loaded. $($NeverChangingPasswords.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No users with never changing passwords found." -Type Info
                    $Global:TextBlockStatus.Text = "No users with never changing passwords found."
                }
            } catch {
                $ErrorMessage = "Error loading never changing passwords: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler fÃ¼r ButtonQuickEmptyGroups
    if ($null -ne $Global:ButtonQuickEmptyGroups) {
        $ButtonQuickEmptyGroups.add_Click({
            Write-ADReportLog -Message "Loading empty groups..." -Type Info
            try {
                $Global:RadioButtonGroup.IsChecked = $true
                
                $EmptyGroups = Get-EmptyGroups
                if ($EmptyGroups -and $EmptyGroups.Count -gt 0) {
                    Update-ADReportResults -Results $EmptyGroups
                    Write-ADReportLog -Message "Empty groups loaded. $($EmptyGroups.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $EmptyGroups
                    $Global:TextBlockStatus.Text = "Empty groups loaded. $($EmptyGroups.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No empty groups found." -Type Info
                    $Global:TextBlockStatus.Text = "No empty groups found."
                }
            } catch {
                $ErrorMessage = "Error loading empty groups: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler fÃ¼r ButtonQuickServiceAccountsOverview
    if ($null -ne $Global:ButtonQuickServiceAccountsOverview) {
        $ButtonQuickServiceAccountsOverview.add_Click({
            Write-ADReportLog -Message "Loading service accounts overview..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $ServiceAccounts = Get-ServiceAccountsOverview
                if ($ServiceAccounts -and $ServiceAccounts.Count -gt 0) {
                    Update-ADReportResults -Results $ServiceAccounts
                    Write-ADReportLog -Message "Service accounts loaded. $($ServiceAccounts.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $ServiceAccounts
                    $Global:TextBlockStatus.Text = "Service accounts loaded. $($ServiceAccounts.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No service accounts found." -Type Info
                    $Global:TextBlockStatus.Text = "No service accounts found."
                }
            } catch {
                $ErrorMessage = "Error loading service accounts: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler fÃ¼r ButtonQuickGPOOverview
    if ($null -ne $Global:ButtonQuickGPOOverview) {
        $ButtonQuickGPOOverview.add_Click({
            Write-ADReportLog -Message "Loading GPO overview..." -Type Info
            try {
                $GPOOverview = Get-GPOOverview
                if ($GPOOverview -and $GPOOverview.Count -gt 0) {
                    Update-ADReportResults -Results $GPOOverview
                    Write-ADReportLog -Message "GPO overview loaded. $($GPOOverview.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $GPOOverview
                    $Global:TextBlockStatus.Text = "GPO overview loaded. $($GPOOverview.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No GPOs found." -Type Info
                    $Global:TextBlockStatus.Text = "No GPOs found."
                }
            } catch {
                $ErrorMessage = "Error loading GPO overview: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler fÃ¼r ButtonQuickHoneyTokens
    if ($null -ne $Global:ButtonQuickHoneyTokens) {
        $ButtonQuickHoneyTokens.add_Click({
            Write-ADReportLog -Message "Analyzing potential honey tokens..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $HoneyTokens = Get-HoneyTokens
                if ($HoneyTokens -and $HoneyTokens.Count -gt 0) {
                    Update-ADReportResults -Results $HoneyTokens
                    Write-ADReportLog -Message "Honey token analysis completed. $($HoneyTokens.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $HoneyTokens
                    $Global:TextBlockStatus.Text = "Honey token analysis completed. $($HoneyTokens.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No honey tokens detected." -Type Info
                    $Global:TextBlockStatus.Text = "No honey tokens detected."
                }
            } catch {
                $ErrorMessage = "Error analyzing honey tokens: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler fÃ¼r weitere Quick Report Buttons
    if ($null -ne $Global:ButtonQuickExpiringAccounts) {
        $ButtonQuickExpiringAccounts.add_Click({
            Write-ADReportLog -Message "Analyzing expiring accounts..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $ExpiringAccounts = Get-ExpiringAccounts
                if ($ExpiringAccounts -and $ExpiringAccounts.Count -gt 0) {
                    Update-ADReportResults -Results $ExpiringAccounts
                    Write-ADReportLog -Message "Expiring accounts analysis completed. $($ExpiringAccounts.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $ExpiringAccounts
                    $Global:TextBlockStatus.Text = "Expiring accounts analysis completed. $($ExpiringAccounts.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No expiring accounts found." -Type Info
                    $Global:TextBlockStatus.Text = "No expiring accounts found."
                }
            } catch {
                $ErrorMessage = "Error analyzing expiring accounts: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Definiere den ListBox Selection Changed Handler als Script Variable
    $script:ListBoxSelectionChangedHandler = $null
    
    # Fenster anzeigen
    $null = $Window.ShowDialog()
}

# Funktion zur Visualisierung der Ergebnisse (Placeholder fÃ¼r zukÃ¼nftige Implementierung)
Function Update-ResultVisualization {
    param (
        [Parameter(Mandatory=$false)]
        $Results
    )
    
    # Diese Funktion ist ein Placeholder fÃ¼r zukÃ¼nftige Visualisierungen
    # Momentan wird nur ein Debug-Log-Eintrag erstellt
    Write-DebugLog "Update-ResultVisualization aufgerufen mit $($Results.Count) Ergebnissen" "Visualization"
}

# Platzhalter-Funktion fÃ¼r initiale Visualisierung beim Start
Function Initialize-ResultVisualization {
    [CmdletBinding()]
    param()
    
    # Diese Funktion ist ein Placeholder fÃ¼r zukÃ¼nftige Visualisierungen
    Write-DebugLog "Initialize-ResultVisualization aufgerufen" "Visualization"
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

    # Ruft Initialize-ADReportForm auf, welche die UI lÃ¤dt, Elemente zuweist und fÃ¼llt.
    Initialize-ADReportForm -XAMLContent $Global:XAML
}

# --- Neue Roadmap Features - Benutzer-Reports ---
Function Get-StalePasswords {
    [CmdletBinding()]
    param([int]$Days = 90)
    
    try {
        Write-ADReportLog -Message "Analyzing users with stale passwords (older than $Days days)..." -Type Info -Terminal
        $CutoffDate = (Get-Date).AddDays(-$Days)
        $Users = Get-ADUser -Filter "Enabled -eq `$true" -Properties PasswordLastSet, DisplayName, SamAccountName, Department, Title -ErrorAction Stop
        $StalePasswordUsers = $Users | Where-Object { $_.PasswordLastSet -lt $CutoffDate }
        
        $Results = $StalePasswordUsers | Select-Object DisplayName, SamAccountName, Department, Title, 
            @{Name="PasswordLastSet";Expression={$_.PasswordLastSet}},
            @{Name="DaysSinceLastChange";Expression={(New-TimeSpan -Start $_.PasswordLastSet -End (Get-Date)).Days}}
        
        return $Results
    } catch {
        Write-ADReportLog -Message "Error analyzing stale passwords: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-NeverChangingPasswords {
    [CmdletBinding()]
    param([int]$Days = 365)
    
    try {
        Write-ADReportLog -Message "Analyzing users with passwords that never change or are very old (older than $Days days)..." -Type Info -Terminal
        $CutoffDate = (Get-Date).AddDays(-$Days)
        $Users = Get-ADUser -Filter "Enabled -eq `$true" -Properties PasswordLastSet, WhenCreated, DisplayName, SamAccountName, Department, Title, PasswordNeverExpires -ErrorAction Stop
        
        $NeverChangedUsers = $Users | Where-Object { 
            # Passwort nie geÃ¤ndert seit Erstellung
            ($_.PasswordLastSet -eq $_.WhenCreated) -or 
            # Passwort ist null
            ($_.PasswordLastSet -eq $null) -or
            # Passwort ist sehr alt
            ($_.PasswordLastSet -lt $CutoffDate) -or
            # Passwort lÃ¤uft nie ab und ist alt
            ($_.PasswordNeverExpires -eq $true -and $_.PasswordLastSet -lt $CutoffDate)
        }
        
        $Results = foreach ($user in $NeverChangedUsers) {
            $daysSinceChange = if ($user.PasswordLastSet) { 
                (New-TimeSpan -Start $user.PasswordLastSet -End (Get-Date)).Days 
            } else { 
                if ($user.WhenCreated) { (New-TimeSpan -Start $user.WhenCreated -End (Get-Date)).Days } else { 9999 }
            }
            
            [PSCustomObject]@{
                DisplayName = $user.DisplayName
                SamAccountName = $user.SamAccountName
                Department = $user.Department
                Title = $user.Title
                WhenCreated = $user.WhenCreated
                PasswordLastSet = if ($user.PasswordLastSet) { $user.PasswordLastSet } else { "Never" }
                PasswordNeverExpires = $user.PasswordNeverExpires
                DaysSinceLastChange = $daysSinceChange
                RiskLevel = if ($daysSinceChange -gt 730) { "Critical" } elseif ($daysSinceChange -gt 365) { "High" } else { "Medium" }
            }
        }
        
        Write-ADReportLog -Message "Never changing passwords analysis completed. $($Results.Count) accounts found." -Type Info -Terminal
        return $Results | Sort-Object DaysSinceLastChange -Descending
    } catch {
        Write-ADReportLog -Message "Error analyzing never changing passwords: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-ExpiringAccounts {
    [CmdletBinding()]
    param([int]$Days = 30)
    
    try {
        Write-ADReportLog -Message "Analyzing accounts expiring within $Days days..." -Type Info -Terminal
        $FutureDate = (Get-Date).AddDays($Days)
        $Users = Get-ADUser -Filter "Enabled -eq `$true -and AccountExpirationDate -like '*'" -Properties AccountExpirationDate, DisplayName, SamAccountName, Department -ErrorAction Stop
        $ExpiringUsers = $Users | Where-Object { $_.AccountExpirationDate -le $FutureDate -and $_.AccountExpirationDate -gt (Get-Date) }
        
        $Results = $ExpiringUsers | Select-Object DisplayName, SamAccountName, Department,
            @{Name="AccountExpirationDate";Expression={$_.AccountExpirationDate}},
            @{Name="DaysUntilExpiration";Expression={(New-TimeSpan -Start (Get-Date) -End $_.AccountExpirationDate).Days}}
        
        return $Results
    } catch {
        Write-ADReportLog -Message "Error analyzing expiring accounts: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-ReversibleEncryptionUsers {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing users with reversible encryption..." -Type Info -Terminal
        $Users = Get-ADUser -Filter "Enabled -eq `$true" -Properties AllowReversiblePasswordEncryption, DisplayName, SamAccountName, Department -ErrorAction Stop
        $ReversibleUsers = $Users | Where-Object { $_.AllowReversiblePasswordEncryption -eq $true }
        
        $Results = $ReversibleUsers | Select-Object DisplayName, SamAccountName, Department,
            @{Name="AllowReversiblePasswordEncryption";Expression={$_.AllowReversiblePasswordEncryption}}
        
        return $Results
    } catch {
        Write-ADReportLog -Message "Error analyzing reversible encryption users: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-KerberosDESUsers {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing users with Kerberos DES encryption..." -Type Info -Terminal
        $Users = Get-ADUser -Filter "Enabled -eq `$true" -Properties UserAccountControl, DisplayName, SamAccountName, Department -ErrorAction Stop
        $DESUsers = $Users | Where-Object { ($_.UserAccountControl -band 0x200000) -ne 0 }
        
        $Results = $DESUsers | Select-Object DisplayName, SamAccountName, Department,
            @{Name="UsesDESEncryption";Expression={"True"}}
        
        return $Results
    } catch {
        Write-ADReportLog -Message "Error analyzing Kerberos DES users: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-UsersWithSPN {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing users with Service Principal Names..." -Type Info -Terminal
        $Users = Get-ADUser -Filter "Enabled -eq `$true -and ServicePrincipalName -like '*'" -Properties ServicePrincipalName, DisplayName, SamAccountName, Department -ErrorAction Stop
        
        $Results = $Users | Select-Object DisplayName, SamAccountName, Department,
            @{Name="ServicePrincipalNames";Expression={$_.ServicePrincipalName -join "; "}}
        
        return $Results
    } catch {
        Write-ADReportLog -Message "Error analyzing users with SPN: $($_.Exception.Message)" -Type Error
        return @()
    }
}

# --- Weitere User Report Funktionen ---
Function Get-GuestAccountStatus {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing Guest account status..." -Type Info -Terminal
        
        # Suche nach dem Guest Account (kann verschiedene Namen haben)
        $GuestAccounts = @()
        
        # Standard Guest Account Namen in verschiedenen Sprachen
        $GuestNames = @("Guest", "Gast", "InvitÃ©", "Invitado", "Ospite")
        
        foreach ($name in $GuestNames) {
            $account = Get-ADUser -Filter "SamAccountName -eq '$name'" -Properties * -ErrorAction SilentlyContinue
            if ($account) {
                $GuestAccounts += $account
            }
        }
        
        # Auch nach SID suchen (Well-known Guest SID endet mit -501)
        $DomainSID = (Get-ADDomain).DomainSID.Value
        $GuestSID = "$DomainSID-501"
        
        try {
            $GuestBySID = Get-ADUser -Identity $GuestSID -Properties * -ErrorAction SilentlyContinue
            if ($GuestBySID -and -not ($GuestAccounts | Where-Object { $_.SID -eq $GuestBySID.SID })) {
                $GuestAccounts += $GuestBySID
            }
        } catch {
            # SID nicht gefunden ist OK
        }
        
        if ($GuestAccounts.Count -eq 0) {
            Write-ADReportLog -Message "No Guest account found in domain." -Type Info -Terminal
            return @([PSCustomObject]@{
                AccountName = "No Guest Account"
                Status = "Not Found"
                SecurityRisk = "Low"
                Recommendation = "Guest account not present in domain (good security practice)"
            })
        }
        
        $Results = foreach ($guest in $GuestAccounts) {
            # Analysiere Sicherheitsrisiken
            $risks = @()
            $riskLevel = "Low"
            $recommendations = @()
            
            if ($guest.Enabled) {
                $risks += "Account is enabled"
                $riskLevel = "High"
                $recommendations += "Disable Guest account immediately"
            }
            
            if ($guest.PasswordNeverExpires) {
                $risks += "Password never expires"
                if ($riskLevel -ne "High") { $riskLevel = "Medium" }
                $recommendations += "Set password expiration"
            }
            
            if ($guest.PasswordNotRequired) {
                $risks += "Password not required"
                $riskLevel = "Critical"
                $recommendations += "Enforce password requirement"
            }
            
            if ($guest.LastLogonDate -and $guest.LastLogonDate -gt (Get-Date).AddDays(-30)) {
                $risks += "Recently used (within 30 days)"
                if ($riskLevel -eq "Low") { $riskLevel = "Medium" }
                $recommendations += "Investigate recent usage"
            }
            
            # PrÃ¼fe Gruppenmitgliedschaften
            $groups = Get-ADPrincipalGroupMembership -Identity $guest -ErrorAction SilentlyContinue
            $privilegedGroups = @()
            
            foreach ($group in $groups) {
                if ($group.Name -ne "Domain Guests" -and $group.Name -ne "Guests") {
                    $privilegedGroups += $group.Name
                }
            }
            
            if ($privilegedGroups.Count -gt 0) {
                $risks += "Member of additional groups"
                $riskLevel = "Critical"
                $recommendations += "Remove from all groups except Domain Guests"
            }
            
            [PSCustomObject]@{
                AccountName = $guest.SamAccountName
                DisplayName = $guest.DisplayName
                Enabled = $guest.Enabled
                PasswordLastSet = $guest.PasswordLastSet
                LastLogonDate = $guest.LastLogonDate
                PasswordNeverExpires = $guest.PasswordNeverExpires
                PasswordNotRequired = $guest.PasswordNotRequired
                AccountLockedOut = $guest.LockedOut
                SID = $guest.SID
                Groups = if ($privilegedGroups) { $privilegedGroups -join ", " } else { "Domain Guests only" }
                RiskLevel = $riskLevel
                SecurityIssues = if ($risks) { $risks -join "; " } else { "None" }
                Recommendations = if ($recommendations) { $recommendations -join "; " } else { "Account properly secured" }
                WhenCreated = $guest.WhenCreated
                WhenChanged = $guest.WhenChanged
            }
        }
        
        Write-ADReportLog -Message "Guest account analysis completed. $($Results.Count) account(s) found." -Type Info -Terminal
        return $Results
        
    } catch {
        Write-ADReportLog -Message "Error analyzing Guest account status: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-UsersByDepartment {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing users by department..." -Type Info -Terminal
        
        # Lade alle Benutzer mit Department-Attribut
        $Users = Get-ADUser -Filter * -Properties Department, DisplayName, SamAccountName, Title, Enabled, LastLogonDate, Manager -ErrorAction Stop
        
        # Gruppiere nach Department
        $DepartmentGroups = $Users | Group-Object Department
        
        $Results = foreach ($deptGroup in $DepartmentGroups) {
            $deptName = if ([string]::IsNullOrWhiteSpace($deptGroup.Name)) { "(No Department)" } else { $deptGroup.Name }
            
            # Statistiken fÃ¼r die Abteilung
            $enabledCount = ($deptGroup.Group | Where-Object { $_.Enabled }).Count
            $disabledCount = ($deptGroup.Group | Where-Object { -not $_.Enabled }).Count
            $activeCount = ($deptGroup.Group | Where-Object { $_.LastLogonDate -and $_.LastLogonDate -gt (Get-Date).AddDays(-30) }).Count
            
            # Manager in der Abteilung identifizieren
            $managers = @()
            $managersDict = @{}
            
            foreach ($user in $deptGroup.Group) {
                if ($user.Manager) {
                    if (-not $managersDict.ContainsKey($user.Manager)) {
                        try {
                            $managerObj = Get-ADUser -Identity $user.Manager -Properties DisplayName -ErrorAction SilentlyContinue
                            if ($managerObj) {
                                $managersDict[$user.Manager] = $managerObj.DisplayName
                            }
                        } catch {
                            # Manager nicht gefunden
                        }
                    }
                }
            }
            
            $uniqueManagers = $managersDict.Values | Sort-Object -Unique
            
            # Erstelle EintrÃ¤ge fÃ¼r jeden Benutzer in der Abteilung
            foreach ($user in $deptGroup.Group) {
                $managerName = "None"
                if ($user.Manager -and $managersDict.ContainsKey($user.Manager)) {
                    $managerName = $managersDict[$user.Manager]
                }
                
                [PSCustomObject]@{
                    Department = $deptName
                    DisplayName = $user.DisplayName
                    SamAccountName = $user.SamAccountName
                    Title = $user.Title
                    Manager = $managerName
                    Enabled = $user.Enabled
                    LastLogonDate = $user.LastLogonDate
                    DeptUserCount = $deptGroup.Count
                    DeptEnabledCount = $enabledCount
                    DeptActiveCount = $activeCount
                    ActivityStatus = if ($user.LastLogonDate -and $user.LastLogonDate -gt (Get-Date).AddDays(-30)) { "Active" } 
                                    elseif ($user.LastLogonDate -and $user.LastLogonDate -gt (Get-Date).AddDays(-90)) { "Inactive" } 
                                    else { "Very Inactive" }
                }
            }
        }
        
        Write-ADReportLog -Message "Users by department analysis completed. $($Results.Count) users in $($DepartmentGroups.Count) departments." -Type Info -Terminal
        return $Results | Sort-Object Department, DisplayName
        
    } catch {
        Write-ADReportLog -Message "Error analyzing users by department: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-UsersByManager {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing users by manager..." -Type Info -Terminal
        
        # Lade alle Benutzer mit Manager-Attribut
        $Users = Get-ADUser -Filter * -Properties Manager, DisplayName, SamAccountName, Department, Title, Enabled, LastLogonDate -ErrorAction Stop
        
        # Erstelle Dictionary fÃ¼r Manager-Namen
        $managersDict = @{}
        $usersWithManager = $Users | Where-Object { $_.Manager }
        
        foreach ($user in $usersWithManager) {
            if (-not $managersDict.ContainsKey($user.Manager)) {
                try {
                    $managerObj = Get-ADUser -Identity $user.Manager -Properties DisplayName, Department, Title -ErrorAction SilentlyContinue
                    if ($managerObj) {
                        $managersDict[$user.Manager] = $managerObj
                    }
                } catch {
                    # Manager nicht gefunden
                }
            }
        }
        
        # Gruppiere nach Manager
        $ManagerGroups = $usersWithManager | Group-Object Manager
        
        $Results = @()
        
        # Verarbeite Benutzer mit Manager
        foreach ($mgrGroup in $ManagerGroups) {
            $managerInfo = $managersDict[$mgrGroup.Name]
            $managerName = if ($managerInfo) { $managerInfo.DisplayName } else { "Unknown Manager" }
            $managerDept = if ($managerInfo) { $managerInfo.Department } else { "Unknown" }
            $managerTitle = if ($managerInfo) { $managerInfo.Title } else { "Unknown" }
            
            # Statistiken fÃ¼r diesen Manager
            $enabledCount = ($mgrGroup.Group | Where-Object { $_.Enabled }).Count
            $activeCount = ($mgrGroup.Group | Where-Object { $_.LastLogonDate -and $_.LastLogonDate -gt (Get-Date).AddDays(-30) }).Count
            
            foreach ($user in $mgrGroup.Group) {
                $Results += [PSCustomObject]@{
                    ManagerName = $managerName
                    ManagerDepartment = $managerDept
                    ManagerTitle = $managerTitle
                    DirectReports = $mgrGroup.Count
                    ActiveReports = $activeCount
                    UserDisplayName = $user.DisplayName
                    UserSamAccountName = $user.SamAccountName
                    UserDepartment = $user.Department
                    UserTitle = $user.Title
                    UserEnabled = $user.Enabled
                    UserLastLogon = $user.LastLogonDate
                    UserActivityStatus = if ($user.LastLogonDate -and $user.LastLogonDate -gt (Get-Date).AddDays(-30)) { "Active" } 
                                        elseif ($user.LastLogonDate -and $user.LastLogonDate -gt (Get-Date).AddDays(-90)) { "Inactive" } 
                                        else { "Very Inactive" }
                }
            }
        }
        
        # FÃ¼ge Benutzer ohne Manager hinzu
        $usersWithoutManager = $Users | Where-Object { -not $_.Manager }
        foreach ($user in $usersWithoutManager) {
            $Results += [PSCustomObject]@{
                ManagerName = "(No Manager)"
                ManagerDepartment = "N/A"
                ManagerTitle = "N/A"
                DirectReports = 0
                ActiveReports = 0
                UserDisplayName = $user.DisplayName
                UserSamAccountName = $user.SamAccountName
                UserDepartment = $user.Department
                UserTitle = $user.Title
                UserEnabled = $user.Enabled
                UserLastLogon = $user.LastLogonDate
                UserActivityStatus = if ($user.LastLogonDate -and $user.LastLogonDate -gt (Get-Date).AddDays(-30)) { "Active" } 
                                    elseif ($user.LastLogonDate -and $user.LastLogonDate -gt (Get-Date).AddDays(-90)) { "Inactive" } 
                                    else { "Very Inactive" }
            }
        }
        
        Write-ADReportLog -Message "Users by manager analysis completed. $($Results.Count) users analyzed." -Type Info -Terminal
        return $Results | Sort-Object ManagerName, UserDisplayName
        
    } catch {
        Write-ADReportLog -Message "Error analyzing users by manager: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-RemoteAccessUsers {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing remote access users..." -Type Info -Terminal
        
        # Remote Access kann durch verschiedene Gruppenmitgliedschaften ermÃ¶glicht werden
        $RemoteAccessGroups = @(
            "Remote Desktop Users",
            "Remotedesktopbenutzer", # Deutsch
            "VPN Users",
            "Remote Access Users",
            "Remote Management Users",
            "Terminal Server Users"
        )
        
        $RemoteAccessUsers = @()
        $ProcessedUsers = @{}
        
        # Suche nach Benutzern in Remote Access Gruppen
        foreach ($groupName in $RemoteAccessGroups) {
            try {
                $group = Get-ADGroup -Filter "Name -eq '$groupName'" -ErrorAction SilentlyContinue
                if ($group) {
                    $members = Get-ADGroupMember -Identity $group -Recursive -ErrorAction SilentlyContinue | 
                               Where-Object { $_.objectClass -eq "user" }
                    
                    foreach ($member in $members) {
                        if (-not $ProcessedUsers.ContainsKey($member.SamAccountName)) {
                            $userDetails = Get-ADUser -Identity $member.SamAccountName `
                                -Properties DisplayName, Department, Title, Enabled, LastLogonDate, `
                                           PasswordLastSet, AccountExpirationDate, Description, `
                                           msNPAllowDialin, msRADIUSFramedIPAddress -ErrorAction SilentlyContinue
                            
                            if ($userDetails) {
                                $ProcessedUsers[$member.SamAccountName] = @{
                                    User = $userDetails
                                    Groups = @($groupName)
                                }
                            }
                        } else {
                            $ProcessedUsers[$member.SamAccountName].Groups += $groupName
                        }
                    }
                }
            } catch {
                Write-ADReportLog -Message "Could not process group $groupName : $($_.Exception.Message)" -Type Warning
            }
        }
        
        # Suche auch nach Benutzern mit Dial-In Berechtigung
        try {
            $dialInUsers = Get-ADUser -Filter "msNPAllowDialin -eq `$true" `
                -Properties DisplayName, Department, Title, Enabled, LastLogonDate, `
                           PasswordLastSet, AccountExpirationDate, Description, `
                           msNPAllowDialin, msRADIUSFramedIPAddress -ErrorAction SilentlyContinue
            
            foreach ($user in $dialInUsers) {
                if (-not $ProcessedUsers.ContainsKey($user.SamAccountName)) {
                    $ProcessedUsers[$user.SamAccountName] = @{
                        User = $user
                        Groups = @("(Dial-In Permission)")
                    }
                } else {
                    $ProcessedUsers[$user.SamAccountName].Groups += "(Dial-In Permission)"
                }
            }
        } catch {
            Write-ADReportLog -Message "Could not check dial-in permissions: $($_.Exception.Message)" -Type Warning
        }
        
        # Erstelle Ergebnisse
        foreach ($entry in $ProcessedUsers.GetEnumerator()) {
            $user = $entry.Value.User
            $groups = $entry.Value.Groups | Sort-Object -Unique
            
            # Risikobewertung
            $riskLevel = "Low"
            $riskFactors = @()
            
            if (-not $user.Enabled) {
                $riskFactors += "Account disabled but has remote access"
                $riskLevel = "Medium"
            }
            
            if ($user.AccountExpirationDate -and $user.AccountExpirationDate -lt (Get-Date)) {
                $riskFactors += "Account expired"
                $riskLevel = "High"
            }
            
            if ($user.PasswordLastSet -and $user.PasswordLastSet -lt (Get-Date).AddDays(-180)) {
                $riskFactors += "Password older than 180 days"
                if ($riskLevel -eq "Low") { $riskLevel = "Medium" }
            }
            
            if (-not $user.LastLogonDate -or $user.LastLogonDate -lt (Get-Date).AddDays(-90)) {
                $riskFactors += "Inactive for >90 days"
                if ($riskLevel -eq "Low") { $riskLevel = "Medium" }
            }
            
            $RemoteAccessUsers += [PSCustomObject]@{
                DisplayName = $user.DisplayName
                SamAccountName = $user.SamAccountName
                Department = $user.Department
                Title = $user.Title
                Enabled = $user.Enabled
                RemoteAccessGroups = $groups -join ", "
                LastLogonDate = $user.LastLogonDate
                PasswordLastSet = $user.PasswordLastSet
                AccountExpiration = $user.AccountExpirationDate
                DialInPermission = if ($user.msNPAllowDialin) { "Allowed" } else { "Not Set" }
                RiskLevel = $riskLevel
                RiskFactors = if ($riskFactors) { $riskFactors -join "; " } else { "None" }
                ActivityStatus = if ($user.LastLogonDate -and $user.LastLogonDate -gt (Get-Date).AddDays(-30)) { "Active" } 
                                elseif ($user.LastLogonDate -and $user.LastLogonDate -gt (Get-Date).AddDays(-90)) { "Inactive" } 
                                else { "Very Inactive" }
                Description = $user.Description
            }
        }
        
        if ($RemoteAccessUsers.Count -eq 0) {
            Write-ADReportLog -Message "No remote access users found." -Type Info -Terminal
            return @([PSCustomObject]@{
                DisplayName = "No Results"
                SamAccountName = "N/A"
                RemoteAccessGroups = "No remote access users found"
                RiskLevel = "N/A"
                ActivityStatus = "N/A"
            })
        }
        
        Write-ADReportLog -Message "Remote access users analysis completed. $($RemoteAccessUsers.Count) users found." -Type Info -Terminal
        return $RemoteAccessUsers | Sort-Object RiskLevel, DisplayName
        
    } catch {
        Write-ADReportLog -Message "Error analyzing remote access users: $($_.Exception.Message)" -Type Error
        return @()
    }
}

# --- Neue Roadmap Features - Gruppen-Reports ---
Function Get-EmptyGroups {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing empty groups..." -Type Info -Terminal
        $Groups = Get-ADGroup -Filter * -Properties Members, Name, GroupCategory, GroupScope, Description, whenCreated, whenChanged -ErrorAction Stop
        
        $EmptyGroups = $Groups | Where-Object { 
            # PrÃ¼fe sowohl direkte Members als auch Ã¼ber Get-ADGroupMember
            $_.Members.Count -eq 0 -and 
            @(Get-ADGroupMember -Identity $_.DistinguishedName -ErrorAction SilentlyContinue).Count -eq 0
        }
        
        $Results = foreach ($group in $EmptyGroups) {
            $ageInDays = if ($group.whenCreated) { (New-TimeSpan -Start $group.whenCreated -End (Get-Date)).Days } else { 0 }
            
            [PSCustomObject]@{
                Name = $group.Name
                GroupCategory = $group.GroupCategory
                GroupScope = $group.GroupScope
                Description = $group.Description
                WhenCreated = $group.whenCreated
                WhenChanged = $group.whenChanged
                AgeInDays = $ageInDays
                MemberCount = 0
                CleanupRecommended = if ($ageInDays -gt 90 -and [string]::IsNullOrWhiteSpace($group.Description)) { "Yes" } else { "Review" }
            }
        }
        
        Write-ADReportLog -Message "Empty groups analysis completed. $($Results.Count) groups found." -Type Info -Terminal
        return $Results | Sort-Object AgeInDays -Descending
    } catch {
        Write-ADReportLog -Message "Error analyzing empty groups: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-NestedGroups {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing nested group memberships..." -Type Info -Terminal
        $Groups = Get-ADGroup -Filter * -Properties MemberOf, Name, GroupCategory -ErrorAction Stop
        $NestedGroups = $Groups | Where-Object { $_.MemberOf.Count -gt 0 }
        
        $Results = $NestedGroups | Select-Object Name, GroupCategory,
            @{Name="MemberOfGroups";Expression={$_.MemberOf.Count}},
            @{Name="ParentGroups";Expression={($_.MemberOf | ForEach-Object { (Get-ADGroup $_).Name }) -join "; "}}
        
        return $Results
    } catch {
        Write-ADReportLog -Message "Error analyzing nested groups: $($_.Exception.Message)" -Type Error
        return @()
    }
}

# --- Neue Roadmap Features - Computer-Reports ---
Function Get-OSSummary {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Generating Operating System summary..." -Type Info -Terminal
        $Computers = Get-ADComputer -Filter "Enabled -eq `$true" -Properties OperatingSystem, OperatingSystemVersion -ErrorAction Stop
        
        $OSGroups = $Computers | Group-Object OperatingSystem | Sort-Object Count -Descending
        $Results = $OSGroups | Select-Object @{Name="OperatingSystem";Expression={$_.Name}},
            @{Name="Count";Expression={$_.Count}},
            @{Name="Percentage";Expression={[math]::Round(($_.Count / $Computers.Count) * 100, 2)}}
        
        return $Results
    } catch {
        Write-ADReportLog -Message "Error generating OS summary: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-BitLockerStatus {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing BitLocker status (placeholder - requires additional modules)..." -Type Info -Terminal
        # Placeholder - wÃ¼rde BitLocker-spezifische Abfragen erfordern
        $Results = @()
        Write-ADReportLog -Message "BitLocker analysis requires additional PowerShell modules and permissions." -Type Warning
        return $Results
    } catch {
        Write-ADReportLog -Message "Error analyzing BitLocker status: $($_.Exception.Message)" -Type Error
        return @()
    }
}

# --- Neue Roadmap Features - Service Account Reports ---
Function Get-ServiceAccountsOverview {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Generating comprehensive Service Accounts overview..." -Type Info -Terminal
        
        # Service Accounts identifizieren durch verschiedene Kriterien
        $PotentialServiceAccounts = @()
        
        # 1. Accounts mit Service-Ã¤hnlichen Namen
        $NameBasedSvcAccounts = Get-ADUser -Filter "(Name -like '*svc*') -or (Name -like '*service*') -or (Name -like '*sql*') -or (Name -like '*iis*') -or (Name -like '*app*') -or (Name -like '*web*')" -Properties Description, LastLogonDate, PasswordLastSet, ServicePrincipalName, PasswordNeverExpires, Enabled, whenCreated, Department -ErrorAction SilentlyContinue
        $PotentialServiceAccounts += $NameBasedSvcAccounts
        
        # 2. Accounts mit SPNs (Service Principal Names)
        $SPNAccounts = Get-ADUser -Filter "ServicePrincipalName -like '*'" -Properties Description, LastLogonDate, PasswordLastSet, ServicePrincipalName, PasswordNeverExpires, Enabled, whenCreated, Department -ErrorAction SilentlyContinue
        $PotentialServiceAccounts += $SPNAccounts
        
        # 3. Accounts mit service-bezogenen Beschreibungen
        $DescBasedSvcAccounts = Get-ADUser -Filter "Description -like '*service*' -or Description -like '*application*' -or Description -like '*sql*' -or Description -like '*database*'" -Properties Description, LastLogonDate, PasswordLastSet, ServicePrincipalName, PasswordNeverExpires, Enabled, whenCreated, Department -ErrorAction SilentlyContinue
        $PotentialServiceAccounts += $DescBasedSvcAccounts
        
        # Duplikate entfernen
        $UniqueServiceAccounts = $PotentialServiceAccounts | Sort-Object DistinguishedName | Get-Unique -AsString
        
        $Results = foreach ($account in $UniqueServiceAccounts) {
            $daysSincePasswordChange = if ($account.PasswordLastSet) { 
                (New-TimeSpan -Start $account.PasswordLastSet -End (Get-Date)).Days 
            } else { 9999 }
            
            $daysSinceLastLogon = if ($account.LastLogonDate) { 
                (New-TimeSpan -Start $account.LastLogonDate -End (Get-Date)).Days 
            } else { 9999 }
            
            # Risk Assessment
            $riskScore = 0
            $riskFactors = @()
            
            if ($daysSincePasswordChange -gt 365) { $riskScore += 3; $riskFactors += "Old Password" }
            if ($account.PasswordNeverExpires) { $riskScore += 2; $riskFactors += "Password Never Expires" }
            if ($daysSinceLastLogon -gt 90) { $riskScore += 2; $riskFactors += "Inactive" }
            if ($account.ServicePrincipalName.Count -gt 0) { $riskScore += 1; $riskFactors += "Has SPN" }
            if ($account.Enabled -eq $false) { $riskScore -= 2; $riskFactors += "Disabled" }
            
            $riskLevel = switch ($riskScore) {
                {$_ -ge 5} { "High" }
                {$_ -ge 3} { "Medium" }
                {$_ -ge 1} { "Low" }
                default { "Minimal" }
            }
            
            [PSCustomObject]@{
                Name = $account.Name
                SamAccountName = $account.SamAccountName
                Description = $account.Description
                Department = $account.Department
                Enabled = $account.Enabled
                LastLogonDate = if ($account.LastLogonDate) { $account.LastLogonDate } else { "Never" }
                DaysSinceLastLogon = $daysSinceLastLogon
                PasswordLastSet = if ($account.PasswordLastSet) { $account.PasswordLastSet } else { "Never" }
                DaysSincePasswordChange = $daysSincePasswordChange
                PasswordNeverExpires = $account.PasswordNeverExpires
                HasSPN = $account.ServicePrincipalName.Count -gt 0
                SPNCount = $account.ServicePrincipalName.Count
                RiskLevel = $riskLevel
                RiskFactors = $riskFactors -join ", "
                WhenCreated = $account.whenCreated
            }
        }
        
        Write-ADReportLog -Message "Service Accounts overview completed. $($Results.Count) accounts found." -Type Info -Terminal
        return $Results | Sort-Object RiskLevel, DaysSincePasswordChange -Descending
    } catch {
        Write-ADReportLog -Message "Error analyzing service accounts: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-ManagedServiceAccounts {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing Managed Service Accounts..." -Type Info -Terminal
        # gMSA und sMSA haben spezielle ObjectClasses
        $MSAs = Get-ADObject -Filter "ObjectClass -eq 'msDS-ManagedServiceAccount' -or ObjectClass -eq 'msDS-GroupManagedServiceAccount'" -Properties Name, ObjectClass, LastLogonDate -ErrorAction Stop
        
        $Results = $MSAs | Select-Object Name,
            @{Name="AccountType";Expression={if($_.ObjectClass -eq "msDS-GroupManagedServiceAccount"){"Group MSA"}else{"Standalone MSA"}}},
            @{Name="LastLogonDate";Expression={$_.LastLogonDate}}
        
        return $Results
    } catch {
        Write-ADReportLog -Message "Error analyzing managed service accounts: $($_.Exception.Message)" -Type Error
        return @()
    }
}

# --- Platzhalter fÃ¼r weitere neue Features ---
Function Get-HoneyTokens {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing potential honey token accounts and suspicious activities..." -Type Info -Terminal
        
        # Honey Tokens sind oft spezielle Accounts die zur Angriffserkennung verwendet werden
        # Wir suchen nach Accounts mit verdÃ¤chtigen Eigenschaften oder Namensmustern
        
        $SuspiciousPatterns = @(
            "*honey*", "*canary*", "*trap*", "*decoy*", "*bait*", 
            "*test*", "*dummy*", "*fake*", "*monitor*", "*audit*"
        )
        
        $PotentialHoneyTokens = @()
        
        # 1. Accounts mit verdÃ¤chtigen Namen
        foreach ($pattern in $SuspiciousPatterns) {
            $accounts = Get-ADUser -Filter "Name -like '$pattern'" -Properties Description, LastLogonDate, PasswordLastSet, Enabled, whenCreated, Department -ErrorAction SilentlyContinue
            $PotentialHoneyTokens += $accounts
        }
        
        # 2. Accounts mit verdÃ¤chtigen Beschreibungen
        $DescriptionBasedAccounts = Get-ADUser -Filter "Description -like '*honey*' -or Description -like '*canary*' -or Description -like '*monitoring*' -or Description -like '*security*'" -Properties Description, LastLogonDate, PasswordLastSet, Enabled, whenCreated, Department -ErrorAction SilentlyContinue
        $PotentialHoneyTokens += $DescriptionBasedAccounts
        
        # 3. Niemals verwendete Admin-Ã¤hnliche Accounts (verdÃ¤chtig)
        $UnusedAdminAccounts = Get-ADUser -Filter "Name -like '*admin*' -and LastLogonDate -eq `$null -and Enabled -eq `$true" -Properties Description, LastLogonDate, PasswordLastSet, Enabled, whenCreated, Department -ErrorAction SilentlyContinue
        $PotentialHoneyTokens += $UnusedAdminAccounts
        
        # Duplikate entfernen
        $UniqueAccounts = $PotentialHoneyTokens | Sort-Object DistinguishedName | Get-Unique -AsString
        
        $Results = foreach ($account in $UniqueAccounts) {
            $ageInDays = if ($account.whenCreated) { (New-TimeSpan -Start $account.whenCreated -End (Get-Date)).Days } else { 0 }
            $daysSinceLastLogon = if ($account.LastLogonDate) { 
                (New-TimeSpan -Start $account.LastLogonDate -End (Get-Date)).Days 
            } else { 9999 }
            
            # VerdÃ¤chtigkeits-Score
            $suspicionScore = 0
            $indicators = @()
            
            # Name-basierte Indikatoren
            if ($account.Name -match "(honey|canary|trap|decoy|bait)") { $suspicionScore += 5; $indicators += "Suspicious Name Pattern" }
            if ($account.Name -match "(test|dummy|fake)") { $suspicionScore += 3; $indicators += "Test Account Pattern" }
            if ($account.Name -match "(monitor|audit)") { $suspicionScore += 4; $indicators += "Monitoring Pattern" }
            
            # Verhalten-basierte Indikatoren
            if ($account.LastLogonDate -eq $null -and $ageInDays -gt 30) { $suspicionScore += 3; $indicators += "Never Used" }
            if ($account.Enabled -and $daysSinceLastLogon -gt 180) { $suspicionScore += 2; $indicators += "Long Inactive" }
            if ([string]::IsNullOrWhiteSpace($account.Department)) { $suspicionScore += 1; $indicators += "No Department" }
            
            # Beschreibung-basierte Indikatoren
            if ($account.Description -match "(honey|canary|monitoring|security)") { $suspicionScore += 4; $indicators += "Security Description" }
            
            $suspicionLevel = switch ($suspicionScore) {
                {$_ -ge 7} { "Very High" }
                {$_ -ge 5} { "High" }
                {$_ -ge 3} { "Medium" }
                {$_ -ge 1} { "Low" }
                default { "Minimal" }
            }
            
            [PSCustomObject]@{
                Name = $account.Name
                SamAccountName = $account.SamAccountName
                Description = $account.Description
                Department = $account.Department
                Enabled = $account.Enabled
                WhenCreated = $account.whenCreated
                LastLogonDate = if ($account.LastLogonDate) { $account.LastLogonDate } else { "Never" }
                DaysSinceLastLogon = $daysSinceLastLogon
                AgeInDays = $ageInDays
                SuspicionLevel = $suspicionLevel
                SuspicionScore = $suspicionScore
                Indicators = $indicators -join ", "
                PotentialHoneyToken = if ($suspicionScore -ge 5) { "Likely" } elseif ($suspicionScore -ge 3) { "Possible" } else { "Unlikely" }
            }
        }
        
        if ($Results.Count -eq 0) {
            Write-ADReportLog -Message "No potential honey token accounts detected." -Type Info -Terminal
            return @([PSCustomObject]@{
                Name = "No Results"
                SamAccountName = "N/A"
                Description = "No potential honey token accounts found in current environment"
                PotentialHoneyToken = "None Detected"
                SuspicionLevel = "N/A"
                Indicators = "Analysis completed - no suspicious patterns detected"
            })
        }
        
        Write-ADReportLog -Message "Honey token analysis completed. $($Results.Count) potential accounts found." -Type Info -Terminal
        return $Results | Sort-Object SuspicionScore, SuspicionLevel -Descending
    } catch {
        Write-ADReportLog -Message "Error in honey token detection: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-GPOOverview {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Generating comprehensive GPO overview..." -Type Info -Terminal
        
        # PrÃ¼fe ob GroupPolicy Module verfÃ¼gbar ist
        if (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
            Write-ADReportLog -Message "GroupPolicy PowerShell module not available. Attempting alternative approach..." -Type Warning
            
            # Alternative: Ãœber AD direkt abfragen
            $GPOs = Get-ADObject -SearchBase "CN=Policies,CN=System,$((Get-ADDomain).DistinguishedName)" -Filter "ObjectClass -eq 'groupPolicyContainer'" -Properties DisplayName, whenCreated, whenChanged, gPCFileSysPath -ErrorAction Stop
            
            $Results = foreach ($gpo in $GPOs) {
                [PSCustomObject]@{
                    DisplayName = $gpo.DisplayName
                    Id = $gpo.Name
                    CreationTime = $gpo.whenCreated
                    ModificationTime = $gpo.whenChanged
                    GPOStatus = "Unknown (Manual Check Required)"
                    FileSysPath = $gpo.gPCFileSysPath
                    AgeInDays = if ($gpo.whenCreated) { (New-TimeSpan -Start $gpo.whenCreated -End (Get-Date)).Days } else { 0 }
                    DaysSinceModification = if ($gpo.whenChanged) { (New-TimeSpan -Start $gpo.whenChanged -End (Get-Date)).Days } else { 0 }
                }
            }
            
            return $Results | Sort-Object DaysSinceModification
        }
        
        Import-Module GroupPolicy -ErrorAction Stop
        $GPOs = Get-GPO -All -ErrorAction Stop
        
        $Results = foreach ($gpo in $GPOs) {
            # GPO Links analysieren
            $gpoLinks = @()
            try {
                $gpoReport = Get-GPOReport -Guid $gpo.Id -ReportType Xml -ErrorAction SilentlyContinue
                if ($gpoReport) {
                    # Vereinfachte Link-Analyse
                    $linkCount = ([xml]$gpoReport).GPO.LinksTo.SOMPath.Count
                } else {
                    $linkCount = 0
                }
            } catch {
                $linkCount = 0
            }
            
            $ageInDays = if ($gpo.CreationTime) { (New-TimeSpan -Start $gpo.CreationTime -End (Get-Date)).Days } else { 0 }
            $daysSinceModification = if ($gpo.ModificationTime) { (New-TimeSpan -Start $gpo.ModificationTime -End (Get-Date)).Days } else { 0 }
            
            # Status-Bewertung
            $healthStatus = "Good"
            $issues = @()
            
            if ($daysSinceModification -gt 365) { $issues += "Not modified for over 1 year"; $healthStatus = "Review" }
            if ($linkCount -eq 0) { $issues += "No links found"; $healthStatus = "Orphaned" }
            if ($gpo.GpoStatus -eq "AllSettingsDisabled") { $issues += "All settings disabled"; $healthStatus = "Disabled" }
            
            [PSCustomObject]@{
                DisplayName = $gpo.DisplayName
                Id = $gpo.Id
                CreationTime = $gpo.CreationTime
                ModificationTime = $gpo.ModificationTime
                GPOStatus = $gpo.GpoStatus
                AgeInDays = $ageInDays
                DaysSinceModification = $daysSinceModification
                LinkCount = $linkCount
                HealthStatus = $healthStatus
                Issues = $issues -join "; "
                Domain = $gpo.DomainName
                Owner = $gpo.Owner
            }
        }
        
        Write-ADReportLog -Message "GPO overview completed. $($Results.Count) GPOs found." -Type Info -Terminal
        return $Results | Sort-Object HealthStatus, DaysSinceModification -Descending
    } catch {
        Write-ADReportLog -Message "Error generating GPO overview: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-GDPRCompliance {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "GDPR Compliance analysis requires custom business rules..." -Type Info -Terminal
        $Results = @()
        Write-ADReportLog -Message "GDPR compliance analysis requires custom business rule implementation." -Type Warning
        return $Results
    } catch {
        Write-ADReportLog -Message "Error in GDPR compliance analysis: $($_.Exception.Message)" -Type Error
        return @()
    }
}

# --- Weitere Gruppen-Report Funktionen ---
Function Get-CircularGroups {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing circular group references..." -Type Info -Terminal
        
        # Dictionary zur Speicherung der Gruppenmitgliedschaften
        $groupMemberships = @{}
        $circularReferences = @()
        
        # Lade alle Gruppen mit ihren Mitgliedschaften
        $allGroups = Get-ADGroup -Filter * -Properties MemberOf, Members, Description, GroupCategory, GroupScope -ErrorAction Stop
        
        # Baue Dictionary auf
        foreach ($group in $allGroups) {
            $groupMemberships[$group.DistinguishedName] = @{
                Group = $group
                MemberOf = $group.MemberOf
                Visited = $false
                InStack = $false
            }
        }
        
        # Funktion zur rekursiven Suche nach zirkulären Referenzen
        function Find-CircularReference {
            param(
                [string]$GroupDN,
                [System.Collections.ArrayList]$Path
            )
            
            if (-not $groupMemberships.ContainsKey($GroupDN)) { return }
            
            $groupData = $groupMemberships[$GroupDN]
            
            # Wenn bereits im Stack, haben wir eine zirkuläre Referenz gefunden
            if ($groupData.InStack) {
                $circularPath = $Path.Clone()
                $circularPath.Add($GroupDN) | Out-Null
                
                # Finde den Start des Zyklus
                $startIndex = $circularPath.IndexOf($GroupDN)
                $cycle = $circularPath[$startIndex..($circularPath.Count - 1)]
                
                # Erstelle eine eindeutige ID für den Zyklus
                $cycleId = ($cycle | Sort-Object) -join '|'
                
                if (-not ($circularReferences | Where-Object { $_.CycleId -eq $cycleId })) {
                    $groupNames = $cycle | ForEach-Object {
                        if ($groupMemberships.ContainsKey($_)) {
                            $groupMemberships[$_].Group.Name
                        } else {
                            "Unknown"
                        }
                    }
                    
                    $circularReferences += [PSCustomObject]@{
                        CycleId = $cycleId
                        CircularPath = $groupNames -join " → "
                        NumberOfGroups = $cycle.Count
                        FirstGroup = $groupNames[0]
                        GroupsInvolved = $groupNames -join ", "
                        Risk = if ($cycle.Count -gt 3) { "High" } elseif ($cycle.Count -eq 3) { "Medium" } else { "Low" }
                        Recommendation = "Break circular reference by removing one membership link"
                    }
                }
                return
            }
            
            # Markiere als besucht und im Stack
            $groupData.Visited = $true
            $groupData.InStack = $true
            $Path.Add($GroupDN) | Out-Null
            
            # Durchlaufe alle Parent-Gruppen
            foreach ($parentDN in $groupData.MemberOf) {
                Find-CircularReference -GroupDN $parentDN -Path $Path
            }
            
            # Entferne aus Stack
            $groupData.InStack = $false
            $Path.RemoveAt($Path.Count - 1)
        }
        
        # Suche nach zirkulären Referenzen für alle Gruppen
        foreach ($groupDN in $groupMemberships.Keys) {
            if (-not $groupMemberships[$groupDN].Visited) {
                $path = New-Object System.Collections.ArrayList
                Find-CircularReference -GroupDN $groupDN -Path $path
            }
        }
        
        if ($circularReferences.Count -eq 0) {
            Write-ADReportLog -Message "No circular group references found." -Type Info -Terminal
            return @([PSCustomObject]@{
                CircularPath = "None Found"
                NumberOfGroups = 0
                FirstGroup = "N/A"
                GroupsInvolved = "No circular references detected"
                Risk = "None"
                Recommendation = "Environment is healthy - no circular references"
            })
        }
        
        Write-ADReportLog -Message "Circular group reference analysis completed. $($circularReferences.Count) circular reference(s) found." -Type Info -Terminal
        return $circularReferences | Sort-Object NumberOfGroups, FirstGroup
        
    } catch {
        Write-ADReportLog -Message "Error analyzing circular group references: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-GroupsByTypeScope {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing groups by type and scope..." -Type Info -Terminal
        
        $Groups = Get-ADGroup -Filter * -Properties GroupCategory, GroupScope, Description, ManagedBy, whenCreated, whenChanged, Members -ErrorAction Stop
        
        # Gruppiere nach Type und Scope
        $GroupStats = $Groups | Group-Object -Property GroupCategory, GroupScope
        
        $Results = foreach ($statGroup in $GroupStats) {
            $category = $statGroup.Group[0].GroupCategory
            $scope = $statGroup.Group[0].GroupScope
            
            # Berechne durchschnittliche Mitgliederzahl
            $memberCounts = @()
            foreach ($group in $statGroup.Group) {
                $memberCount = @(Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue).Count
                $memberCounts += $memberCount
            }
            
            $avgMembers = if ($memberCounts.Count -gt 0) { 
                [math]::Round(($memberCounts | Measure-Object -Average).Average, 1) 
            } else { 0 }
            
            $maxMembers = if ($memberCounts.Count -gt 0) { 
                ($memberCounts | Measure-Object -Maximum).Maximum 
            } else { 0 }
            
            # Best Practice Empfehlungen
            $recommendation = switch ("$category-$scope") {
                "Security-Global" { "Standard for most security groups" }
                "Security-DomainLocal" { "Good for resource permissions" }
                "Security-Universal" { "Use for cross-forest scenarios" }
                "Distribution-Global" { "Standard for email distribution" }
                "Distribution-DomainLocal" { "Rarely used - review necessity" }
                "Distribution-Universal" { "Good for cross-forest distribution" }
                default { "Review group configuration" }
            }
            
            [PSCustomObject]@{
                GroupCategory = $category
                GroupScope = $scope
                Count = $statGroup.Count
                Percentage = [math]::Round(($statGroup.Count / $Groups.Count) * 100, 2)
                AverageMemberCount = $avgMembers
                MaxMemberCount = $maxMembers
                Recommendation = $recommendation
                Examples = ($statGroup.Group | Select-Object -First 3 | ForEach-Object { $_.Name }) -join ", "
            }
        }
        
        Write-ADReportLog -Message "Groups by type/scope analysis completed. $($Results.Count) combinations found." -Type Info -Terminal
        return $Results | Sort-Object GroupCategory, GroupScope
        
    } catch {
        Write-ADReportLog -Message "Error analyzing groups by type and scope: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-DynamicDistGroups {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing dynamic distribution groups..." -Type Info -Terminal
        
        # Hinweis: Dynamic Distribution Groups sind ein Exchange-Feature
        # Diese Funktion würde Exchange-Cmdlets benötigen
        
        Write-ADReportLog -Message "Dynamic Distribution Groups require Exchange PowerShell module." -Type Warning
        
        return @([PSCustomObject]@{
            Name = "Not Available"
            Status = "Exchange Required"
            Description = "Dynamic Distribution Groups are an Exchange feature"
            Recommendation = "Use Exchange Management Shell for this report"
        })
        
    } catch {
        Write-ADReportLog -Message "Error analyzing dynamic distribution groups: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-MailEnabledGroups {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing mail-enabled groups..." -Type Info -Terminal
        
        # Suche nach Gruppen mit Mail-Attributen
        $MailGroups = Get-ADGroup -Filter "mail -like '*'" -Properties mail, proxyAddresses, legacyExchangeDN, msExchRecipientTypeDetails, GroupCategory, GroupScope, ManagedBy, Description -ErrorAction Stop
        
        if ($MailGroups.Count -eq 0) {
            Write-ADReportLog -Message "No mail-enabled groups found." -Type Info -Terminal
            return @([PSCustomObject]@{
                Name = "No Results"
                Mail = "N/A"
                GroupCategory = "N/A"
                Status = "No mail-enabled groups found"
                Recommendation = "Mail-enabled groups may be managed through Exchange"
            })
        }
        
        $Results = foreach ($group in $MailGroups) {
            # Prüfe Gruppenmitgliederzahl
            $memberCount = @(Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue).Count
            
            # Analysiere Mail-Eigenschaften
            $hasMultipleProxies = ($group.proxyAddresses -and $group.proxyAddresses.Count -gt 1)
            $primarySMTP = $group.proxyAddresses | Where-Object { $_ -clike "SMTP:*" } | Select-Object -First 1
            
            # Manager auflösen
            $managerName = "None"
            if ($group.ManagedBy) {
                try {
                    $manager = Get-ADUser -Identity $group.ManagedBy -Properties DisplayName -ErrorAction SilentlyContinue
                    if ($manager) {
                        $managerName = $manager.DisplayName
                    }
                } catch {
                    $managerName = "Unknown"
                }
            }
            
            # Status und Empfehlungen
            $status = "Active"
            $recommendations = @()
            
            if ([string]::IsNullOrWhiteSpace($group.mail)) {
                $status = "Incomplete"
                $recommendations += "Mail attribute is empty"
            }
            
            if ($memberCount -eq 0) {
                $recommendations += "Group has no members"
            }
            
            if ($memberCount -gt 1000) {
                $recommendations += "Large distribution list - consider breaking up"
            }
            
            if ([string]::IsNullOrWhiteSpace($group.ManagedBy)) {
                $recommendations += "No manager assigned"
            }
            
            [PSCustomObject]@{
                Name = $group.Name
                Mail = $group.mail
                PrimarySMTP = if ($primarySMTP) { $primarySMTP -replace "SMTP:", "" } else { $group.mail }
                GroupCategory = $group.GroupCategory
                GroupScope = $group.GroupScope
                MemberCount = $memberCount
                ManagedBy = $managerName
                ProxyAddressCount = if ($group.proxyAddresses) { $group.proxyAddresses.Count } else { 0 }
                Status = $status
                Recommendations = if ($recommendations) { $recommendations -join "; " } else { "Properly configured" }
                Description = $group.Description
            }
        }
        
        Write-ADReportLog -Message "Mail-enabled groups analysis completed. $($Results.Count) groups found." -Type Info -Terminal
        return $Results | Sort-Object MemberCount -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing mail-enabled groups: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-GroupsWithoutOwners {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing groups without owners/managers..." -Type Info -Terminal
        
        $Groups = Get-ADGroup -Filter * -Properties ManagedBy, Description, GroupCategory, GroupScope, whenCreated, whenChanged, info -ErrorAction Stop
        
        # Filtere Gruppen ohne ManagedBy
        $UnmanagedGroups = $Groups | Where-Object { [string]::IsNullOrWhiteSpace($_.ManagedBy) }
        
        $Results = foreach ($group in $UnmanagedGroups) {
            # Prüfe Gruppengröße
            $memberCount = @(Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue).Count
            
            # Alter der Gruppe
            $ageInDays = if ($group.whenCreated) { 
                (New-TimeSpan -Start $group.whenCreated -End (Get-Date)).Days 
            } else { 0 }
            
            # Risikobewertung
            $riskLevel = "Low"
            $riskFactors = @()
            
            if ($group.GroupCategory -eq "Security") {
                $riskLevel = "Medium"
                $riskFactors += "Security group"
            }
            
            if ($memberCount -gt 50) {
                if ($riskLevel -eq "Low") { $riskLevel = "Medium" }
                $riskFactors += "Large membership ($memberCount members)"
            }
            
            if ($memberCount -gt 100 -and $group.GroupCategory -eq "Security") {
                $riskLevel = "High"
            }
            
            if ($ageInDays -gt 365) {
                $riskFactors += "Old group (>1 year)"
            }
            
            # Empfehlungen
            $recommendations = @("Assign a group manager/owner")
            
            if ($memberCount -eq 0 -and $ageInDays -gt 90) {
                $recommendations += "Consider deleting empty old group"
            }
            
            if ($group.GroupCategory -eq "Security" -and $memberCount -gt 20) {
                $recommendations += "Security groups should have designated owners"
            }
            
            [PSCustomObject]@{
                Name = $group.Name
                GroupCategory = $group.GroupCategory
                GroupScope = $group.GroupScope
                MemberCount = $memberCount
                Description = if ($group.Description) { $group.Description } else { "(No description)" }
                WhenCreated = $group.whenCreated
                AgeInDays = $ageInDays
                LastModified = $group.whenChanged
                RiskLevel = $riskLevel
                RiskFactors = if ($riskFactors) { $riskFactors -join "; " } else { "None" }
                Recommendations = $recommendations -join "; "
            }
        }
        
        Write-ADReportLog -Message "Groups without owners analysis completed. $($Results.Count) groups found." -Type Info -Terminal
        return $Results | Sort-Object RiskLevel, MemberCount -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing groups without owners: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-LargeGroups {
    [CmdletBinding()]
    param([int]$Threshold = 100)
    
    try {
        Write-ADReportLog -Message "Analyzing large groups (threshold: $Threshold members)..." -Type Info -Terminal
        
        $Groups = Get-ADGroup -Filter * -Properties Description, GroupCategory, GroupScope, ManagedBy, whenCreated -ErrorAction Stop
        
        $LargeGroups = @()
        
        foreach ($group in $Groups) {
            $memberCount = @(Get-ADGroupMember -Identity $group.DistinguishedName -Recursive -ErrorAction SilentlyContinue).Count
            
            if ($memberCount -ge $Threshold) {
                # Nested group analysis
                $directMembers = @(Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue)
                $nestedGroups = @($directMembers | Where-Object { $_.objectClass -eq "group" })
                
                # Manager info
                $managerName = "None"
                if ($group.ManagedBy) {
                    try {
                        $manager = Get-ADUser -Identity $group.ManagedBy -Properties DisplayName -ErrorAction SilentlyContinue
                        if ($manager) {
                            $managerName = $manager.DisplayName
                        }
                    } catch {
                        $managerName = "Unknown"
                    }
                }
                
                # Performance impact assessment
                $performanceImpact = "Low"
                if ($memberCount -gt 1000) {
                    $performanceImpact = "High"
                } elseif ($memberCount -gt 500) {
                    $performanceImpact = "Medium"
                }
                
                # Recommendations
                $recommendations = @()
                if ($memberCount -gt 5000) {
                    $recommendations += "Consider breaking into smaller groups"
                }
                if ($nestedGroups.Count -gt 10) {
                    $recommendations += "High number of nested groups may impact performance"
                }
                if ($group.GroupScope -eq "Global" -and $memberCount -gt 5000) {
                    $recommendations += "Consider using Universal scope for large groups"
                }
                
                $LargeGroups += [PSCustomObject]@{
                    Name = $group.Name
                    GroupCategory = $group.GroupCategory
                    GroupScope = $group.GroupScope
                    TotalMembers = $memberCount
                    DirectMembers = $directMembers.Count
                    NestedGroups = $nestedGroups.Count
                    ManagedBy = $managerName
                    Description = if ($group.Description) { $group.Description } else { "(No description)" }
                    WhenCreated = $group.whenCreated
                    PerformanceImpact = $performanceImpact
                    Recommendations = if ($recommendations) { $recommendations -join "; " } else { "Monitor group size" }
                }
            }
        }
        
        if ($LargeGroups.Count -eq 0) {
            Write-ADReportLog -Message "No large groups found (threshold: $Threshold members)." -Type Info -Terminal
            return @([PSCustomObject]@{
                Name = "No Results"
                TotalMembers = 0
                Status = "No groups with $Threshold or more members found"
                Recommendations = "All groups are within size limits"
            })
        }
        
        Write-ADReportLog -Message "Large groups analysis completed. $($LargeGroups.Count) groups found." -Type Info -Terminal
        return $LargeGroups | Sort-Object TotalMembers -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing large groups: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-RecentlyModifiedGroups {
    [CmdletBinding()]
    param([int]$Days = 7)
    
    try {
        Write-ADReportLog -Message "Analyzing recently modified groups (last $Days days)..." -Type Info -Terminal
        
        $CutoffDate = (Get-Date).AddDays(-$Days)
        $Groups = Get-ADGroup -Filter "whenChanged -gt '$CutoffDate'" -Properties whenChanged, whenCreated, ManagedBy, Description, GroupCategory, GroupScope, modifyTimeStamp -ErrorAction Stop
        
        if ($Groups.Count -eq 0) {
            Write-ADReportLog -Message "No groups modified in the last $Days days." -Type Info -Terminal
            return @([PSCustomObject]@{
                Name = "No Results"
                Status = "No groups modified in the last $Days days"
                LastModified = "N/A"
                Recommendations = "Normal - no recent group modifications"
            })
        }
        
        $Results = foreach ($group in $Groups) {
            # Versuche die Art der Änderung zu ermitteln
            $changeType = "Modified"
            if ($group.whenCreated -and $group.whenCreated -gt $CutoffDate) {
                $changeType = "Created"
            }
            
            # Manager info
            $managerName = "None"
            if ($group.ManagedBy) {
                try {
                    $manager = Get-ADUser -Identity $group.ManagedBy -Properties DisplayName -ErrorAction SilentlyContinue
                    if ($manager) {
                        $managerName = $manager.DisplayName
                    }
                } catch {
                    $managerName = "Unknown"
                }
            }
            
            # Mitgliederzahl für Kontext
            $memberCount = @(Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue).Count
            
            # Tage seit Änderung
            $daysSinceChange = [math]::Round((New-TimeSpan -Start $group.whenChanged -End (Get-Date)).TotalDays, 1)
            
            [PSCustomObject]@{
                Name = $group.Name
                GroupCategory = $group.GroupCategory
                GroupScope = $group.GroupScope
                ChangeType = $changeType
                LastModified = $group.whenChanged
                DaysSinceChange = $daysSinceChange
                CreatedDate = $group.whenCreated
                MemberCount = $memberCount
                ManagedBy = $managerName
                Description = if ($group.Description) { $group.Description } else { "(No description)" }
                ModificationTime = $group.whenChanged.ToString("HH:mm:ss")
            }
        }
        
        Write-ADReportLog -Message "Recently modified groups analysis completed. $($Results.Count) groups found." -Type Info -Terminal
        return $Results | Sort-Object LastModified -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing recently modified groups: $($_.Exception.Message)" -Type Error
        return @()
    }
}

# --- Weitere Computer-Report Funktionen ---
Function Get-ComputersByOSVersion {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing computers by OS version..." -Type Info -Terminal
        
        $Computers = Get-ADComputer -Filter * -Properties OperatingSystem, OperatingSystemVersion, OperatingSystemServicePack, LastLogonDate, Enabled, whenCreated -ErrorAction Stop
        
        # Gruppiere nach OS und Version
        $OSGroups = $Computers | Group-Object -Property OperatingSystem, OperatingSystemVersion
        
        $Results = foreach ($osGroup in $OSGroups) {
            $os = $osGroup.Group[0].OperatingSystem
            $version = $osGroup.Group[0].OperatingSystemVersion
            $sp = $osGroup.Group[0].OperatingSystemServicePack
            
            # Statistiken für diese OS-Version
            $enabledCount = @($osGroup.Group | Where-Object { $_.Enabled }).Count
            $activeCount = @($osGroup.Group | Where-Object { $_.LastLogonDate -and $_.LastLogonDate -gt (Get-Date).AddDays(-30) }).Count
            
            # Support-Status ermitteln
            $supportStatus = "Unknown"
            $eolDate = "Unknown"
            
            if ($os -like "*Windows 7*") {
                $supportStatus = "End of Life"
                $eolDate = "2020-01-14"
            } elseif ($os -like "*Windows 8*" -and $os -notlike "*8.1*") {
                $supportStatus = "End of Life"
                $eolDate = "2016-01-12"
            } elseif ($os -like "*Windows 8.1*") {
                $supportStatus = "End of Life"
                $eolDate = "2023-01-10"
            } elseif ($os -like "*Windows 10*") {
                # Verschiedene Windows 10 Versionen haben unterschiedliche EOL-Daten
                $supportStatus = "Check Version"
                $eolDate = "Version-dependent"
            } elseif ($os -like "*Windows 11*") {
                $supportStatus = "Supported"
                $eolDate = "Active"
            } elseif ($os -like "*Server 2008*" -and $os -notlike "*R2*") {
                $supportStatus = "End of Life"
                $eolDate = "2015-07-14"
            } elseif ($os -like "*Server 2008 R2*") {
                $supportStatus = "End of Life"
                $eolDate = "2020-01-14"
            } elseif ($os -like "*Server 2012*" -and $os -notlike "*R2*") {
                $supportStatus = "End of Life"
                $eolDate = "2023-10-10"
            } elseif ($os -like "*Server 2012 R2*") {
                $supportStatus = "End of Life"
                $eolDate = "2023-10-10"
            } elseif ($os -like "*Server 2016*") {
                $supportStatus = "Supported"
                $eolDate = "2027-01-12"
            } elseif ($os -like "*Server 2019*") {
                $supportStatus = "Supported"
                $eolDate = "2029-01-09"
            } elseif ($os -like "*Server 2022*") {
                $supportStatus = "Supported"
                $eolDate = "2031-10-14"
            }
            
            # Risikobewertung
            $riskLevel = "Low"
            if ($supportStatus -eq "End of Life") {
                $riskLevel = "Critical"
            } elseif ($supportStatus -eq "Check Version") {
                $riskLevel = "Medium"
            }
            
            [PSCustomObject]@{
                OperatingSystem = if ($os) { $os } else { "Unknown" }
                Version = if ($version) { $version } else { "Unknown" }
                ServicePack = if ($sp) { $sp } else { "None" }
                Count = $osGroup.Count
                EnabledCount = $enabledCount
                ActiveCount = $activeCount
                InactiveCount = $osGroup.Count - $activeCount
                Percentage = [math]::Round(($osGroup.Count / $Computers.Count) * 100, 2)
                SupportStatus = $supportStatus
                EndOfLifeDate = $eolDate
                RiskLevel = $riskLevel
                Recommendation = if ($supportStatus -eq "End of Life") { 
                    "Urgent: Upgrade or replace systems" 
                } elseif ($supportStatus -eq "Check Version") { 
                    "Verify specific version support status" 
                } else { 
                    "Keep systems updated" 
                }
            }
        }
        
        Write-ADReportLog -Message "Computers by OS version analysis completed. $($Results.Count) unique OS versions found." -Type Info -Terminal
        return $Results | Sort-Object RiskLevel, Count -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing computers by OS version: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-ComputersNeverLoggedOn {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing computers that never logged on..." -Type Info -Terminal
        
        $Computers = Get-ADComputer -Filter "LastLogonDate -notlike '*'" -Properties LastLogonDate, whenCreated, OperatingSystem, Enabled, Description, DistinguishedName -ErrorAction Stop
        
        if ($Computers.Count -eq 0) {
            Write-ADReportLog -Message "No computers found that never logged on." -Type Info -Terminal
            return @([PSCustomObject]@{
                Name = "No Results"
                Status = "All computers have logged on at least once"
                Recommendation = "Environment is healthy"
            })
        }
        
        $Results = foreach ($computer in $Computers) {
            # Alter des Computer-Objekts
            $ageInDays = if ($computer.whenCreated) {
                (New-TimeSpan -Start $computer.whenCreated -End (Get-Date)).Days
            } else { "Unknown" }
            
            # OU-Pfad extrahieren
            $ouPath = if ($computer.DistinguishedName -match 'CN=[^,]+,(.+)$') { $matches[1] } else { "Unknown" }
            
            # Risikobewertung
            $riskLevel = "Low"
            $recommendations = @()
            
            if ($computer.Enabled) {
                $riskLevel = "Medium"
                $recommendations += "Consider disabling unused computer"
            }
            
            if ($ageInDays -is [int] -and $ageInDays -gt 90) {
                if ($riskLevel -eq "Low") { $riskLevel = "Medium" }
                if ($computer.Enabled) { $riskLevel = "High" }
                $recommendations += "Old unused computer account - consider deletion"
            }
            
            if ($ageInDays -is [int] -and $ageInDays -gt 180) {
                $riskLevel = "High"
                $recommendations = @("Delete unused computer account")
            }
            
            [PSCustomObject]@{
                Name = $computer.Name
                OperatingSystem = if ($computer.OperatingSystem) { $computer.OperatingSystem } else { "Unknown" }
                Enabled = $computer.Enabled
                WhenCreated = $computer.whenCreated
                AgeInDays = $ageInDays
                Description = if ($computer.Description) { $computer.Description } else { "(No description)" }
                OrganizationalUnit = $ouPath
                RiskLevel = $riskLevel
                Recommendations = if ($recommendations) { $recommendations -join "; " } else { "Monitor account" }
                Status = if ($computer.Enabled) { "Enabled but never used" } else { "Disabled and never used" }
            }
        }
        
        Write-ADReportLog -Message "Computers never logged on analysis completed. $($Results.Count) computers found." -Type Info -Terminal
        return $Results | Sort-Object RiskLevel, AgeInDays -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing computers never logged on: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-DuplicateComputerNames {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing duplicate computer names..." -Type Info -Terminal
        
        $Computers = Get-ADComputer -Filter * -Properties DNSHostName, Enabled, OperatingSystem, LastLogonDate, whenCreated, DistinguishedName, IPv4Address -ErrorAction Stop
        
        # Suche nach Duplikaten basierend auf Name (ohne $)
        $ComputerNames = $Computers | ForEach-Object { 
            [PSCustomObject]@{
                Computer = $_
                CleanName = $_.Name.TrimEnd('$').ToUpper()
            }
        }
        
        # Gruppiere nach bereinigtem Namen
        $DuplicateGroups = $ComputerNames | Group-Object CleanName | Where-Object { $_.Count -gt 1 }
        
        if ($DuplicateGroups.Count -eq 0) {
            Write-ADReportLog -Message "No duplicate computer names found." -Type Info -Terminal
            return @([PSCustomObject]@{
                Name = "No Duplicates"
                Status = "No duplicate computer names detected"
                Count = 0
                Recommendation = "Environment is healthy"
            })
        }
        
        $Results = foreach ($dupGroup in $DuplicateGroups) {
            foreach ($item in $dupGroup.Group) {
                $computer = $item.Computer
                
                # OU-Pfad extrahieren
                $ouPath = if ($computer.DistinguishedName -match 'CN=[^,]+,(.+)$') { $matches[1] } else { "Unknown" }
                
                # Bestimme welcher Computer der "aktive" ist
                $isActive = $computer.Enabled -and $computer.LastLogonDate -and 
                           $computer.LastLogonDate -gt (Get-Date).AddDays(-30)
                
                [PSCustomObject]@{
                    Name = $computer.Name
                    CleanName = $item.CleanName
                    DNSHostName = $computer.DNSHostName
                    IPv4Address = if ($computer.IPv4Address) { $computer.IPv4Address } else { "No IP" }
                    Enabled = $computer.Enabled
                    OperatingSystem = if ($computer.OperatingSystem) { $computer.OperatingSystem } else { "Unknown" }
                    LastLogonDate = $computer.LastLogonDate
                    WhenCreated = $computer.whenCreated
                    OrganizationalUnit = $ouPath
                    DuplicateCount = $dupGroup.Count
                    Status = if ($isActive) { "Active Duplicate" } 
                            elseif ($computer.Enabled) { "Enabled but Inactive" } 
                            else { "Disabled" }
                    Recommendation = if (-not $computer.Enabled -and (-not $computer.LastLogonDate -or 
                                      $computer.LastLogonDate -lt (Get-Date).AddDays(-90))) {
                                        "Delete disabled/inactive duplicate"
                                    } elseif ($dupGroup.Count -gt 2) {
                                        "Multiple duplicates - immediate investigation required"
                                    } else {
                                        "Investigate and resolve duplicate"
                                    }
                }
            }
        }
        
        Write-ADReportLog -Message "Duplicate computer names analysis completed. $($DuplicateGroups.Count) duplicate groups found." -Type Info -Terminal
        return $Results | Sort-Object CleanName, Status
        
    } catch {
        Write-ADReportLog -Message "Error analyzing duplicate computer names: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-ComputersByLocation {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing computers by location..." -Type Info -Terminal
        
        $Computers = Get-ADComputer -Filter * -Properties Location, Enabled, OperatingSystem, LastLogonDate, Description, DistinguishedName -ErrorAction Stop
        
        # Gruppiere nach Location
        $LocationGroups = $Computers | Group-Object Location
        
        $Results = foreach ($locGroup in $LocationGroups) {
            $location = if ([string]::IsNullOrWhiteSpace($locGroup.Name)) { "(No Location)" } else { $locGroup.Name }
            
            # Statistiken für diese Location
            $enabledCount = @($locGroup.Group | Where-Object { $_.Enabled }).Count
            $activeCount = @($locGroup.Group | Where-Object { 
                $_.LastLogonDate -and $_.LastLogonDate -gt (Get-Date).AddDays(-30) 
            }).Count
            
            # OS-Verteilung in dieser Location
            $osDistribution = $locGroup.Group | Group-Object OperatingSystem | 
                             Sort-Object Count -Descending | 
                             Select-Object -First 3 | 
                             ForEach-Object { "$($_.Name): $($_.Count)" }
            
            # Durchschnittliche Inaktivität
            $inactiveDays = @()
            foreach ($comp in $locGroup.Group) {
                if ($comp.LastLogonDate) {
                    $inactiveDays += (New-TimeSpan -Start $comp.LastLogonDate -End (Get-Date)).Days
                }
            }
            
            $avgInactiveDays = if ($inactiveDays.Count -gt 0) {
                [math]::Round(($inactiveDays | Measure-Object -Average).Average, 1)
            } else { "N/A" }
            
            [PSCustomObject]@{
                Location = $location
                TotalComputers = $locGroup.Count
                EnabledComputers = $enabledCount
                ActiveComputers = $activeCount
                InactiveComputers = $locGroup.Count - $activeCount
                Percentage = [math]::Round(($locGroup.Count / $Computers.Count) * 100, 2)
                TopOperatingSystems = if ($osDistribution) { $osDistribution -join "; " } else { "Unknown" }
                AverageInactiveDays = $avgInactiveDays
                HealthStatus = if ($activeCount -lt ($enabledCount * 0.5)) { "Poor" } 
                              elseif ($activeCount -lt ($enabledCount * 0.8)) { "Fair" } 
                              else { "Good" }
                Recommendation = if ($enabledCount -gt 0 -and $activeCount -lt ($enabledCount * 0.5)) {
                                   "Many inactive computers - review and cleanup"
                               } elseif ([string]::IsNullOrWhiteSpace($locGroup.Name)) {
                                   "Update location information for better tracking"
                               } else {
                                   "Location properly maintained"
                               }
            }
        }
        
        Write-ADReportLog -Message "Computers by location analysis completed. $($Results.Count) locations found." -Type Info -Terminal
        return $Results | Sort-Object TotalComputers -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing computers by location: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-VirtualVsPhysical {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing virtual vs physical computers..." -Type Info -Terminal
        
        $Computers = Get-ADComputer -Filter * -Properties OperatingSystem, Manufacturer, Model, Enabled, LastLogonDate, Description -ErrorAction Stop
        
        # Kategorisiere Computer als virtuell oder physisch
        $ComputerTypes = foreach ($computer in $Computers) {
            $type = "Physical"  # Standard-Annahme
            $platform = "Unknown"
            
            # Prüfe verschiedene Indikatoren für virtuelle Maschinen
            # Hinweis: Diese Properties sind möglicherweise nicht immer verfügbar
            if ($computer.Manufacturer) {
                if ($computer.Manufacturer -like "*VMware*") {
                    $type = "Virtual"
                    $platform = "VMware"
                } elseif ($computer.Manufacturer -like "*Microsoft Corporation*" -and $computer.Model -like "*Virtual*") {
                    $type = "Virtual"
                    $platform = "Hyper-V"
                } elseif ($computer.Manufacturer -like "*Xen*") {
                    $type = "Virtual"
                    $platform = "Xen"
                } elseif ($computer.Manufacturer -like "*QEMU*" -or $computer.Manufacturer -like "*KVM*") {
                    $type = "Virtual"
                    $platform = "KVM/QEMU"
                } elseif ($computer.Manufacturer -like "*Oracle*" -and $computer.Model -like "*VirtualBox*") {
                    $type = "Virtual"
                    $platform = "VirtualBox"
                } elseif ($computer.Manufacturer -like "*Amazon*" -or $computer.Manufacturer -like "*EC2*") {
                    $type = "Virtual"
                    $platform = "AWS EC2"
                } elseif ($computer.Manufacturer -like "*Google*" -or $computer.Model -like "*Google*") {
                    $type = "Virtual"
                    $platform = "Google Cloud"
                } elseif ($computer.Manufacturer -like "*Microsoft*" -and $computer.Model -like "*Azure*") {
                    $type = "Virtual"
                    $platform = "Azure"
                }
            }
            
            # Zusätzliche Prüfung über Modell
            if ($computer.Model) {
                if ($computer.Model -like "*Virtual*" -or $computer.Model -like "*VM*") {
                    if ($type -eq "Physical") {
                        $type = "Virtual"
                        $platform = "Generic VM"
                    }
                }
            }
            
            [PSCustomObject]@{
                Computer = $computer
                Type = $type
                Platform = $platform
            }
        }
        
        # Gruppiere nach Typ
        $TypeGroups = $ComputerTypes | Group-Object Type
        
        # Erstelle Zusammenfassung
        $Summary = foreach ($typeGroup in $TypeGroups) {
            $computers = $typeGroup.Group.Computer
            
            # Statistiken
            $enabledCount = @($computers | Where-Object { $_.Enabled }).Count
            $activeCount = @($computers | Where-Object { 
                $_.LastLogonDate -and $_.LastLogonDate -gt (Get-Date).AddDays(-30) 
            }).Count
            
            # Platform-Verteilung für virtuelle Maschinen
            $platformDist = if ($typeGroup.Name -eq "Virtual") {
                $typeGroup.Group | Group-Object Platform | 
                Sort-Object Count -Descending | 
                ForEach-Object { "$($_.Name): $($_.Count)" }
            } else { @("N/A") }
            
            # OS-Verteilung
            $osDist = $computers | Group-Object OperatingSystem | 
                     Sort-Object Count -Descending | 
                     Select-Object -First 5 | 
                     ForEach-Object { "$($_.Name): $($_.Count)" }
            
            [PSCustomObject]@{
                Type = $typeGroup.Name
                Count = $typeGroup.Count
                Percentage = [math]::Round(($typeGroup.Count / $Computers.Count) * 100, 2)
                Enabled = $enabledCount
                Active = $activeCount
                Inactive = $typeGroup.Count - $activeCount
                Platforms = if ($platformDist) { $platformDist -join "; " } else { "N/A" }
                TopOperatingSystems = if ($osDist) { $osDist -join "; " } else { "Unknown" }
                Recommendations = if ($typeGroup.Name -eq "Virtual" -and $platformDist.Count -gt 3) {
                                    "Multiple virtualization platforms detected - consider standardization"
                                } elseif ($activeCount -lt ($enabledCount * 0.7)) {
                                    "Many inactive $($typeGroup.Name.ToLower()) machines - review for cleanup"
                                } else {
                                    "Environment properly maintained"
                                }
            }
        }
        
        Write-ADReportLog -Message "Virtual vs Physical analysis completed. Found $($TypeGroups.Count) types." -Type Info -Terminal
        return $Summary  # Nur Summary zurückgeben für bessere Übersicht
        
    } catch {
        Write-ADReportLog -Message "Error analyzing virtual vs physical computers: $($_.Exception.Message)" -Type Error
        return @()
    }
}




# --- Skriptstart ---
Start-ADReportGUI




# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDKpcp1XDnv31BY
# Oi44eTCqpaZP7ZuQCB5kOaHbURrVxaCCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# DQEJBDEiBCCZeQG+l9zI40rsiae6Ps7NWv8F5Y607R4uwHGdjZTcvjANBgkqhkiG
# 9w0BAQEFAASCAQC4/MXaOcxORZwUKf0mWOsFGYNnX6yqVhNZXrbbSg2ym6myzRHr
# hfyID8GTA9dGf7gXN1wJRx8fZ879OOT0p3GM+snxw7C0n5KlbrlOgu/2PpV3dYBp
# p8bONLfaTJT3XQ4R1r4mekZsEQmP1+H3K5a7i4hZyLJMpPITgvvFyVkyyxMLMEhm
# cyA+GM/kxWRvEeREd77ILSs5EiEVZ4GmHnOvgqUt3vrEadrsa8P4DamnZI8m0OmR
# 1EAB5xpOOCCgqQVqTbwnoiyegcOPv9Ket+fi6/R/Sc5ZiH1SZMe/VFmCTY15IfGI
# qWrgH9ivpRuA5hX7S0kbnd/7HPynKj9L8ctooYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTEzOVowLwYJKoZIhvcNAQkEMSIEIPN+tjSGh8omi04OcI44HiSZRbkD
# j5z/Wjd10ny537I9MA0GCSqGSIb3DQEBAQUABIICAJuh43LKpoyMT+/IrtIfB49I
# YbymBsS36VRKPQFdKW5ok1XgUSlDwoFOP6yaeqfZ5hlLuiqQ9ZEUR4q3r9RPBgCb
# DaREaqYQspDH52GigvMKVeUfo9lBbzoqBNuUoFUms897rd0mTCbnQDVnz8KsP0OZ
# epB5OJz0Px8TUW20QSx/gKfufHre6eRDnotIhhwYXnXoCsMYhSz+91XRgha3ZoT1
# Q3/WQ0kbjSCIDBubVlyypfCBpBUtjiyC6Kz52yEkWv5pe66LEP589vdU6tI7nvlt
# TiE0qaS59qDS9tBTM2d3X6GjzXSbUx+HP8rWtVfEt/diPXuFSTXumfX/fQhVkMaN
# bBtkl0SekHkw30xxcj0tqstaw0IdxqOtwhGRMQXfIPkvIPZcMZhVJOrM6EPGYlk4
# LCVCmKlGtdNQAXPZda7Ta4KtwwFdnldyiuHjyf22+cIQD7Dkm5C67uc+ZqIChUv2
# G7aaBM65eY7yeJ6hs3wKzfJnvUPU/irfSd+lXo5bqviXr2YlF81SP7NGkurNmnvB
# cvjGHDU8S0N8LG9RaK86zAyLM6YPKG8/q5cGgzU0MpgAwZ2tVH1MyHlUzaounIBp
# s+RlveUp1l7fF6n1ynLZhOhk2VR8+EO5u8jhf0UFJkVsTABe0M7H/4GqlUVo2Itq
# tSTykEcIDVkA0WEOWYG7
# SIG # End signature block

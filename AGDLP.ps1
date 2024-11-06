$runAsAdmin = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
$adminRole = [Security.Principal.WindowsBuiltInRole]::Administrator

if (-not $runAsAdmin.IsInRole($adminRole)) {
    # Relaunch the script as Administrator
    $arguments = "$($myinvocation.MyCommand.Definition)"
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File $arguments" -Verb RunAs
    exit
}

# Import-Module und GUI-Komponenten laden
Add-Type -AssemblyName System.Windows.Forms
Import-Module ActiveDirectory

# Funktion zum Abrufen der aktuellen Domäne
function Get-DomainName {
    $domain = (Get-ADDomain).DNSRoot
    return $domain
}

# Funktion zum Abrufen von Gruppen aus AD
function Get-Groups {
    $groups = Get-ADGroup -Filter *
    return $groups | ForEach-Object { $_.Name }
}

# Funktion zum Abrufen von Benutzern aus AD
function Get-Users {
    $users = Get-ADUser -Filter *
    return $users | ForEach-Object { $_.SamAccountName }
}

$domain = Get-DomainName
$ggList = Get-ADGroup -Filter { GroupScope -eq 'Global' }
$ggList = $ggList | Select-Object -ExpandProperty Name
$dlList = Get-Groups
$userList = Get-Users

# GUI zur Zuordnung von GG zu DL und Benutzern erstellen
$form = New-Object System.Windows.Forms.Form
$form.Text = "GG zu DL Zuordnung und Benutzer"
$form.Size = New-Object System.Drawing.Size(805, 435)  # Fenstergröße anpassen

# Benutzer-Eingabe und Liste (ganz oben)
$userLabel = New-Object System.Windows.Forms.Label
$userLabel.Text = "Benutzer zu Globalgruppe zuordnen:"
$userLabel.Location = New-Object System.Drawing.Point(10, 10)
$userLabel.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($userLabel)

$userTextBox = New-Object System.Windows.Forms.TextBox
$userTextBox.Location = New-Object System.Drawing.Point(10, 30)
$userTextBox.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($userTextBox)

$userListBox = New-Object System.Windows.Forms.ListBox
$userListBox.Location = New-Object System.Drawing.Point(10, 55)
$userListBox.Size = New-Object System.Drawing.Size(250, 200)  # Anpassung der Größe
$userListBox.Items.AddRange($userList)
$userListBox.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
$form.Controls.Add($userListBox)

$userTextBox.Add_TextChanged({
    $userListBox.Items.Clear()
    $filteredUsers = $userList | Where-Object { $_ -like "*$($userTextBox.Text)*" }
    $userListBox.Items.AddRange($filteredUsers)
})

# GG-Eingabe und Liste (mitte)
$ggLabel = New-Object System.Windows.Forms.Label
$ggLabel.Text = "Globalgruppe (GG):"
$ggLabel.Location = New-Object System.Drawing.Point(270, 10)
$ggLabel.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($ggLabel)

$ggTextBox = New-Object System.Windows.Forms.TextBox
$ggTextBox.Location = New-Object System.Drawing.Point(270, 30)
$ggTextBox.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($ggTextBox)

$ggListBox = New-Object System.Windows.Forms.ListBox
$ggListBox.Location = New-Object System.Drawing.Point(270, 55)
$ggListBox.Size = New-Object System.Drawing.Size(250, 200)
$ggListBox.Items.AddRange($ggList)
$ggListBox.SelectionMode = [System.Windows.Forms.SelectionMode]::One
$form.Controls.Add($ggListBox)

$ggTextBox.Add_TextChanged({
    $ggListBox.Items.Clear()
    $filteredGGs = $ggList | Where-Object { $_ -like "*$($ggTextBox.Text)*" }
    $ggListBox.Items.AddRange($filteredGGs)
})

# DL-Eingabe und Liste (ganz unten)
$dlLabel = New-Object System.Windows.Forms.Label
$dlLabel.Text = "Domänenlokale Gruppen (DL):"
$dlLabel.Location = New-Object System.Drawing.Point(530, 10)
$dlLabel.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($dlLabel)

$dlTextBox = New-Object System.Windows.Forms.TextBox
$dlTextBox.Location = New-Object System.Drawing.Point(530, 30)
$dlTextBox.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($dlTextBox)

$dlListBox = New-Object System.Windows.Forms.ListBox
$dlListBox.Location = New-Object System.Drawing.Point(530, 55)
$dlListBox.Size = New-Object System.Drawing.Size(250, 200)
$dlListBox.Items.AddRange($dlList)
$dlListBox.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
$form.Controls.Add($dlListBox)

$dlTextBox.Add_TextChanged({
    $dlListBox.Items.Clear()
    $filteredDLs = $dlList | Where-Object { $_ -like "*$($dlTextBox.Text)*" }
    $dlListBox.Items.AddRange($filteredDLs)
})

# RichTextBox für Debug-Informationen hinzufügen
$debugTextBox = New-Object System.Windows.Forms.RichTextBox
$debugTextBox.Location = New-Object System.Drawing.Point(10, 265)
$debugTextBox.Size = New-Object System.Drawing.Size(770, 80)
$debugTextBox.ReadOnly = $true
$form.Controls.Add($debugTextBox)

# Methode zur Ausgabe von Nachrichten im RichTextBox
function Add-DebugMessage {
    param (
        [string]$message,
        [string]$color
    )

    $debugTextBox.SelectionStart = $debugTextBox.TextLength
    $debugTextBox.SelectionLength = 0
    $debugTextBox.SelectionColor = $color
    $debugTextBox.AppendText("$message`r`n")
}

# OK-Button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(10, 355)
$okButton.Size = New-Object System.Drawing.Size(380, 30)
$okButton.Add_Click({
    $gg = $ggListBox.SelectedItem
    $dl = $dlListBox.SelectedItems
    $users = $userListBox.SelectedItems

    # Zuerst Benutzer zur Globalgruppe hinzufügen
    foreach ($user in $users) {
        if (-not (Get-ADGroupMember -Identity $gg -Recursive | Where-Object { $_.SamAccountName -eq $user })) {
            Add-ADGroupMember -Identity $gg -Members $user
            Add-DebugMessage "Benutzer $user wurde der Globalgruppe $gg hinzugefügt." "Green"
        } else {
            Add-DebugMessage "Benutzer $user ist bereits Mitglied der Globalgruppe $gg." "Orange"
        }
    }

    # Dann Globalgruppe zu einer anderen Globalgruppe hinzufügen
    foreach ($dlGroup in $dl) {
        if (Get-ADGroup -Filter "Name -eq '$dlGroup'") {
            $groupType = (Get-ADGroup -Identity $dlGroup).GroupScope
            if ($groupType -eq 'Global') {
                if (-not (Get-ADGroupMember -Identity $dlGroup -Recursive | Where-Object { $_.SamAccountName -eq $gg })) {
                    Add-ADGroupMember -Identity $dlGroup -Members $gg
                    Add-DebugMessage "Globalgruppe $gg wurde zur Globalgruppe $dlGroup hinzugefügt." "Green"
                } else {
                    Add-DebugMessage "Globalgruppe $gg ist bereits Mitglied der Globalgruppe $dlGroup." "Orange"
                }
            }
        }
    }

    # Schließlich Globalgruppe zu einer Domänenlokalen Gruppe hinzufügen
    foreach ($dlGroup in $dl) {
        if (Get-ADGroup -Filter "Name -eq '$dlGroup'") {
            $groupType = (Get-ADGroup -Identity $dlGroup).GroupScope
            if ($groupType -eq 'DomainLocal') {
                if (-not (Get-ADGroupMember -Identity $dlGroup -Recursive | Where-Object { $_.SamAccountName -eq $gg })) {
                    Add-ADGroupMember -Identity $dlGroup -Members $gg
                    Add-DebugMessage "Globalgruppe $gg wurde zur Domänenlokalen Gruppe $dlGroup hinzugefügt." "Green"
                } else {
                    Add-DebugMessage "Globalgruppe $gg ist bereits Mitglied der Domänenlokalen Gruppe $dlGroup." "Orange"
                }
            }
        }
    }

    # Trennstrich
    $separator = '+' * 127  #mal das Zeichen '+' wiederholen
    Add-DebugMessage $separator "Black"

    [System.Windows.Forms.MessageBox]::Show("Vorgang abgeschlossen.")
})
$form.Controls.Add($okButton)

# Beenden-Button
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Text = "Beenden"
$exitButton.Location = New-Object System.Drawing.Point(400, 355)
$exitButton.Size = New-Object System.Drawing.Size(380, 30)
$exitButton.Add_Click({
    $form.Close()
})
$form.Controls.Add($exitButton)

$form.ShowDialog()

# Version 1.0:
# Created 2024 by Tim Eertmoed, Germany to work on Windows Server 2019 as an user creating script.

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
$domain = Get-DomainName
$ggList = Get-ADGroup -Filter { GroupScope -eq 'Global' }
$ggList = $ggList | Select-Object -ExpandProperty Name
$dlList = Get-Groups

# GUI zur Zuordnung von GG zu DL erstellen
$form = New-Object System.Windows.Forms.Form
$form.Text = "GG zu DL Zuordnung"
$form.Size = New-Object System.Drawing.Size(550, 400)  # Großes komfortables Fenster

# GG-Eingabe und Liste
$ggLabel = New-Object System.Windows.Forms.Label
$ggLabel.Text = "Globalgruppe (GG):"
$ggLabel.Location = New-Object System.Drawing.Point(10, 10)
$ggLabel.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($ggLabel)

$ggTextBox = New-Object System.Windows.Forms.TextBox
$ggTextBox.Location = New-Object System.Drawing.Point(10, 30)
$ggTextBox.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($ggTextBox)

$ggListBox = New-Object System.Windows.Forms.ListBox
$ggListBox.Location = New-Object System.Drawing.Point(10, 55)
$ggListBox.Size = New-Object System.Drawing.Size(250, 200)
$ggListBox.Items.AddRange($ggList)
$ggListBox.SelectionMode = [System.Windows.Forms.SelectionMode]::One
$form.Controls.Add($ggListBox)

$ggTextBox.Add_TextChanged({
    $ggListBox.Items.Clear()
    $filteredGGs = $ggList | Where-Object { $_ -like "*$($ggTextBox.Text)*" }
    $ggListBox.Items.AddRange($filteredGGs)
})

# DL-Eingabe und Liste
$dlLabel = New-Object System.Windows.Forms.Label
$dlLabel.Text = "Domänenlokale Gruppen (DL):"
$dlLabel.Location = New-Object System.Drawing.Point(270, 10)
$dlLabel.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($dlLabel)

$dlTextBox = New-Object System.Windows.Forms.TextBox
$dlTextBox.Location = New-Object System.Drawing.Point(270, 30)
$dlTextBox.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($dlTextBox)

$dlListBox = New-Object System.Windows.Forms.ListBox
$dlListBox.Location = New-Object System.Drawing.Point(270, 55)
$dlListBox.Size = New-Object System.Drawing.Size(250, 200)
$dlListBox.Items.AddRange($dlList)
$dlListBox.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
$form.Controls.Add($dlListBox)

$dlTextBox.Add_TextChanged({
    $dlListBox.Items.Clear()
    $filteredDLs = $dlList | Where-Object { $_ -like "*$($dlTextBox.Text)*" }
    $dlListBox.Items.AddRange($filteredDLs)
})


# OK-Button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(80, 300)
$okButton.Size = New-Object System.Drawing.Size(100, 30)
$okButton.Add_Click({
    $gg = $ggListBox.SelectedItem
    $dl = $dlListBox.SelectedItems

    foreach ($dlGroup in $dl) {
        if (-not (Get-ADGroup -Filter "Name -eq '$dlGroup'")) {
            New-ADGroup -Name $dlGroup -GroupScope DomainLocal -GroupCategory Security -Path "OU=DL_Groups,DC=deineDomäne,DC=local"
            Write-Output "Gruppe $dlGroup wurde erstellt."
        } else {
            Write-Output "Gruppe $dlGroup existiert bereits."
        }

        if (-not (Get-ADGroupMember -Identity $dlGroup -Recursive | Where-Object { $_.SamAccountName -eq $gg })) {
            Add-ADGroupMember -Identity $dlGroup -Members $gg
            Write-Output "Globalgruppe $gg wurde zur Domänenlokalen Gruppe $dlGroup hinzugefügt."
        } else {
            Write-Output "Globalgruppe $gg ist bereits Mitglied der Domänenlokalen Gruppe $dlGroup."
        }
    }

    [System.Windows.Forms.MessageBox]::Show("Globalgruppe $gg wurde den Domänenlokalen Gruppe/n $($dl -join ', ') hinzugefügt.")
})
$form.Controls.Add($okButton)

# Beenden-Button
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Text = "Beenden"
$exitButton.Location = New-Object System.Drawing.Point(320, 300)
$exitButton.Size = New-Object System.Drawing.Size(100, 30)
$exitButton.Add_Click({
    $form.Close()
})
$form.Controls.Add($exitButton)

$form.ShowDialog()

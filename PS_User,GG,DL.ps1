###Benötigt eine befüllte UGD_Blaupause.csv###

# Import-Module und GUI-Komponenten laden
Add-Type -AssemblyName System.Windows.Forms
Import-Module ActiveDirectory

# Funktion, um die Domain aus dem OU-Pfad zu extrahieren
function Get-DomainFromOUPath {
    param ([string]$ouPath)
    $ouParts = $ouPath -split ","
    $dcParts = $ouParts | Where-Object { $_ -like "DC=*" }
    $domain = ($dcParts -join ".").Replace("DC=", "")
    return $domain
}

# Funktion zur OU-Auswahl
function Select-OU {
    $ous = Get-ADOrganizationalUnit -Filter *
    $ouNames = $ous | ForEach-Object { $_.DistinguishedName }
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select OU"
    $form.Size = New-Object System.Drawing.Size(400, 300)
    
    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Dock = "Fill"
    $listBox.Items.AddRange($ouNames)
    $form.Controls.Add($listBox)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Dock = "Bottom"
    $okButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })
    $form.Controls.Add($okButton)
    
    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $listBox.SelectedItem
    } else {
        return $null
    }
}

# CSV-Datei Pfad auswählen über ein Auswahlfenster
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
$openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
$openFileDialog.FilterIndex = 1
$openFileDialog.Multiselect = $false
if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $csvFilePath = $openFileDialog.FileName
} else {
    Write-Output "Es wurde keine Datei ausgewählt. Skript wird beendet."
    exit
}

# Standardpasswort per Texteingabe in einem Fenster
$form = New-Object System.Windows.Forms.Form
$form.Text = "Passworteingabe"
$form.Size = New-Object System.Drawing.Size(300, 150)

$label = New-Object System.Windows.Forms.Label
$label.Text = "Bitte geben Sie das Standardpasswort ein:"
$label.Location = New-Object System.Drawing.Point(10, 20)
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10, 50)
$textBox.Size = New-Object System.Drawing.Size(260, 20)
$textBox.UseSystemPasswordChar = $true
$form.Controls.Add($textBox)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(100, 80)
$okButton.Add_Click({
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})
$form.Controls.Add($okButton)

if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $standardPasswordText = $textBox.Text
    $standardPassword = (ConvertTo-SecureString $standardPasswordText -AsPlainText -Force)
} else {
    Write-Output "Es wurde kein Passwort eingegeben. Skript wird beendet."
    exit
}

# CSV-Datei einlesen mit Semikolon als Delimiter
$agdplTable = Import-Csv -Path $csvFilePath -Delimiter ';'

# OU auswählen
$ouPath = Select-OU
if ($ouPath -eq $null) {
    Write-Output "Es wurde keine OU ausgewählt. Skript wird beendet."
    exit
}

# Domain für die E-Mail-Adressen aus dem OU-Pfad generieren
$emailDomain = Get-DomainFromOUPath -ouPath $ouPath

# IGLDA ausführen und Gruppen bei Bedarf erstellen
foreach ($entry in $agdplTable) {
    $vorname = $entry.Vorname
    $nachname = $entry.Nachname
    $globalGroup = $entry.GlobalGroup
    $domainLocalGroups = $entry.DomainLocalGroups -split "," # DL-Gruppen aufteilen

    if ([string]::IsNullOrEmpty($vorname) -or [string]::IsNullOrEmpty($nachname) -or [string]::IsNullOrEmpty($globalGroup)) {
        continue
    }
    
    $account = $vorname.Substring(0,1) + $nachname
    $email = $vorname.Substring(0,1) + $nachname + "@" + $emailDomain

    # Benutzerkonto erstellen, wenn nicht vorhanden
    $user = Get-ADUser -Filter {SamAccountName -eq $account}
    if (-not $user) {
        New-ADUser -Name "$vorname $nachname" -GivenName $vorname -Surname $nachname -SamAccountName $account -UserPrincipalName $email -EmailAddress $email -AccountPassword $standardPassword -Enabled $true -Path $ouPath
        Write-Output "Benutzer $vorname $nachname mit E-Mail $email wurde erstellt."
    } else {
        Write-Output "Benutzer $vorname $nachname existiert bereits."
    }

    # Globalgruppe erstellen, wenn nicht vorhanden
    $globalGroupObj = Get-ADGroup -Filter {Name -eq $globalGroup}
    if (-not $globalGroupObj) {
        New-ADGroup -Name $globalGroup -GroupScope Global -Path $ouPath
        Write-Output "Globalgruppe $globalGroup wurde erstellt."
    } else {
        Write-Output "Globalgruppe $globalGroup existiert bereits."
    }

    # Konto zu Globalgruppe hinzufügen
    if (-not (Get-ADGroupMember -Identity $globalGroup -Recursive | Where-Object { $_.SamAccountName -eq $account })) {
        Add-ADGroupMember -Identity $globalGroup -Members $account
        Write-Output "Benutzer $account wurde zur Globalgruppe $globalGroup hinzugefügt."
    } else {
        Write-Output "Benutzer $account ist bereits Mitglied der Globalgruppe $globalGroup."
    }

   # Domänenlokale Gruppen erstellen und verknüpfen
foreach ($domainLocalGroup in $domainLocalGroups) {
    $domainLocalGroup = $domainLocalGroup.Trim() # Leerzeichen entfernen
    if ([string]::IsNullOrEmpty($domainLocalGroup)) {
        continue # Ignoriert leere Felder
    }
    try {
        if (-not (Get-ADGroup -Filter {Name -eq $domainLocalGroup})) {
            New-ADGroup -Name $domainLocalGroup -GroupScope DomainLocal -Path $ouPath
            Write-Output "Domänenlokale Gruppe $domainLocalGroup wurde erstellt."
        } else {
            Write-Output "Domänenlokale Gruppe $domainLocalGroup existiert bereits."
        }

        # Globalgruppe zu Domänenlokaler Gruppe hinzufügen
        if (-not (Get-ADGroupMember -Identity $domainLocalGroup -Recursive | Where-Object { $_.SamAccountName -eq $globalGroup })) {
            Add-ADGroupMember -Identity $domainLocalGroup -Members $globalGroup
            Write-Output "Globalgruppe $globalGroup wurde zur Domänenlokalen Gruppe $domainLocalGroup hinzugefügt."
        } else {
            Write-Output "Globalgruppe $globalGroup ist bereits Mitglied der Domänenlokalen Gruppe $domainLocalGroup."
        }
    } catch {
        Write-Output "Fehler beim Erstellen oder Verknüpfen der Domänenlokalen Gruppe $domainLocalGroup: $_"
    }
}

Write-Output "Skript erfolgreich ausgeführt!"


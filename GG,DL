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
    $globalGroup = $entry.GlobalGroup
    $domainLocalGroups = $entry.DomainLocalGroups -split "," # DL-Gruppen aufteilen

    if ([string]::IsNullOrEmpty($globalGroup)) {
        continue
    }

    # Globalgruppe erstellen, wenn nicht vorhanden
    $globalGroupObj = Get-ADGroup -Filter {Name -eq $globalGroup}
    if (-not $globalGroupObj) {
        New-ADGroup -Name $globalGroup -GroupScope Global -Path $ouPath
        Write-Output "Globalgruppe $globalGroup wurde erstellt."
    } else {
        Write-Output "Globalgruppe $globalGroup existiert bereits."
    }

    # Domänenlokale Gruppen erstellen und verknüpfen
    foreach ($domainLocalGroup in $domainLocalGroups) {
        $domainLocalGroup = $domainLocalGroup.Trim() # Leerzeichen entfernen
        if ([string]::IsNullOrEmpty($domainLocalGroup)) {
            continue # Ignoriert leere Felder
        }
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
    }
}

Write-Output "Skript erfolgreich ausgeführt!"

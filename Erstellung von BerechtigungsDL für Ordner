# Import-Module und GUI-Komponenten laden
Add-Type -AssemblyName System.Windows.Forms
Import-Module ActiveDirectory

# Funktion, um die aktuelle Domäne auszulesen
function Get-DomainName {
    $domain = (Get-ADDomain).DNSRoot
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
# Freigabepfad auf dem Server abfragen
$sharePath = Read-Host "Bitte den Pfad zur Freigabe eingeben (z.B. \\Server\Freigabe)"
$domain = Get-DomainName

# OU auswählen
$ouPath = Select-OU
if ($ouPath -eq $null) {
    Write-Output "Es wurde keine OU ausgewählt. Skript wird beendet."
    exit
}

# Berechtigungen festlegen
$permissions = @{
    "RO" = [System.Security.AccessControl.FileSystemRights]::Read
    "RW" = [System.Security.AccessControl.FileSystemRights]::ReadAndExecute
    "RX" = [System.Security.AccessControl.FileSystemRights]::Modify
    "FA" = [System.Security.AccessControl.FileSystemRights]::FullControl
}
# Ordner in der Freigabe durchlaufen
Get-ChildItem -Path $sharePath -Directory | ForEach-Object {
    $folder = $_.Name
    $folderPath = $_.FullName
    
    foreach ($permission in $permissions.GetEnumerator()) {
        $groupName = "$folder-$($permission.Key)"

        # DL-Gruppe erstellen, falls nicht vorhanden
        try {
            if (-not (Get-ADGroup -Filter "Name -eq '$groupName'")) {
                New-ADGroup -Name $groupName -GroupScope DomainLocal -GroupCategory Security -Path $ouPath
                Write-Output "Gruppe $groupName wurde erstellt."
            } else {
                Write-Output "Gruppe $groupName existiert bereits."
            }
        } catch {
            $errorMessage = $_.Exception.Message
            Write-Output ("Fehler beim Erstellen der Gruppe {0}: {1}" -f $groupName, $errorMessage)
            continue
        }

        # Berechtigungen setzen
        try {
            $acl = Get-Acl -Path $folderPath
            $permissionRule = New-Object System.Security.AccessControl.FileSystemAccessRule("${domain}\${groupName}", $permission.Value, "ContainerInherit, ObjectInherit", "None", "Allow")
            $acl.AddAccessRule($permissionRule)
            Set-Acl -Path $folderPath -AclObject $acl
            Write-Output "Berechtigungen für $groupName auf $folderPath wurden gesetzt."
        } catch {
            $errorMessage = $_.Exception.Message
            Write-Output ("Fehler beim Setzen der Berechtigungen für {0} auf {1}: {2}" -f $groupName, $folderPath, $errorMessage)
        }
    }
}

Write-Host "Skript erfolgreich ausgeführt!"

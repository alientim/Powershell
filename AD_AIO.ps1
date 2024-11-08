# PowerShell GUI Script für Auswahlbuttons: Create User, Create Groups, ADGDL

# Sicherstellen, dass das Skript als Administrator ausgeführt wird
$myWindowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$myPrincipal = New-Object System.Security.Principal.WindowsPrincipal($myWindowsIdentity)

if (-not $myPrincipal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -ArgumentList $arguments -Verb runAs
    Exit
}

# Windows.Forms Assembly laden
Add-Type -AssemblyName System.Windows.Forms
Import-Module ActiveDirectory

# Funktion, die beim Klick auf "Create User" ausgeführt wird
function Create-User {
    Write-Host "Erstelle Benutzer..."
        $myWindowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $myPrincipal = New-Object System.Security.Principal.WindowsPrincipal($myWindowsIdentity)

        if (-not $myPrincipal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
            $arguments = "& '" + $myinvocation.mycommand.definition + "'"
            Start-Process powershell -ArgumentList $arguments -Verb runAs
            Exit
        }
        Add-Type -AssemblyName System.Windows.Forms
        Import-Module ActiveDirectory

        # Master-OU und Domäne ermitteln
        $masterOU = "OU=Benutzer,DC=domain,DC=com"  # Beispielwert, passe ihn an eure Umgebung an
        $masterGroupOU = "OU=Gruppen,DC=domain,DC=com"  # Beispielwert, passe ihn an eure Umgebung an

        # Holen Sie sich die Domäne aus dem Active Directory
        $domain = (Get-ADDomain).DNSRoot

        # GUI erstellen
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Benutzererstellung"
        $form.Size = New-Object System.Drawing.Size(990, 600)  # Gesamtgröße der Form anpassen

        # Funktion zur Berechnung der maximalen Breite der ComboBox
        function Set-DropDownWidth {
            param (
                [System.Windows.Forms.ComboBox]$comboBox
            )

            $maxLength = 0
            foreach ($item in $comboBox.Items) {
                $itemLength = $item.Length
                if ($itemLength -gt $maxLength) {
                    $maxLength = $itemLength
                }
            }

            # Setze die DropDownWidth entsprechend der maximalen Länge
            $comboBox.DropDownWidth = $maxLength * 6  # Schätzbreite (Zeichen * 8 Pixel)
        }

        # Masterkennwort Eingabefeld
        $masterPasswordLabel = New-Object System.Windows.Forms.Label
        $masterPasswordLabel.Text = "Masterkennwort:"
        $masterPasswordLabel.Location = New-Object System.Drawing.Point(10, 12)
        $masterPasswordLabel.Size = New-Object System.Drawing.Size(100, 20)
        $form.Controls.Add($masterPasswordLabel)

        $masterPasswordTextBox = New-Object System.Windows.Forms.TextBox
        $masterPasswordTextBox.Location = New-Object System.Drawing.Point(110, 10)
        $masterPasswordTextBox.Size = New-Object System.Drawing.Size(200, 20)
        $masterPasswordTextBox.UseSystemPasswordChar = $true  # Passwortfeld
        $form.Controls.Add($masterPasswordTextBox)

        # Master-OU für Benutzer Auswahl
        $masterOULabel = New-Object System.Windows.Forms.Label
        $masterOULabel.Text = "Master-OU (Benutzer):"
        $masterOULabel.Location = New-Object System.Drawing.Point(320, 12)
        $masterOULabel.Size = New-Object System.Drawing.Size(120, 20)
        $form.Controls.Add($masterOULabel)

        $masterOUComboBox = New-Object System.Windows.Forms.ComboBox
        $masterOUComboBox.Location = New-Object System.Drawing.Point(440, 10)
        $masterOUComboBox.Size = New-Object System.Drawing.Size(200, 20)  # ComboBox-Größe festgelegt
        $masterOUComboBox.DropDownStyle = 'DropDownList'

        # OUs aus dem AD für Benutzer abrufen und in die ComboBox einfügen
        $ouList = Get-ADOrganizationalUnit -Filter * | Select-Object -ExpandProperty DistinguishedName
        $masterOUComboBox.Items.AddRange($ouList)
        $masterOUComboBox.SelectedItem = $masterOU  # Standardwerte setzen
        $form.Controls.Add($masterOUComboBox)

        # Berechne und setze die DropDownWidth basierend auf der maximalen Länge
        Set-DropDownWidth -comboBox $masterOUComboBox

        # Master-OU für Gruppen Auswahl
        $masterGroupOULabel = New-Object System.Windows.Forms.Label
        $masterGroupOULabel.Text = "Master-OU (Gruppen):"
        $masterGroupOULabel.Location = New-Object System.Drawing.Point(645, 12)
        $masterGroupOULabel.Size = New-Object System.Drawing.Size(120, 20)
        $form.Controls.Add($masterGroupOULabel)

        $masterGroupOUComboBox = New-Object System.Windows.Forms.ComboBox
        $masterGroupOUComboBox.Location = New-Object System.Drawing.Point(765, 10)
        $masterGroupOUComboBox.Size = New-Object System.Drawing.Size(200, 20)  # ComboBox-Größe für Gruppen
        $masterGroupOUComboBox.DropDownStyle = 'DropDownList'

        # OUs aus dem AD für Gruppen abrufen und in die ComboBox einfügen
        $groupOUList = Get-ADOrganizationalUnit -Filter * | Select-Object -ExpandProperty DistinguishedName
        $masterGroupOUComboBox.Items.AddRange($groupOUList)
        $masterGroupOUComboBox.SelectedItem = $masterGroupOU  # Standardwerte setzen
        $form.Controls.Add($masterGroupOUComboBox)

        # Berechne und setze die DropDownWidth basierend auf der maximalen Länge
        Set-DropDownWidth -comboBox $masterGroupOUComboBox

        # DataGridView erstellen
        $dataGridView = New-Object System.Windows.Forms.DataGridView
        $dataGridView.Size = New-Object System.Drawing.Size(954, 300)  # Breite und Höhe des DataGridViews festgelegt
        $dataGridView.Location = New-Object System.Drawing.Point(10, 40)  # Position des DataGridViews
        $dataGridView.Anchor = [System.Windows.Forms.AnchorStyles]::Top
        $form.Controls.Add($dataGridView)

        # Die Breite der DropDown-Liste manuell anpassen, da DropDownWidth für DataGridViewComboBoxColumn nicht unterstützt wird 
        $maxLength = 0 
        foreach ($item in $masterGroupOUComboBox.Items) {
            $itemLength = $item.ToString().Length
            if ($itemLength -gt $maxLength) {
            $maxLength = $itemLength
            }
        }

        # Definieren der Spalten für die DataGridView
        $dataGridView.ColumnCount = 6
        $dataGridView.Columns[0].Name = "Titel"            # Titel in der ersten Spalte
        $dataGridView.Columns[1].Name = "Vorname"          # Vorname
        $dataGridView.Columns[2].Name = "Nachname"         # Nachname
        $dataGridView.Columns[3].Name = "Globalgruppe"     # Globalgruppe
        $dataGridView.Columns[4].Name = "Standardpasswort" # Standardpasswort
        $dataGridView.Columns[5].Name = "OU"               # OU (wird als ComboBox hinzugefügt)
        $dataGridView.Columns[5].Width = $ouComboBoxColumn

        # Dropdown für die OU in der DataGridView-ComboBox-Spalte
        $ouComboBoxColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
        $ouComboBoxColumn.HeaderText = "OU"
        $ouComboBoxColumn.Items.AddRange($ouList)  # Hier fügen wir alle OUs hinzu
        $dataGridView.Columns.RemoveAt(5)  # Entfernt die ursprüngliche OU-Spalte
        $dataGridView.Columns.Insert(5, $ouComboBoxColumn)  # Fügt die ComboBox-Spalte an der richtigen Stelle ein
        $ouComboBoxColumn.Width = $maxLength * 6

        # RichTextBox für Ausgaben (anstelle von TextBox)
        $outputTextBox = New-Object System.Windows.Forms.RichTextBox
        $outputTextBox.Multiline = $true
        $outputTextBox.Location = New-Object System.Drawing.Point(10, 350)  # Position unterhalb des DataGridViews
        $outputTextBox.Size = New-Object System.Drawing.Size(954, 150)  # Größe der RichTextBox festgelegt
        $outputTextBox.ScrollBars = 'Vertical'
        $form.Controls.Add($outputTextBox)

        # OK-Button erstellen
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "Benutzer erstellen"
        $okButton.Size = New-Object System.Drawing.Size(380, 30)  # Größe des Buttons festgelegt
        $okButton.Location = New-Object System.Drawing.Point(10, 510)  # Position des Buttons unter der TextBox

        # Beenden-Button erstellen
        $exitButton = New-Object System.Windows.Forms.Button
        $exitButton.Text = "Beenden"
        $exitButton.Size = New-Object System.Drawing.Size(380, 30)  # Größe des Buttons festgelegt
        $exitButton.Location = New-Object System.Drawing.Point(585, 510)  # Position des Beenden-Buttons

        # Buttons zum Formular hinzufügen
        $form.Controls.Add($okButton)
        $form.Controls.Add($exitButton)

        # Importiere das Active Directory Modul, falls noch nicht geschehen
        Import-Module ActiveDirectory

        # Event-Handler für den OK-Button hinzufügen
        $okButton.Add_Click({
            $masterPassword = $masterPasswordTextBox.Text  # Masterkennwort holen
            foreach ($row in $dataGridView.Rows) {
                if ($row.Index -lt $dataGridView.RowCount - 1) {  # Nicht für die leere letzte Zeile
                    $title = $row.Cells[0].Value
                    $firstName = $row.Cells[1].Value
                    $lastName = $row.Cells[2].Value
                    $globalGroup = $row.Cells[3].Value
                    $password = $row.Cells[4].Value
                    $ou = $row.Cells[5].Value  # Hier wird die ausgewählte OU aus der ComboBox abgerufen

                    # Wenn keine OU in der Zeile gewählt wurde, nehme die Master-OU
                    if (-not $ou) {
                        $ou = $masterOUComboBox.SelectedItem
                    }

                    # Wenn kein Passwort angegeben wurde, nutze das Masterkennwort
                    if (-not $password) {
                        $password = $masterPassword
                    }

                    $groupOU = $masterGroupOUComboBox.SelectedItem

                    # Überprüfen, ob notwendige Felder ausgefüllt sind
                    if (-not $firstName -or -not $lastName) {
                        $missingField = if (-not $firstName) { "Vorname" } elseif (-not $lastName) { "Nachname" }
                        $outputTextBox.AppendText("Fehler bei der Erstellung des Benutzers '$firstName $lastName': $missingField fehlt.`r`n")
                        continue
                    }
            
                    # Benutzername generieren: erster Buchstabe des Vornamens + Nachname
                    $username = ($firstName.Substring(0, 1) + $lastName).ToLower()

                    # E-Mail-Adresse generieren
                    $email = Get-EmailAddress -username $username

                    # Benutzer erstellen
                    try {
                        # Versuchen, den Benutzer abzurufen (auch wenn er bereits existiert)
                        $user = Get-ADUser -Filter { SamAccountName -eq $username }
                        if (-not $user) {
                            # Benutzer erstellen, falls er nicht existiert
                            try {
                                New-ADUser -Name "$firstName $lastName" `
                                           -GivenName "$firstName" `
                                           -Surname "$lastName" `
                                           -SamAccountName "$username" `
                                           -UserPrincipalName "$email" `
                                           -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) `
                                           -Enabled $true `
                                           -Path "$ou"
                                   
                                $outputTextBox.SelectionColor = 'Green'
                                $outputTextBox.AppendText("Benutzer $username wurde erfolgreich erstellt.`r`n")
                            }
                            catch {
                                $outputTextBox.SelectionColor = 'Red'
                                $outputTextBox.AppendText("Fehler bei der Erstellung des Benutzers '$firstName $lastName':`r`n")
                                $outputTextBox.AppendText("Fehlerdetails: $_`r`n")
                            }
                        }
                        else {
                            # Erfolgsnachricht für vorhandenen Benutzer
                            $outputTextBox.SelectionColor = 'RED'
                            $outputTextBox.AppendText("Benutzer $username existiert bereits.`r`n")
                            $outputTextBox.AppendText("Fehlerdetails: $_`r`n")
                        }

                        # Gruppenzuordnung durchführen, auch wenn der Benutzer schon existiert
                        if ($globalGroup) {
                            # Gruppen-OU immer auf die Master-OU setzen
                            try {
                                # Überprüfen, ob die Gruppe existiert
                                $groupOU = $masterGroupOUComboBox.SelectedItem
                                $group = Get-ADGroup -Filter { Name -eq $globalGroup }
                                try {
                                    # Gruppe erstellen, falls sie nicht existiert
                                    $groupName = "GG_" + $globalGroup
                                    New-ADGroup -Name $groupName `
                                                -GroupScope Global `
                                                -Path $groupOU `
                                                -Description "Globale Gruppe für $groupName"
                            
                                    # Erfolgsnachricht
                                    $outputTextBox.SelectionColor = 'Green'
                                    $outputTextBox.AppendText("Globale Gruppe '$groupName' wurde erfolgreich erstellt.`r`n")
                                }
                                catch {
                                    $outputTextBox.SelectionColor = 'Red'
                                    $outputTextBox.AppendText("Fehler bei der Erstellung der Gruppe '$groupName':`r`n")
                                    $outputTextBox.AppendText("Fehlerdetails: $_`r`n")
                                }
                            }
                            catch {
                                # Fehler bei der Gruppenzuordnung
                                $outputTextBox.SelectionColor = 'Red'
                                $outputTextBox.AppendText("Fehler bei der Gruppenzuordnung für Benutzer '$username':`r`n")
                                $outputTextBox.AppendText("Fehlerdetails: $_`r`n")
                            }
                        }
                    }
                    catch {
                        # Fehler bei der Benutzererstellung
                        $outputTextBox.SelectionColor = 'Red'
                        $outputTextBox.AppendText("Fehler bei der Erstellung des Benutzers '$firstName $lastName':`r`n")
                        $outputTextBox.AppendText("Fehlerdetails: $_`r`n")
                    }
                }
            }
        })

        # Event-Handler für den Beenden-Button hinzufügen
        $exitButton.Add_Click({
            $form.Close()  # Formular schließen
        })

        # Funktion zum Erstellen der E-Mail-Adresse
        function Get-EmailAddress {
            param (
                [string]$username
            )
    
            return "$username@$domain"
        }

        # Formular anzeigen
        $form.ShowDialog()
}

# Funktion, die beim Klick auf "Create Groups" ausgeführt wird
function Create-Groups {
    Write-Host "Erstelle Gruppen..."
        Add-Type -AssemblyName System.Windows.Forms
        Import-Module ActiveDirectory

        # Holen Sie sich die Domäne aus dem Active Directory
        $domain = (Get-ADDomain).DNSRoot

        # GUI erstellen
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Gruppen Erstellung"
        $form.Size = New-Object System.Drawing.Size(600, 555)  # Gesamtgröße der Form anpassen

        # DataGridView für Gruppen erstellen
        $dataGridView = New-Object System.Windows.Forms.DataGridView
        $dataGridView.Size = New-Object System.Drawing.Size(564, 300)  # Breite und Höhe des DataGridViews
        $dataGridView.Location = New-Object System.Drawing.Point(10, 10)  # Position des DataGridViews
        $form.Controls.Add($dataGridView)

        # Definieren der Spalten für die DataGridView
        $dataGridView.ColumnCount = 3
        $dataGridView.Columns[0].Name = "Typ"          # Typ in der ersten Spalte
        $dataGridView.Columns[1].Name = "Gruppenname"  # Gruppenname in der zweiten Spalte
        $dataGridView.Columns[2].Name = "OU"           # OU in der dritten Spalte

        # Breite der Spalten festlegen
        $dataGridView.Columns[0].Width = 75   # Typ-Spalte auf 50 Pixel setzen
        $dataGridView.Columns[1].Width = 100  # Gruppenname-Spalte auf 100 Pixel setzen
        $dataGridView.Columns[2].Width = 345  # OU-Spalte auf 400 Pixel setzen

        # Dropdown für den Typ (GG oder DL)
        $typColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
        $typColumn.HeaderText = "Typ"
        $typColumn.Items.Add("GG")
        $typColumn.Items.Add("DL")
        $typColumn.Items.Add("DL_OS")
        $dataGridView.Columns.RemoveAt(0)
        $dataGridView.Columns.Insert(0, $typColumn)
        # Breite der Typ-Spalte setzen
        $typColumn.Width = 75  # Du kannst hier den Wert nach Bedarf anpassen

        # Dropdown für die OU-Auswahl in der dritten Spalte
        $ouColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
        $ouColumn.HeaderText = "OU"
        $groupOUList = Get-ADOrganizationalUnit -Filter * | Select-Object -ExpandProperty DistinguishedName
        foreach ($ou in $groupOUList) {
            $ouColumn.Items.Add($ou)
        }
        $dataGridView.Columns.RemoveAt(2)
        $dataGridView.Columns.Insert(2, $ouColumn)

        # Breite der OU-Spalte setzen
        $ouColumn.Width = 345  # Du kannst hier den Wert nach Bedarf anpassen

        # RichTextBox für Ausgaben
        $outputTextBox = New-Object System.Windows.Forms.RichTextBox
        $outputTextBox.Multiline = $true
        $outputTextBox.Location = New-Object System.Drawing.Point(10, 320)  # Position unterhalb des DataGridViews
        $outputTextBox.Size = New-Object System.Drawing.Size(564, 150)  # Größe der RichTextBox
        $outputTextBox.ScrollBars = 'Vertical'
        $form.Controls.Add($outputTextBox)

        # OK-Button für Gruppen erstellen
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "Gruppen erstellen"
        $okButton.Size = New-Object System.Drawing.Size(280, 30)  # Größe des Buttons
        $okButton.Location = New-Object System.Drawing.Point(10, 475)  # Position des Buttons
        $okButton.Add_Click({
            foreach ($row in $dataGridView.Rows) {
                if ($row.Index -lt $dataGridView.RowCount - 1) {  # Nicht für die leere letzte Zeile
                    $groupType = $row.Cells[0].Value
                    $groupName = $row.Cells[1].Value
                    $groupOU = $row.Cells[2].Value

                    if (-not $groupType -or -not $groupName -or -not $groupOU) {
                        $outputTextBox.SelectionColor = 'Red'
                        $outputTextBox.AppendText("Fehler: Alle Felder (Typ, Gruppenname, OU) mÃ¼ssen ausgefÃ¼llt sein.`r`n")
                        continue
                    }

                    # Präfix vor den Gruppennamen setzen, basierend auf dem Typ
                    if ($groupType -eq "GG") {
                        $groupName = "GG_" + $groupName  # Für globale Gruppen "GG_" voranstellen
                    } elseif ($groupType -eq "DL_OS") {
                        $groupName = "DL_" + $groupName  # Für globale Gruppen "DL_OS" voranstellen
                    } elseif ($groupType -eq "DL") {
                        $groupName = "DL_" + $groupName  # Für DomainLocal Gruppen "DL_" voranstellen
                    }

                    try {
                        if ($groupType -eq "GG") {
                            # Globale Gruppe erstellen (keine Suffixe, aber Präfix wird hinzugefügt)
                            $group = Get-ADGroup -Filter { Name -eq $groupName }
                            if (-not $group) {
                                New-ADGroup -Name $groupName `
                                            -GroupScope Global `
                                            -Path $groupOU `
                                            -Description "Globale Gruppe fÃ¼r $groupName"
                                $outputTextBox.SelectionColor = 'Green'
                                $outputTextBox.AppendText("Globale Gruppe '$groupName' wurde erfolgreich erstellt.`r`n")
                            } else {
                                $outputTextBox.SelectionColor = 'Green'
                                $outputTextBox.AppendText("Globale Gruppe '$groupName' existiert bereits.`r`n")
                            }
                        } elseif ($groupType -eq "DL_OS") {
                            # DomainLocal Gruppen ohne Suffix erstellen
                                $domainLocalGroupName = "$groupName"
                                if (-not $group) {
                                    New-ADGroup -Name $domainLocalGroupName `
                                                -GroupScope DomainLocal `
                                                -Path $groupOU `
                                                -Description "DomainLocal Gruppe fÃ¼r $domainLocalGroupName"
                                    $outputTextBox.SelectionColor = 'Green'
                                    $outputTextBox.AppendText("DomainLocal Gruppe '$domainLocalGroupName' wurde erfolgreich erstellt.`r`n")
                                } else {
                                    $outputTextBox.SelectionColor = 'Green'
                                    $outputTextBox.AppendText("DomainLocal Gruppe '$domainLocalGroupName' existiert bereits.`r`n")
                                }
                        } elseif ($groupType -eq "DL") {
                            # DomainLocal Gruppen mit Suffixen erstellen
                            $suffixes = "_FA", "_RW", "_RX", "_RO"
                            foreach ($suffix in $suffixes) {
                                $domainLocalGroupName = "$groupName$suffix"
                                $group = Get-ADGroup -Filter { Name -eq $domainLocalGroupName }
                                if (-not $group) {
                                    New-ADGroup -Name $domainLocalGroupName `
                                                -GroupScope DomainLocal `
                                                -Path $groupOU `
                                                -Description "DomainLocal Gruppe fÃ¼r $domainLocalGroupName"
                                    $outputTextBox.SelectionColor = 'Green'
                                    $outputTextBox.AppendText("DomainLocal Gruppe '$domainLocalGroupName' wurde erfolgreich erstellt.`r`n")
                                } else {
                                    $outputTextBox.SelectionColor = 'Green'
                                    $outputTextBox.AppendText("DomainLocal Gruppe '$domainLocalGroupName' existiert bereits.`r`n")
                                }
                            }
                        }
                    } catch {
                        # Fehler bei der Gruppen-Erstellung
                        $outputTextBox.SelectionColor = 'Red'
                        $outputTextBox.AppendText("Fehler bei der Erstellung der Gruppe '$groupName': $_.Exception.Message`r`n")
                    }
                }
            }
        })

        # Beenden-Button erstellen
        $exitButton = New-Object System.Windows.Forms.Button
        $exitButton.Text = "Beenden"
        $exitButton.Size = New-Object System.Drawing.Size(280, 30)
        $exitButton.Location = New-Object System.Drawing.Point(295, 475)  # Position des Beenden-Buttons
        $exitButton.Add_Click({
            $form.Close()  # Formular schließen
        })

        # Buttons zum Formular hinzufügen
        $form.Controls.Add($okButton)
        $form.Controls.Add($exitButton)

        # Formular anzeigen
        $form.ShowDialog()
}

# Funktion, die beim Klick auf "ADGDL" ausgeführt wird
function ADGDL {
    Write-Host "Führe AGDLP aus..."
        Add-Type -AssemblyName System.Windows.Forms
        Import-Module ActiveDirectory

        # Funktion zum Abrufen der aktuellen DomÃ¤ne
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
        $dlLabel.Text = "DomÃ¤nenlokale Gruppen (DL):"
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

        # RichTextBox für Debug-Informationen hinzufÃ¼gen
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
                    Add-DebugMessage "Benutzer $user wurde der Globalgruppe $gg hinzugefÃ¼gt." "Green"
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
                            Add-DebugMessage "Globalgruppe $gg wurde zur Globalgruppe $dlGroup hinzugefÃ¼gt." "Green"
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
                            Add-DebugMessage "Globalgruppe $gg wurde zur DomÃ¤nenlokalen Gruppe $dlGroup hinzugefÃ¼gt." "Green"
                        } else {
                            Add-DebugMessage "Globalgruppe $gg ist bereits Mitglied der DomÃ¤nenlokalen Gruppe $dlGroup." "Orange"
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
}

# Erstellen des Formulars
$form = New-Object Windows.Forms.Form
$form.Text = 'Administrator Tools'
$form.Size = New-Object Drawing.Size(300, 200)

# Erstellen des Buttons für "Create User"
$btnCreateUser = New-Object Windows.Forms.Button
$btnCreateUser.Text = 'Create User'
$btnCreateUser.Size = New-Object Drawing.Size(250, 40)
$btnCreateUser.Location = New-Object Drawing.Point(20, 10)
$btnCreateUser.Add_Click({ Create-User })

# Erstellen des Buttons für "Create Groups"
$btnCreateGroups = New-Object Windows.Forms.Button
$btnCreateGroups.Text = 'Create Groups'
$btnCreateGroups.Size = New-Object Drawing.Size(250, 40)
$btnCreateGroups.Location = New-Object Drawing.Point(20, 60)
$btnCreateGroups.Add_Click({ Create-Groups })

# Erstellen des Buttons für "AGDLP"
$btnADGDL = New-Object Windows.Forms.Button
$btnADGDL.Text = 'AGDLP'
$btnADGDL.Size = New-Object Drawing.Size(250, 40)
$btnADGDL.Location = New-Object Drawing.Point(20, 110)
$btnADGDL.Add_Click({ ADGDL })

# Hinzufügen der Buttons zum Formular
$form.Controls.Add($btnCreateUser)
$form.Controls.Add($btnCreateGroups)
$form.Controls.Add($btnADGDL)

# Anzeigen des Formulars
$form.ShowDialog()

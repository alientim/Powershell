# PowerShell GUI Script für Auswahlbuttons: Create User, Create Groups, ADGDL

# Version 1.0:
# Created 2024 by Tim Eertmoed, Germany to work on Windows Server 2019/2022 as an user creating script.

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
        $outputTextBox.ReadOnly = $true
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
						# Versuchen, den Benutzer abzurufen (auch wenn er bereits existiert)
						$user = Get-ADUser -Filter { SamAccountName -eq $username }

						if (-not $user) {
							# Benutzer erstellen, falls er nicht existiert
							try {
								# Erstelle die Parameter für New-ADUser
								$createUserParams = @{
									Name              = "$firstName $lastName"
									GivenName         = "$firstName"
									Surname           = "$lastName"
									SamAccountName    = "$username"
									UserPrincipalName = "$email"
									AccountPassword   = (ConvertTo-SecureString $password -AsPlainText -Force)
									Enabled           = $true
									Path              = "$ou"
								}

								# Überprüfen, ob der Jobtitel gesetzt wurde und hinzufügen
								if ($title) {
									$createUserParams["Title"] = $title
								}

								# Benutzer erstellen
								New-ADUser @createUserParams

								# Erfolgsnachricht
								$outputTextBox.SelectionColor = 'Green'
								$outputTextBox.AppendText("Benutzer $username wurde erfolgreich erstellt.`r`n")
							}
							catch {
								# Fehler bei der Benutzererstellung
								$outputTextBox.SelectionColor = 'Red'
								$outputTextBox.AppendText("Fehler bei der Erstellung des Benutzers '$firstName $lastName':`r`n")
								$outputTextBox.AppendText("Fehlerdetails: $_`r`n")
                                continue
							}
						}
						else {
							# Benutzer existiert bereits
							$outputTextBox.SelectionColor = 'Orange'
							$outputTextBox.AppendText("Benutzer $username existiert bereits.`r`n")
						}
                            # Gruppenzuordnung durchführen, auch wenn der Benutzer schon existiert
                        if ($globalGroup) {
                            # Gruppen-OU immer auf die Master-OU setzen
                            try {
                                # Der Name der Gruppe mit dem Präfix GG_ erstellen
                                $groupName = "GG_" + $globalGroup

                                # Überprüfen, ob die Gruppe bereits existiert
                                $group = Get-ADGroup -Filter { Name -eq $groupName } -ErrorAction SilentlyContinue

                                # Wenn die Gruppe nicht existiert, wird sie erstellt
                                if (-not $group) {
                                    try {
                                        # Erstelle die Gruppe
                                        New-ADGroup -Name $groupName `
                                                    -GroupScope Global `
                                                    -Path $groupOU `
                                                    -Description "Globale Gruppe für $groupName"
                
                                        # Erfolgsnachricht für Gruppenerstellung
                                        $outputTextBox.SelectionColor = 'Green'
                                        $outputTextBox.AppendText("Globale Gruppe '$groupName' wurde erfolgreich erstellt.`r`n")
                                    }
                                    catch {
                                        # Fehler bei der Erstellung der Gruppe
                                        $outputTextBox.SelectionColor = 'Red'
                                        $outputTextBox.AppendText("Fehler bei der Erstellung der Gruppe '$groupName':`r`n")
                                        $outputTextBox.AppendText("Fehlerdetails: $_`r`n")
                                        continue
                                    }
                                }
                                else {
                                    # Erfolgsnachricht, falls die Gruppe bereits existiert
                                    $outputTextBox.SelectionColor = 'Orange'
                                    $outputTextBox.AppendText("Gruppe '$groupName' existiert bereits.`r`n")
                                }

                                # Benutzer zur Gruppe hinzufügen (nur wenn der Benutzername definiert ist)
                                if ($username) {
                                    try {
                                        Add-ADGroupMember -Identity $groupName -Members $username
                                
                                        # Erfolgsnachricht für das Hinzufügen des Benutzers
                                        $outputTextBox.SelectionColor = 'Green'
                                        $outputTextBox.AppendText("Benutzer '$username' wurde erfolgreich zur Gruppe '$groupName' hinzugefügt.`r`n")
                                    }
                                    catch {
                                        # Fehler bei der Hinzufügung des Benutzers
                                        $outputTextBox.SelectionColor = 'Red'
                                        $outputTextBox.AppendText("Fehler bei der Hinzufügung des Benutzers '$username' zur Gruppe '$groupName':`r`n")
                                        $outputTextBox.AppendText("Fehlerdetails: $_`r`n")
                                        continue
                                    }
                                } else {
                                    $outputTextBox.SelectionColor = 'Red'
                                    $outputTextBox.AppendText("Benutzername '$username' ist nicht definiert.`r`n")
                                }
                            }
                            catch {
                                # Fehler bei der Gruppenzuordnung oder übergeordneten Fehler
                                $outputTextBox.SelectionColor = 'Red'
                                $outputTextBox.AppendText("Fehler bei der Verarbeitung der Gruppe '$globalGroup':`r`n")
                                $outputTextBox.AppendText("Fehlerdetails: $_`r`n")
                                continue
                            }
                        }
                    }
                    catch {
                        # Fehler bei der Benutzererstellung
                        $outputTextBox.SelectionColor = 'Red'
                        $outputTextBox.AppendText("Fehler bei der Erstellung des Benutzers '$firstName $lastName':`r`n")
                        $outputTextBox.AppendText("Fehlerdetails: $_`r`n")
                        continue
                    }
                }
            }
            [System.Windows.Forms.MessageBox]::Show("Vorgang abgeschlossen.")
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
        $outputTextBox.ReadOnly = $true
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
                        $outputTextBox.AppendText("Fehler: Alle Felder (Typ, Gruppenname, OU) müssen ausgefüllt sein.`r`n")
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
                                            -Description "Globale Gruppe für $groupName"
                                $outputTextBox.SelectionColor = 'Green'
                                $outputTextBox.AppendText("Globale Gruppe '$groupName' wurde erfolgreich erstellt.`r`n")
                            } else {
                                $outputTextBox.SelectionColor = 'Orange'
                                $outputTextBox.AppendText("Globale Gruppe '$groupName' existiert bereits.`r`n")
                            }
                        } elseif ($groupType -eq "DL_OS") {
                            # DomainLocal Gruppen ohne Suffix erstellen
                                $domainLocalGroupName = "$groupName"
                                $group = Get-ADGroup -Filter { Name -eq $domainLocalGroupName }
                                if (-not $group) {
                                    New-ADGroup -Name $domainLocalGroupName `
                                                -GroupScope DomainLocal `
                                                -Path $groupOU `
                                                -Description "DomainLocal Gruppe für $domainLocalGroupName"
                                    $outputTextBox.SelectionColor = 'Green'
                                    $outputTextBox.AppendText("DomainLocal Gruppe '$domainLocalGroupName' wurde erfolgreich erstellt.`r`n")
                                } else {
                                    $outputTextBox.SelectionColor = 'Orange'
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
                                                -Description "DomainLocal Gruppe für $domainLocalGroupName"
                                    $outputTextBox.SelectionColor = 'Green'
                                    $outputTextBox.AppendText("DomainLocal Gruppe '$domainLocalGroupName' wurde erfolgreich erstellt.`r`n")
                                } else {
                                    $outputTextBox.SelectionColor = 'Orange'
                                    $outputTextBox.AppendText("DomainLocal Gruppe '$domainLocalGroupName' existiert bereits.`r`n")
                                }
                            }
                        }
                    } catch {
                        # Fehler bei der Gruppen-Erstellung
                        $outputTextBox.SelectionColor = 'Red'
                        $outputTextBox.AppendText("Fehler bei der Erstellung der Gruppe '$groupName': $_.Exception.Message`r`n")
                        continue
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
        $form.Size = New-Object System.Drawing.Size(805, 505)  # Fenstergröße anpassen

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
        $debugTextBox.Multiline = $true
        $debugTextBox.Location = New-Object System.Drawing.Point(10, 265)
        $debugTextBox.Size = New-Object System.Drawing.Size(770, 150)
        $debugTextBox.ScrollBars = 'Vertical'
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
        $okButton.Location = New-Object System.Drawing.Point(10, 425)
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
        $exitButton.Location = New-Object System.Drawing.Point(400, 425)
        $exitButton.Size = New-Object System.Drawing.Size(380, 30)
        $exitButton.Add_Click({
            $form.Close()
        })
        $form.Controls.Add($exitButton)

        $form.ShowDialog()
}

# Funktion, die beim Klick auf "SMB-Rechte" ausgeführt wird
function SMB-RIGHTS {
    Write-Host "Führe SMB Rechte aus..."
        Add-Type -AssemblyName System.Windows.Forms
        Import-Module ActiveDirectory

        # Hauptform erstellen
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Verzeichnis- und OU-Auswahl"
        $form.Size = New-Object System.Drawing.Size(648, 510)

        # RichTextBox für Debug-Informationen hinzufügen
        $debugTextBox = New-Object System.Windows.Forms.RichTextBox
        $debugTextBox.Multiline = $true
        $debugTextBox.Location = New-Object System.Drawing.Point(10, 280)
        $debugTextBox.Size = New-Object System.Drawing.Size(610, 150)
        $debugTextBox.ScrollBars = 'Vertical'
        $debugTextBox.ReadOnly = $true
        $form.Controls.Add($debugTextBox)

        # Liste für neu angelegte DLs
        $newDLs = @()

        # TreeView für die Ordnerstruktur (links)
        $folderTreeView = New-Object System.Windows.Forms.TreeView
        $folderTreeView.Location = New-Object System.Drawing.Point(10, 10)
        $folderTreeView.Size = New-Object System.Drawing.Size(300, 200)

        # Setze die folgenden Eigenschaften für den schreibgeschützten Zustand
        $folderTreeView.LabelEdit = $false  # Verhindert das Bearbeiten von Knoten
        $folderTreeView.AllowDrop = $false  # Verhindert Drag & Drop
        $folderTreeView.HideSelection = $true  # Deaktiviert die Auswahlanzeige
        $folderTreeView.FullRowSelect = $false  # Verhindert die vollständige Zeilenmarkierung
        $folderTreeView.SelectedNode = $null  # Setzt die Auswahl zurück
        $form.Controls.Add($folderTreeView)

        # TextBox für den Hauptordner
        $folderTextBox = New-Object System.Windows.Forms.TextBox
        $folderTextBox.Location = New-Object System.Drawing.Point(10, 215)
        $folderTextBox.Size = New-Object System.Drawing.Size(300, 30)
        $folderTextBox.ReadOnly = $true
        $form.Controls.Add($folderTextBox)

        # Button für Ordnerauswahl
        $folderButton = New-Object System.Windows.Forms.Button
        $folderButton.Text = "Hauptordner auswählen"
        $folderButton.Location = New-Object System.Drawing.Point(10, 240)
        $folderButton.Size = New-Object System.Drawing.Size(300, 30)
        $form.Controls.Add($folderButton)

        # Ordnerauswahl-Funktion mit FolderBrowserDialog
        $folderButton.Add_Click({
            $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
            $folderBrowser.Description = "Wählen Sie den Hauptordner aus, in dem die Unterordner liegen"

            if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                # Zeige den ausgewählten Ordner in der TextBox
                $folderTextBox.Text = $folderBrowser.SelectedPath
                $debugTextBox.AppendText("Hauptordner ausgewählt: $($folderTextBox.Text)" + [Environment]::NewLine)

                # Leere die TreeView und lade den Ordnerbaum
                $folderTreeView.Nodes.Clear()

                # Füge die Wurzel des Ordners als Node hinzu
                $rootNode = New-Object System.Windows.Forms.TreeNode($folderBrowser.SelectedPath)
                $rootNode.Tag = $folderBrowser.SelectedPath  # Pfad im Tag speichern
                $folderTreeView.Nodes.Add($rootNode)

                # Lade alle Unterordner des ausgewählten Ordners bis zur 2. Ebene (z.B. max. 2 Ebenen)
                Load-Subfolders $folderBrowser.SelectedPath $rootNode 2  # Maximal 2 Ebenen laden

                # Expandiere den Root-Node
                $rootNode.Expand()
            }
        })

        # Funktion zum Laden der Unterordner (max. 2 Ebenen)
        function Load-Subfolders {
            param (
                [string]$parentFolder,
                [System.Windows.Forms.TreeNode]$parentNode,
                [int]$maxDepth,
                [int]$currentDepth = 0
            )

            # Wenn die maximale Tiefe erreicht ist, breche ab
            if ($currentDepth -ge $maxDepth) {
                return
            }

            try {
                # Überprüfen, ob der Ordner existiert
                if (Test-Path $parentFolder) {
                    # Hole alle Unterordner des aktuellen Ordners
                    $subfolders = Get-ChildItem -Path $parentFolder -Directory -ErrorAction SilentlyContinue

                    # Füge alle Unterordner als Knoten hinzu
                    foreach ($subfolder in $subfolders) {
                        # Erstelle den Knoten für den Unterordner
                        $node = New-Object System.Windows.Forms.TreeNode($subfolder.Name)
                        $node.Tag = $subfolder.FullName  # Pfad des Unterordners im Tag speichern
                        $parentNode.Nodes.Add($node)

                        # Lade Unterordner rekursiv für diesen Knoten (mit neuer Tiefe)
                        Load-Subfolders $subfolder.FullName $node $maxDepth ($currentDepth + 1)
                    }
                } else {
                    $debugTextBox.AppendText("Pfad nicht gefunden: $parentFolder" + [Environment]::NewLine)
                }
            } catch {
                $debugTextBox.AppendText("Fehler beim Laden der Ordnerstruktur: $_" + [Environment]::NewLine)
            }
        }

        # Abschnitt für OU Auswahl

        # ListBox für OUs
        $ouListBox = New-Object System.Windows.Forms.ListBox
        $ouListBox.Location = New-Object System.Drawing.Point(320, 10)
        $ouListBox.Size = New-Object System.Drawing.Size(300, 205)
        $ouListBox.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
        $form.Controls.Add($ouListBox)

        # TextBox für die ausgewählte OU
        $ouTextBox = New-Object System.Windows.Forms.TextBox
        $ouTextBox.Location = New-Object System.Drawing.Point(320, 215)
        $ouTextBox.Size = New-Object System.Drawing.Size(300, 30)
        $ouTextBox.ReadOnly = $true
        $form.Controls.Add($ouTextBox)

        # Button für OU-Auswahl
        $ouButton = New-Object System.Windows.Forms.Button
        $ouButton.Text = "OU auswählen"
        $ouButton.Location = New-Object System.Drawing.Point(320, 240)
        $ouButton.Size = New-Object System.Drawing.Size(300, 30)
        $form.Controls.Add($ouButton)

        # Event für den Button: Wenn der Button geklickt wird, öffnet sich das Fenster zur OU-Auswahl
        $ouButton.Add_Click({
            # Funktion zur OU-Auswahl
            function Select-OU {
                Add-Type -AssemblyName System.Windows.Forms
                $ous = Get-ADOrganizationalUnit -Filter *
                $ouNames = $ous | ForEach-Object { $_.DistinguishedName }

                $ouForm = New-Object System.Windows.Forms.Form
                $ouForm.Text = "OU auswählen"
                $ouForm.Size = New-Object System.Drawing.Size(400, 300)

                $listBox = New-Object System.Windows.Forms.ListBox
                $listBox.Dock = "Fill"
                $listBox.Items.AddRange($ouNames)
                $ouForm.Controls.Add($listBox)

                $okButton = New-Object System.Windows.Forms.Button
                $okButton.Size = New-Object System.Drawing.Size(300, 30)
                $okButton.Text = "OK"
                $okButton.Dock = "Bottom"
                $okButton.Add_Click({
                    if ($listBox.SelectedItem) {
                        $ouForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    } else {
                        [System.Windows.Forms.MessageBox]::Show("Bitte wählen Sie eine OU aus.")
                    }
                })
                $ouForm.Controls.Add($okButton)

                # Zeige das Auswahlfenster und gebe die ausgewählte OU zurück
                if ($ouForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    return $listBox.SelectedItem
                } else {
                    return $null
                }
            }

            # Rufe die Select-OU Funktion auf und speichere die gewählte OU in der TextBox
            $selectedOU = Select-OU
            if ($selectedOU) {
                $ouTextBox.Text = $selectedOU  # Setze die ausgewählte OU in die TextBox
                $debugTextBox.AppendText("Ausgewählte OU: $selectedOU" + [Environment]::NewLine)
            }
        })

        # Funktion zum Laden der OUs in die ListBox
        function Load-OUs {
            try {
                $ous = Get-ADOrganizationalUnit -Filter * | Select-Object -ExpandProperty DistinguishedName
                $ouListBox.Items.Clear()
                $ouListBox.Items.AddRange($ous)
            } catch {
                $debugTextBox.AppendText("Fehler beim Laden der OUs: $_" + [Environment]::NewLine)
            }
        }

        Load-OUs

        # Event für ListBox: Wenn eine OU ausgewählt wird
        $ouListBox.Add_SelectedIndexChanged({
            if ($ouListBox.SelectedItem) {
                $ouTextBox.Text = $ouListBox.SelectedItem
            }
        })

        # Schließen Button
        $closeButton = New-Object System.Windows.Forms.Button
        $closeButton.Text = "Beenden"
        $closeButton.Location = New-Object System.Drawing.Point(320, 435)  # Direkt unterhalb der TextBox
        $closeButton.Size = New-Object System.Drawing.Size(300, 30)
        $closeButton.Add_Click({
            $form.Close()
        })
        $form.Controls.Add($closeButton)

        # Starten Button (sichtbar und korrekt positioniert)
        $startButton = New-Object System.Windows.Forms.Button
        $startButton.Text = "Starten"
        $startButton.Location = New-Object System.Drawing.Point(10, 435)  # Direkt unterhalb des TextBox
        $startButton.Size = New-Object System.Drawing.Size(300, 30)
        $startButton.Add_Click({
            $mainFolderPath = $folderTextBox.Text
            $ouPath = $ouTextBox.Text
    
            if (-not $mainFolderPath) {
                # Fehlernachricht (schwarz)
                Add-DebugMessage -message "Bitte wählen Sie einen gültigen Ordner aus." -isPositive $false
                return
            }

            if (-not $ouPath) {
                # Fehlernachricht (schwarz)
                Add-DebugMessage -message "Bitte wählen Sie eine gültige OU aus." -isPositive $false
                return
            }

            # Berechtigungen definieren
            $permissions = @{
                "RO" = [System.Security.AccessControl.FileSystemRights]::Read
                "RX" = [System.Security.AccessControl.FileSystemRights]::ReadAndExecute
                "RW" = [System.Security.AccessControl.FileSystemRights]::Modify
                "FA" = [System.Security.AccessControl.FileSystemRights]::FullControl
            }

            # Liste für neu angelegte DLs und Rechte
            $newDLs = @()
            $newRights = @()

            # Unterordner im Hauptordner durchlaufen
            $subfolders = Get-ChildItem -Path $mainFolderPath -Directory
            foreach ($subfolder in $subfolders) {
                $subfolderName = $subfolder.Name.Trim()
                $subfolderNameClean = $subfolderName -replace ' ', '_'
        
                foreach ($permissionName in $permissions.Keys) {
                    $dlGroupName = "DL_${subfolderNameClean}_${permissionName}"
            
                    # DL-Gruppe erstellen, wenn nicht vorhanden
                    $dlGroup = Get-ADGroup -Filter {Name -eq $dlGroupName}
                    if (-not $dlGroup) {
                        try {
                            New-ADGroup -Name $dlGroupName -GroupScope DomainLocal -Path $ouPath
                            $newDLs += $dlGroupName  # Hinzufügen des DL-Namens zur Liste
                            # Erfolgsnachricht (grün)
                        } catch {
                            # Fehlernachricht (schwarz)
                            Add-DebugMessage -message "Fehler bei der Erstellung der Gruppe '$dlGroupName': $_" -isPositive $false
                        }
                    }

                    # Berechtigungen auf Ordner setzen
                    try {
                        $acl = Get-Acl -Path $subfolder.FullName
                        $permission = $permissions[$permissionName]
                        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule (
                            $dlGroupName, 
                            $permission, 
                            "ContainerInherit,ObjectInherit", 
                            "None", 
                            "Allow"
                        )
                        $acl.SetAccessRule($accessRule)
                        Set-Acl -Path $subfolder.FullName -AclObject $acl
                        # Füge die Berechtigungen zur Liste hinzu
                        $newRights += "Recht '$permissionName' auf $($subfolder.FullName) für $dlGroupName gesetzt" 
                    } catch {
                        # Fehlernachricht (schwarz)
                        Add-DebugMessage -message "Fehler beim Setzen der Berechtigung auf $($subfolder.FullName) für '$dlGroupName': $_" -isPositive $false
                    }
                }
            }

            # Zeige die neu angelegten DLs im Debug-Fenster (grün)
            Add-DebugMessage -message "Folgende DLs wurden erfolgreich angelegt:" -isPositive $true
            Add-DebugMessage -message ($newDLs -join [Environment]::NewLine) -isPositive $true

            # Zeige die neu vergebenen Rechte im Debug-Fenster (grün)
            Add-DebugMessage -message "Folgende Rechte wurden erfolgreich vergeben:" -isPositive $true
            Add-DebugMessage -message ($newRights -join [Environment]::NewLine) -isPositive $true

            # Trennstrich
            $separator = '+' * 90  # Zeichen '+' wiederholen
            Add-DebugMessage -message $separator -isSeparator $true

            [System.Windows.Forms.MessageBox]::Show("Vorgang abgeschlossen.")
        })

        $form.Controls.Add($startButton)

        # Funktion zum Hinzufügen von Text in die Debug-TextBox mit Farben
        function Add-DebugMessage {
            param (
                [string]$message,
                [bool]$isPositive,
                [bool]$isSeparator = $false  # Standardwert für isSeparator auf false setzen
            )

            if ($isSeparator) {
                # Setze die Textfarbe auf Schwarz für den Separator
                $debugTextBox.SelectionColor = 'Black'
            } elseif ($isPositive) {
                # Setze die Textfarbe auf Grün für positive Nachrichten
                $debugTextBox.SelectionColor = 'Green'
            } else {
                # Setze die Textfarbe auf Rot für negative Nachrichten
                $debugTextBox.SelectionColor = 'Red'
            }

            # Füge die Nachricht zum Textfeld hinzu
            $debugTextBox.AppendText($message + [Environment]::NewLine)
        }

        # Formular anzeigen
        $form.ShowDialog()
}


# Erstellen des Formulars
$form = New-Object Windows.Forms.Form
$form.Text = 'Administrator Tools'
$form.Size = New-Object Drawing.Size(300, 250)

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

# Erstellen des Buttons für "SMB-RIGHTS"
$btnSMBRIGHTS = New-Object Windows.Forms.Button
$btnSMBRIGHTS.Text = 'SMB-RIGHTS'
$btnSMBRIGHTS.Size = New-Object Drawing.Size(250, 40)
$btnSMBRIGHTS.Location = New-Object Drawing.Point(20, 160)
$btnSMBRIGHTS.Add_Click({ SMB-RIGHTS })

# Hinzufügen der Buttons zum Formular
$form.Controls.Add($btnCreateUser)
$form.Controls.Add($btnCreateGroups)
$form.Controls.Add($btnADGDL)
$form.Controls.Add($btnSMBRIGHTS)

# Anzeigen des Formulars
$form.ShowDialog()

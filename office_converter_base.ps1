Add-Type -AssemblyName 'System.Windows.Forms'

#####################################################
#                                                   #
#    Create the form : button, label and textbox    #
#                                                   #
#####################################################

$form = New-Object System.Windows.Forms.Form
$form.Text = "Scanner d'ancienne version office"
$form.Size = New-Object System.Drawing.Size(600, 400)

$sourceLabel = New-Object System.Windows.Forms.Label
$sourceLabel.Text = 'Dossier de recherche :'
$sourceLabel.Location = New-Object System.Drawing.Point(10, 20)
$sourceLabel.Width = 120
$form.Controls.Add($sourceLabel)

$sourceTextBox = New-Object System.Windows.Forms.TextBox
$sourceTextBox.Location = New-Object System.Drawing.Point(130, 20)
$sourceTextBox.Width = 320
$form.Controls.Add($sourceTextBox)

$sourceButton = New-Object System.Windows.Forms.Button
$sourceButton.Text = 'Parcourir'
$sourceButton.Location = New-Object System.Drawing.Point(460, 20)
$sourceButton.Width = 100
$form.Controls.Add($sourceButton)

# Create the destination folder label and textbox
$destinationLabel = New-Object System.Windows.Forms.Label
$destinationLabel.Text = 'Dossier du rapport :'
$destinationLabel.Location = New-Object System.Drawing.Point(10, 60)
$destinationLabel.Width = 120
$form.Controls.Add($destinationLabel)

$destinationTextBox = New-Object System.Windows.Forms.TextBox
$destinationTextBox.Location = New-Object System.Drawing.Point(130, 60)
$destinationTextBox.Width = 320
$form.Controls.Add($destinationTextBox)

$destinationButton = New-Object System.Windows.Forms.Button
$destinationButton.Text = 'Parcourir'
$destinationButton.Location = New-Object System.Drawing.Point(460, 60)
$destinationButton.Width = 100
$form.Controls.Add($destinationButton)

# Create the output log TextBox (multiline)
$logTextBox = New-Object System.Windows.Forms.TextBox
$logTextBox.Location = New-Object System.Drawing.Point(10, 100)
$logTextBox.Width = 550
$logTextBox.Height = 200
$logTextBox.Multiline = $true
$logTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$logTextBox.ReadOnly = $true
$form.Controls.Add($logTextBox)

# Add the checkbox for conversion mode
$conversionCheckbox = New-Object System.Windows.Forms.CheckBox
$conversionCheckbox.Text = "Convertir les fichiers"
$conversionCheckbox.Location = New-Object System.Drawing.Point(100, 315)
$conversionCheckbox.Width = 140
$form.Controls.Add($conversionCheckbox)

# Add a button to suppress files after converting (user can choose to do it or not)
$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Text = 'Supprimer les fichiers'
$deleteButton.Location = New-Object System.Drawing.Point(400, 315)
$deleteButton.Width = 140
$deleteButton.Enabled = $false
$form.Controls.Add($deleteButton)












# Browse for Source Folder
$sourceButton.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Sélectionnez le dossier de recherche. Ce dossier et tout ses enfants seront scannés à la recherche de document avec les extensions .doc, .xls, .ppt, .vds puis affichés dans le jou`r`nal."
    
    if ($folderDialog.ShowDialog() -eq 'OK') {
        $sourceTextBox.Text = $folderDialog.SelectedPath
        $logTextBox.AppendText("Dossier sélectionné pour la source : $($folderDialog.SelectedPath)`r`n")
    } else {
        $logTextBox.AppendText("Aucun dossier source sélectionné.`r`n")
    }
})

# Browse for Destination Folder
$destinationButton.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Sélectionnez l'emplacement du rapport."

    if ($folderDialog.ShowDialog() -eq 'OK') {
        $destinationTextBox.Text = $folderDialog.SelectedPath
        $logTextBox.AppendText("Dossier sélectionné pour le rapport : $($folderDialog.SelectedPath)`r`n")
    } else {
        $logTextBox.AppendText("Aucun dossier destination sélectionné.`r`n")
    }
})

# When the user clicks on a button to start processing
$processButton = New-Object System.Windows.Forms.Button
$processButton.Text = 'Lancer la recherche'
$processButton.Location = New-Object System.Drawing.Point(250, 315)
$processButton.Width = 120
$form.Controls.Add($processButton)

$processButton.Add_Click({
    $sourcePath = $sourceTextBox.Text
    $destinationPath = $destinationTextBox.Text
    $convertFiles = $conversionCheckbox.Checked  # Check if the conversion checkbox is checked

    # Section to handle error that would prevent search from being performed
    if ([String]::IsNullOrWhiteSpace($sourcePath)) {
        $logTextBox.AppendText("Erreur : Le dossier source est vide. Veuillez sélectionner un chemin valide pour continuer`r`n")
        return
    }

    if ([String]::IsNullOrWhiteSpace($destinationPath)) {
        $logTextBox.AppendText("Erreur : Le dossier de destination est vide. Veuillez sélectionner un chemin valide pour continuer`r`n")
        return
    }

    if (-not (Test-Path $sourcePath)) {
        $logTextBox.AppendText("Erreur : Le dossier source spécifié n'existe pas.`r`n")
        return
    }

    if (-not (Test-Path $destinationPath)) {
        $logTextBox.AppendText("Erreur : Le dossier destination spécifié n'existe pas.`r`n")
        return
    }

    # Define the filename and join it to the path for the csv creation
    $fileName = "old_files_inventory.csv"
    $outputCsv = Join-Path -Path $destinationPath -ChildPath $fileName
    $logTextBox.AppendText("Fichier de sortie : $outputCsv`r`n")

    # Runspace for file processing
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()

    # Create a PowerShell script block to run the file search asynchronously
    $runspaceScriptBlock = {
        param($sourcePath, $destinationPath, $logTextBox, $outputCsv, $convertFiles, $form, $deleteButton)

        # Arrays for file processing
        $docFiles = @()
        $xlsFiles = @()
        $pptFiles = @()
        $totalSize = 0

        # Process files in the source folder
        Get-ChildItem -Path $sourcePath -Recurse -File | Where-Object {
            $_.Extension -match '(\.doc|\.xls|\.ppt)$'
        } | ForEach-Object {
            $fileSize = $_.Length
            $totalSize += $fileSize

            # Update the logTextBox in the UI thread
            $logTextBox.Invoke([Action]{
                $logTextBox.AppendText("Document trouvé dans $($_.FullName)`r`n")
            })

            if ($convertFiles) {
                # Convert files only if $convertFiles is true
                function CleanUp-OfficeProcesses {
                    $officeProcesses = Get-Process | Where-Object { $_.Name -in @("WINWORD", "EXCEL", "POWERPNT") }
                    if ($officeProcesses) {
                        $officeProcesses | ForEach-Object {
                            try {
                                # Forcefully kill Office processes
                                Stop-Process -Name $_.Name -Force
                                Write-Host "Stopped Office process: $($_.Name)"
                            }
                            catch {
                                Write-Host "Error stopping Office process: $($_.Name)"
                            }
                        }
                    } else {
                        Write-Host "No Office processes running."
                    }
                }
            }

            # Start conversion according to the file type
            $tempPath = $_.FullName 
            switch ($_.Extension.ToLower()) {
                '.doc' {
                    if ($convertFiles){
                        # Convert .doc to .docx
                        $newDocxFile = "$($tempPath.Substring(0, $tempPath.Length - 4)).docx"
                        $word = New-Object -ComObject Word.Application
                        $word.Visible = $false
                        $word.DisplayAlerts = $false
                        $document = $word.Documents.Open($tempPath)
                        $document.SaveAs($newDocxFile, 12)  # 12 corresponds to the Word .docx format
                        $document.Close()
                        $word.Quit()
                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null

                        # Update the logTextBox
                        $logTextBox.Invoke([Action]{
                            $logTextBox.AppendText("Conversion terminée : $($newDocxFile)`r`n")
                        })
                    }
                    $docFiles += [PSCustomObject]@{Extension = 'Word (converted)'; FilePath = $tempPath; Size = $fileSize}
                }
                '.xls' {
                    if ($convertFiles){
                        # Convert .xls to .xlsx
                        $newXlsxFile = "$($tempPath.Substring(0, $tempPath.Length - 4)).xlsx"
                        $excel = New-Object -ComObject Excel.Application
                        $excel.Visible = $false
                        $excel.DisplayAlerts = $false
                        $workbook = $excel.Workbooks.Open($tempPath)
                        $workbook.SaveAs($newXlsxFile, 51)  # 51 corresponds to the Excel .xlsx format
                        $workbook.Close()
                        $excel.Quit()
                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null

                        # Update the logTextBox
                        $logTextBox.Invoke([Action]{
                            $logTextBox.AppendText("Conversion terminée : $($newXlsxFile)`r`n")
                        })
                    }
                    $xlsFiles += [PSCustomObject]@{Extension = 'Excel (converted)'; FilePath = $tempPath; Size = $fileSize}
                }
                '.ppt' {
                    if ($convertFiles) {
                        # Convert .ppt to .pptx
                        $newPptxFile = "$($tempPath.Substring(0, $tempPath.Length - 4)).pptx"
    
                        # Check if the file exists before proceeding
                        if (-Not (Test-Path $tempPath)) {
                            $logTextBox.Invoke([Action]{
                                $logTextBox.AppendText("Le fichier source n'existe pas : $tempPath`r`n")
                            })
                        } else {
                            $powerpoint = New-Object -ComObject PowerPoint.Application
                            $powerpoint.Visible = $false        
                            $powerpoint.DisplayAlerts = $false  

                            try {
                                $presentation = $powerpoint.Presentations.Open($tempPath)
                                $presentation.SaveAs($newPptxFile, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsOpenXMLPresentation)
                                $presentation.Close()                 
                            }
                            catch {
                                $logTextBox.Invoke([Action]{
                                    $logTextBox.AppendText("Erreur lors de la conversion de : $tempPath`r`n")
                                })
                            }
                            finally {
                                $presentation = $null
                                $powerpoint.Quit()  # Ensure PowerPoint quits, even if there was an error
                                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint) | Out-Null
                                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
                            }

                            # Update the logTextBox
                            $logTextBox.Invoke([Action]{
                                $logTextBox.AppendText("Conversion terminée : $($newPptxFile)`r`n")
                            })
                        }
                    }

                    $pptFiles += [PSCustomObject]@{Extension = 'PowerPoint (converted)'; FilePath = $tempPath; Size = $fileSize}
                }
            }

            # Clean up any Office process after all (even though there have been errors)
            CleanUp-OfficeProcesses
        }

        # Merge all file arrays
        $fileList = $docFiles + $xlsFiles + $pptFiles

        # Add total size to the file list
        $sizeInGb = [math]::Round($totalSize / 1GB, 2)
        $formattedSize = "$sizeInGb Go"
        $fileList += [PSCustomObject]@{
            Extension = 'Total Size'
            FilePath = ''
            Size = $formattedSize
        }

        # Export the list to CSV
        $fileList | Export-Csv -Path $outputCsv -NoTypeInformation

        # Inform that the file is created
        $logTextBox.Invoke([Action]{
            $logTextBox.AppendText("Le fichier CSV a été créé à l'emplacement suivant : $outputCsv`r`n")
            $logTextBox.AppendText("Le processus de recherche et de traitement des fichiers est terminé.`r`n`r")
            $logTextBox.AppendText("Activation de la fonctionnalité de suppression....`r`n`r")
            
            
        })



        # Logic to handle file deletion
        $deleteButton.Add_Click({
            $fileList | ForEach-Object {
                try {
                    Remove-Item -Path $_.FilePath -Force
                    $logTextBox.AppendText("Fichier supprimé : $($_.FilePath)`r`n")
                } catch {
                    $logTextBox.AppendText("Erreur lors de la suppression du fichier : $($_.FilePath)`r`n")
                }
            }
        })




        # Enable the delete button now that files are found
        $form.Invoke([Action]{
            $deleteButton.Enabled = $true
        })

        


    }

    # Execute the script block on the runspace
    $runspaceThread = [powershell]::Create().AddScript($runspaceScriptBlock).AddArgument($sourcePath).AddArgument($destinationPath).AddArgument($logTextBox).AddArgument($outputCsv).AddArgument($convertFiles).AddArgument($form).AddArgument($deleteButton)
    $runspaceThread.BeginInvoke()





    # Notify the user that the process has started
    $logTextBox.AppendText("`rLe processus de recherche des fichiers a démarré.`r`n")
})






# Show the form
$form.ShowDialog()

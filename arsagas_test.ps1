# Add required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


# Check if any arguments were passed
if ($Args.Count -eq 0) {
    $MessageBoxOptions = [System.Windows.Forms.MessageBoxButtons]::OK
    [System.Windows.Forms.MessageBox]::Show("No files were passed from File Explorer.", "System Exit", $MessageBoxOptions, [System.Windows.Forms.MessageBoxIcon]::Error)
    Exit
} else {
    $FilesToPrint = $Args
}

# Create a form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Enter Number of Copies'
$form.Size = New-Object System.Drawing.Size(400, 280)  # Adjust size for layout
$form.StartPosition = 'CenterScreen'

# Disable resizing
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.MinimizeBox = $false

# Add a PictureBox for the logo
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
$pictureBox.Size = New-Object System.Drawing.Size(160, 90)  # Adjust size of the logo

# Ensure $form.Width and $pictureBox.Width are integers before performing subtraction
$pictureBox.Location = New-Object System.Drawing.Point([int](($form.Width - $pictureBox.Width) / 2), 10)  # Center horizontally
$pictureBox.Image = [System.Drawing.Image]::FromFile("C:\Users\jerem\Documents\Projects\Arsaga's\Arsaga's Printing\arsagas_logo.png")  # Replace with your logo file path
$form.Controls.Add($pictureBox)

# Add a label
$label = New-Object System.Windows.Forms.Label
$label.Text = 'Number of Copies:'
$label.Location = New-Object System.Drawing.Point(50, 100)  # Adjust for centered layout
$label.Size = New-Object System.Drawing.Size(300, 20)
$label.TextAlign = 'MiddleCenter'
$form.Controls.Add($label)

# Add a text box
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point([int](($form.Width - 260) / 2), 130)  # Center horizontally
$textBox.Size = New-Object System.Drawing.Size(260, 20)
$form.Controls.Add($textBox)

# Add an OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = 'OK'
$okButton.Location = New-Object System.Drawing.Point([int](($form.Width - 100) / 2), 170)  # Center horizontally
$okButton.Size = New-Object System.Drawing.Size(100, 30)
$okButton.Add_Click({
    if (-not [int]::TryParse($textBox.Text, [ref]$null) -or [int]$textBox.Text -le 0) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a valid positive number.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } else {
        $form.Tag = $textBox.Text  # Pass the result via the form's Tag property
        $form.Close()
    }
})
$form.Controls.Add($okButton)

# Show the form and wait for input
$form.ShowDialog() | Out-Null
$copies = $form.Tag

# Validate the result
if (-not $copies) {
    Write-Output "No input provided. Exiting."
    Exit
}

$sumatra = "C:\Users\whole\AppData\Local\SumatraPDF\SumatraPDF.exe"

# Iterate over each file and execute the command
foreach ($file in $FilesToPrint) {
    # Validate the file exists
    if (Test-Path $file) {
        $printerName = "Afinia L801 Label Printer - 12oz Label"  # Replace with your printer name
        $printSettings = "${copies}x,fit"  # Replace with your print settings

        # Construct the SumatraPDF command
        $command = "$sumatra -print-to `"$printerName`" -print-settings `"$printSettings`" `"$file`""

        
        # Show the command in a message box
        [System.Windows.Forms.MessageBox]::Show("The command to be executed is: `"$command`"", "Command", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)


        # Run the command
        Write-Output "Running command: $command"
        # Invoke-Expression $command
    } else {
        Write-Output "File not found: $file"
    }
}

# Pause for debugging if run manually
if ($host.Name -eq "ConsoleHost") {
    Start-Sleep -Seconds 5
}
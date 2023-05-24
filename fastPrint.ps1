Add-Type -AssemblyName System.Windows.Forms

$printerName = ""

$form = New-Object System.Windows.Forms.Form
$form.Text = "Account System Properties"
$form.ClientSize = New-Object System.Drawing.Size(440, 260)

$SubmitButton = New-Object System.Windows.Forms.Button
$SubmitButton.Location = New-Object System.Drawing.Point(200, 10)
$SubmitButton.Size = New-Object System.Drawing.Size(200, 20)
$SubmitButton.Text = "Yazdır"

$TargetInput = New-Object System.Windows.Forms.TextBox
$TargetInput.Location = New-Object System.Drawing.Point(10, 10)
$TargetInput.Size = New-Object System.Drawing.Size(200, 20)
$TargetInput.Text = "Klasör uzantısını giriniz"

$flp = New-Object System.Windows.Forms.FlowLayoutPanel
$flp.Location = New-Object System.Drawing.Point(10, 40)
$flp.Height = 200
$flp.Width = 400
$form.Controls.Add($flp)

$printers = Get-Printer

foreach ($printer in $printers) {
    Write-Output "Printer name: $($printer.Name), HostAddress: $($printer.PortName)"
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $printer.Name
    $button.Tag = $printer.PortName
    $button.Add_Click({
        $global:printerName = $printer.Name
        Write-Host $printer.Name
    })
    $flp.Controls.Add($button)
}

$SubmitButton.Add_Click({
    $target = $TargetInput.Text
    $pdfFiles = Get-ChildItem "$target\*.pdf"
    
    foreach ($file in $pdfFiles) {
        $args = @(
            "-sDEVICE=mswinpr2",
            "-dBATCH",
            "-dNOPAUSE",
            "-dNOPROMPT",
            "-dNumCopies=1",
            "-sOutputFile=`"$printerName`"",
            $file.FullName
        )

         $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = "gswin64c.exe"
        $startInfo.Arguments = $args -join " "
        $startInfo.RedirectStandardOutput = $true
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $startInfo
        $process.Start() | Out-Null
        $process.WaitForExit()
    }
})

$form.Controls.Add($TargetInput)
$form.Controls.Add($SubmitButton)
$form.ShowDialog() | Out-Null

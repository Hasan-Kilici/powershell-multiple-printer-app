Add-Type -AssemblyName System.Windows.Forms

$printerName = ""
$MyPrinterIPAddress = ""
$IPPrinterSource = "" 
$IPPrinterPort = ""

$form = New-Object System.Windows.Forms.Form
$form.Text = "Fast Printer"
$form.ClientSize = New-Object System.Drawing.Size(440, 260)

$SubmitButton = New-Object System.Windows.Forms.Button
$SubmitButton.Location = New-Object System.Drawing.Point(200, 10)
$SubmitButton.Size = New-Object System.Drawing.Size(200, 20)
$SubmitButton.Text = "Yazdır"

$SubmitButton2 = New-Object System.Windows.Forms.Button
$SubmitButton2.Location = New-Object System.Drawing.Point(211, 30)
$SubmitButton2.Size = New-Object System.Drawing.Size(189, 20)
$SubmitButton2.Text = "Uzak Yazıcı ile Yazdır"

$TargetInput = New-Object System.Windows.Forms.TextBox
$TargetInput.Location = New-Object System.Drawing.Point(10, 10)
$TargetInput.Size = New-Object System.Drawing.Size(200, 40)
$TargetInput.Text = "Klasör uzantısını giriniz"

$flp = New-Object System.Windows.Forms.FlowLayoutPanel
$flp.Location = New-Object System.Drawing.Point(10, 60)
$flp.Height = 200
$flp.Width = 400
$form.Controls.Add($flp)

$printers = Get-Printer

foreach ($printer in $printers) {
    Write-Output "Printer name: $($printer.Name), HostAddress: $($printer.PortName)"
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $printer.Name
    $button.Tag = $printer.PortName
    $portName = $printer.PortName
    $printerPort = Get-PrinterPort -Name $portName
    $PrinterLIPAdress = (Get-PrinterPort $portName).PrinterHostAddress
    $PrinterIPAddress = $PrinterLIPAddress.Split(':')[0]
    $sourcePrinter = $printerPort.PrinterHostObjectName
    $button.Add_Click({
        $global:PrinterName = $printer.Name
        $global:MyPrinterIPAddress = $PrinterIPAddress.Split('_')[0]
        $global:IPPrinterSource = $sourcePrinter 
        $global:IPPrinterPort = $printerPort
        Write-Host $printer.Name $PrinterName $MyPrinterIPAddress $printerPort $sourcePrinter 
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

$SubmitButton2.Add_Click({
    $target = $TargetInput.Text
    $pdfFiles = Get-ChildItem "$target\*.pdf"

    foreach ($file in $pdfFiles) {
        $printerUNCPath = "\\$MyPrinterIPAddress\$IPPrinterSource"
        $pdfFilePath = $file.FullName

        $destinationPath = Join-Path $printerUNCPath (Split-Path $pdfFilePath -Leaf)
        Copy-Item -Path $pdfFilePath -Destination $destinationPath

        $shell = New-Object -ComObject Shell.Application
        $printerFolder = $shell.Namespace(0x02)
        $printerItem = $printerFolder.ParseName($printerUNCPath)
        $printerItem.InvokeVerb("Print")
    }
})

$form.Controls.Add($TargetInput)
$form.Controls.Add($SubmitButton)
$form.Controls.Add($SubmitButton2)
$form.ShowDialog() | Out-Null

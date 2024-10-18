# Путь к папке с MSI-файлами
$installerPath = "C:\Windows\Installer"

# Имя автора
$authorName = "SolarWinds"
s
# Поиск MSI-файлов в папке
$msiFiles = Get-ChildItem -Path $installerPath -Filter *.msi -Recurse

# Проверка каждого MSI-файла
foreach ($file in $msiFiles) {
    try {
        $msi = New-Object -ComObject WindowsInstaller.Installer
        $db = $msi.OpenDatabase($file.FullName, 0)
        $view = $db.OpenView("SELECT Property, Value FROM Property")
        $view.Execute()
        $properties = @{}
        while ($record = $view.Fetch()) {
            $propertyName = $record.StringData(1)
            $propertyValue = $record.StringData(2)
            $properties[$propertyName] = $propertyValue
        }
        if ($properties["Manufacturer"] -eq $authorName) {
            Write-Output "File: $($file.Name)"
            Start-Process "msiexec.exe" -ArgumentList "/x $($file.FullName)" -Wait
            "burn/purify" | Set-Clipboard
            Start-Process "C:\SolarWinds.Licensing.Reset.exe" -Wait
            Start-Process "C:\SolarWinds-Dameware-MRC-64bit.exe" -Wait
            "Done!"
        }
        $view.Close()
    } catch {
        Write-Output "Error processing file: $($file.Name)"
    }
}
# Cargar ensamblado para utilizar controles de Windows (diálogo y MessageBox)
Add-Type -AssemblyName System.Windows.Forms

# Mostrar diálogo para seleccionar la carpeta
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Seleccione la carpeta donde se realizará la conversión a PDF"
$folderBrowser.ShowNewFolderButton = $false

$result = $folderBrowser.ShowDialog()
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $rootFolder = $folderBrowser.SelectedPath
} else {
    [System.Windows.Forms.MessageBox]::Show("No se seleccionó ninguna carpeta. El script se cerrará.", "Conversión cancelada", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation)
    exit
}

# Obtener todos los archivos de Word y Excel de manera recursiva
$files = Get-ChildItem -Path $rootFolder -Recurse -Include *.doc, *.docx, *.xls, *.xlsx, *.xlsm
$totalFiles = $files.Count

if ($totalFiles -eq 0) {
    [System.Windows.Forms.MessageBox]::Show("No se encontraron archivos de Word o Excel en la carpeta seleccionada.", "Sin archivos", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    exit
}

# Crear instancias COM de Word y Excel
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = 0

$current = 0
foreach ($file in $files) {
    $current++
    $percent = ($current / $totalFiles) * 100
    Write-Progress -Activity "Convirtiendo archivos a PDF" -Status "Procesando $current de ${totalFiles}: $($file.Name)" -PercentComplete $percent

    $extension = $file.Extension.ToLower()
    # La ruta del PDF se crea en la misma ubicación con el mismo nombre base
    $pdfPath = [System.IO.Path]::ChangeExtension($file.FullName, ".pdf")

    if ($extension -eq ".doc" -or $extension -eq ".docx") {
        try {
            Write-Host "Convirtiendo Word: $($file.FullName)"
            $doc = $word.Documents.Open($file.FullName, [Type]::Missing, $true)
            # 17 es el código para formato PDF en Word
            $doc.ExportAsFixedFormat($pdfPath, 17)
            $doc.Close()
        } catch {
            Write-Host "Error al convertir $($file.FullName): $_"
        }
    } elseif ($extension -eq ".xls" -or $extension -eq ".xlsx" -or $extension -eq ".xlsm") {
        try {
            Write-Host "Convirtiendo Excel: $($file.FullName)"
            $workbook = $excel.Workbooks.Open($file.FullName, [Type]::Missing, $true)
            # 0 representa xlTypePDF en Excel
            $workbook.ExportAsFixedFormat(0, $pdfPath)
            $workbook.Close()
        } catch {
            Write-Host "Error al convertir $($file.FullName): $_"
        }
    }
}

# Cerrar las aplicaciones de Office
$word.Quit()
$excel.Quit()

# Mostrar notificación final
[System.Windows.Forms.MessageBox]::Show("Conversión completada. Se han convertido $totalFiles archivos a PDF.", "Proceso Finalizado", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

# Pausa para ver resultados y errores (si se ejecuta desde la consola)
Read-Host -Prompt "Presione Enter para salir"

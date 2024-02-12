####################################################################################################
#
#       Script para Automatizar el proceso de revisión de administradores
#
#                       SE EJECUTA DESDE LA RUTA ORIGINAL DEL FICHERO
#                              
#       
#       @Since: 30-11-2023
#       @Version: 2.5
#       @Author: BAU Evolutio Bankinter 
#
####################################################################################################

## Importamos las librerias necesarias para ejecutar el programa (IMPORTANTE INSTALAR LOS MODULOS)
# Install-module ImportExcel -Repository PSGallery -force
# Install-Module -Name PSExcel -Force -Scope CurrentUser

Import-Module ImportExcel
Import-Module PSExcel

# Variables por defecto
$d = (Get-Date).Day
$m = (Get-Date).Month
$y = (Get-Date).Year

## Funciones del programa

# Funcion para eliminar duplicados de una columna
function Remove-DuplicatesFromSheet {
    param (
        [string]$sheetName,
        [object]$workbook
    )

    # Obtener la worksheet
    $worksheet = $workbook.Sheets.Item($sheetName)

    # Obtener rango columnas
    $columnRange = $worksheet.Range("A2", $worksheet.Cells.Item($worksheet.UsedRange.Rows.Count, 1))

    # Eliminamos duplicados
    $columnRange.RemoveDuplicates(1, $false)

    # Devolvemos workbook
    return $workbook

}

function Convert-CSV-To-XLSX  {
    param (
        [String]$csvFilePath , 
        [String]$sccmFilePath
    )

    ### Cargar el contenido del CSV
    $data = Import-Csv $csvFilePath
    $data | Export-Excel -Path $sccmFilePath -AutoSize

}

function CAUFile {
    param (
        [String]$sccmFilePath
    )

    # Cargamos excel y workbook
    $excel = New-Object -ComObject Excel.Application
    $workbook = $excel.Workbooks.Open($sccmFilePath)

    # Renombrar fila
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.name = "Administradores Locales"

    # Añadir nuevas columnas
    $worksheetAccount = $workbook.Sheets.Add()
    $worksheetAccount.Name = "Account00"
    $worksheetName = $workbook.Sheets.Add()
    $worksheetName.Name = "Name00"

    # Obtener los rangos de las columnas
    $sourceRangeName = $worksheet.Columns.Item("A")
    $sourceRangeAccount = $worksheet.Columns.Item("B")
    $destinationRangeName = $worksheetName.Columns.Item("A")
    $destinationRangeAccount = $worksheetAccount.Columns.Item("A")

    # Copiar los valores de una columna a otra
    $sourceRangeName.Copy($destinationRangeName) | Out-Null
    $sourceRangeAccount.Copy($destinationRangeAccount) | Out-Null

    # Funciones remover duplicados
    $workbook = Remove-DuplicatesFromSheet -sheetName "Account00" -workbook $workbook
    $workbook = Remove-DuplicatesFromSheet -sheetName "Name00" -workbook $workbook

    # Guardar cambios
    $workbook.Save() 
    $workbook.Close()
    $excel.Quit() 

}

function EVOFile {
    param (
        [String]$sccmFilePath,
        [String]$evoFilePath
    )

    # Definimos hoja de trabajo
    $sheetName = "Administradores Locales"

    # Calumna y valor buscado
    $columnName = "Account00"
    $desiredValue = "Admins. Local EB"

    # Importamos excel
    $data = Import-Excel -Path $sccmFilePath -WorksheetName $sheetName

    # Filtro
    $filteredData = $data | Where-Object { $_.$columnName -eq $desiredValue }
    $groupMembers = Get-ADGroupMember -identity "Admins. Local EB" -recursive | Select-Object name

    # Exportamos y añadimos nueva hoja
    $filteredData | Export-Excel -Path $evoFilePath -WorksheetName $sheetName -ClearSheet -AutoSize
    $groupMembers | Export-Excel -Path $evoFilePath -WorksheetName "Admin Local EB" -AutoSize

}

function CompareAdmins {
    param (
        [String]$sccmFilePath,
        [String]$adminControlPath
    )

    # Columna y valor buscado
    $columnNamesccm = "Account00"
    $desiredPattern = "[a-zA-Z]{2}[0-9]{5}"
    $sccmDataExcel = Import-Excel -Path $sccmFilePath -WorksheetName "Account00"

    # Filtramos
    $sccmUsers = $sccmDataExcel | Where-Object { $_.$columnNamesccm -match $desiredPattern } | ForEach-Object { $_.$columnNamesccm }

    # Inventario
    $columnNameInventario = "Codigo Usuario"
    $inventariodataExcel = Import-Excel -Path $adminControlPath -WorksheetName "Admin. Local"
    $inventarioUsers = $inventariodataExcel.$columnNameInventario

    # Admins

    # Encontrar usuarios en SCCM que no están en Inventario
    $newAdmins = $sccmUsers | Where-Object { $_ -notin $inventarioUsers }

    # Encontrar usuarios en Inventario que no están en SCCM
    $oldAdmins = $inventarioUsers | Where-Object { $_ -notin $sccmUsers }

    Write-Host "Usuarios eliminados:"
    Write-Host $oldAdmins
    Write-Host "Usuarios nuevos:"
    Write-Host $newAdmins

    # Añadir el nuevo elemento a la columna
    foreach ($admin in $newAdmins) {
        $inventariodataExcel += [PSCustomObject]@{
            $columnNameInventario = $admin
            "Fecha privilegio" = (Get-Date).Date
            "Estado Permisos" = "ESPERA JUSTIFICACION"
        }
    }

    # Excepction to these users
    $usersException = @("BK05615", "BK09130", "BK70771", "BK71096", "BK71128", "BK71158", "BK71179")

    # Buscar las filas que cumplen con la condición
    $filaEliminar = $inventariodataExcel | Where-Object { 
        $oldAdmins -contains $_."Codigo Usuario" -and
        $_."Codigo Usuario" -notin $usersException
    }

    # Obtener los códigos de usuario que deseas eliminar
    $codigosUsuariosAEliminar = $filaEliminar."Codigo Usuario"

    # Crear una nueva colección sin las filas que deseas eliminar
    $inventariodataExcel = $inventariodataExcel | Where-Object { $_."Codigo Usuario" -notin $codigosUsuariosAEliminar }

    # Sort rows alphabetically based on "Codigo de Usuario" column
    $inventariodataExcel = $inventariodataExcel | Sort-Object -Property "Codigo Usuario"

    # Export the sorted data to EVOExcel
    $inventariodataExcel | Export-Excel -Path $adminControlPath -WorksheetName "Admin. Local" -ClearSheet -AutoSize

}

function AutomateValues {
    param (
        [String]$sccmFilePath,
        [String]$adminControlPath
    )

    # Nombre de la hoja de Excel que contiene los códigos de usuarios
    $sheetNameNew = "Admin. Local"
    $sheetNameSccm = "Administradores Locales"

    # Cargar el contenido del archivo Excel
    $dataExcelSccm = Import-Excel -Path $sccmFilePath -WorksheetName $sheetNameSccm
    $dataExcelNew = Import-Excel -Path $adminControlPath -WorksheetName $sheetNameNew

    $count = 2
    # Iterar a través de las filas en el archivo Excel
    foreach ($row in $dataExcelNew) {
        # Obtener el código de usuario de la fila actual
        $codigoUsuario = $row.'Codigo Usuario'
        #Write-Host $codigoUsuario

        # Utilizar dsquery para obtener información de Active Directory
        $usuarioProperties =  Get-ADUser -Filter {SamAccountName -eq $codigoUsuario} -Properties GivenName, Surname, Department

        $row.'Nombre' = $usuarioProperties.GivenName
        $row.'Apellidos' = $usuarioProperties.Surname
        $row.'Departamento' = $usuarioProperties.Department
        
        if ($null -eq $usuarioProperties){
            # Agregar la información al objeto de la fila actual
            $row.'Nombre' = "NO LOCALIZADO EN AD"
            $row.'Apellidos' = "NO LOCALIZADO EN AD"
            $row.'Departamento' = "NO LOCALIZADO EN AD"
        }

        # Filtramos por codigo
        $filteredData = $dataExcelSccm | Where-Object { $_."Account00" -eq $codigoUsuario }

        # Export the filtered
        $selectedMachine = $filteredData | Select-Object -ExpandProperty 'Name0' 
        $row.'Maquina/Maquinas' = "PENDIENTE"

        if ($selectedMachine){
            # Agregar la información al objeto de la fila actual
            $row.'Maquina/Maquinas' = $selectedMachine -join ', '
        }

        $count++
    }

    # Exportamos de nuevo
    $dataExcelNew | Export-Excel -Path $adminControlPath -WorksheetName $sheetNameNew -ClearSheet -AutoSize


    # Cargar la hoja de cálculo de Excel
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false # Esto evita que Excel se muestre durante el proceso

    # Crear un nuevo libro de trabajo
    $workbook = $excel.Workbooks.Open($adminControlPath)

    # Obtener la hoja de trabajo activa
    $worksheet = $workbook.Worksheets.Item("Admin. Local")

    # Specify the column you want to format (e.g., column 'A')
    $columnToFormatMachines = "G"
    $columnToFormatDate = "D"
    
    # Get the entire column and set the number format to text
    $columnMachines = $worksheet.Columns.Item($columnToFormatMachines)
    $columnMachines.NumberFormat = '@'

    # Get the entire column and set the number format to text
    $columnDate = $worksheet.Columns.Item($columnToFormatDate)
    $columnDate.NumberFormat = "DD/MM/YYYY" 

    # Guardar el libro de trabajo
    $workbook.Save()
    $workbook.Close()
    
    # Cerrar Excel
    $excel.Quit()

}

function DeletedFile{
    param (
        [String]$deletedFilePath,
        [String]$adminControlPath
    )

    # Hoja a trabajar
    $sheetName = "Admin. Local"

    # Columna y valor a filtrar
    $columnName = "Estado Permisos"
    $desiredValue = "NO"

    # Exportamos datos excel
    $data = Import-Excel -Path $adminControlPath -WorksheetName $sheetName

    # Filtramos
    $filteredData = $data | Where-Object { $_.$columnName -eq $desiredValue } | Select-Object 'Maquina/Maquinas', 'Codigo Usuario'

    # Crear array deletes
    $outputData = @()
    $outputData += "Name0;Account00"

    $filteredData | ForEach-Object {
        # Dividir las máquinas por coma y espacio
        $maquinas = $_.'Maquina/Maquinas' -split ', '

        # Si hay mas de una creamos tantas lineas como
        if ($maquinas.Count -gt 1){
            foreach($maq in $maquinas){
                $outputData += "$($maq);$($_.'Codigo Usuario')"
            }

        }else{
            $outputData += "$($_.'Maquina/Maquinas');$($_.'Codigo Usuario')"
        }
    }

    # Export the data to CSV
    $outputData | Out-File -FilePath $deletedFilePath -Encoding UTF8

}
function RiesgosFiles{
    param (
        [String]$riesgosSP,
        [String]$riesgosPT,
        [String]$adminControlPath
    )

    # Specify the sheet name
    $sheetName = "Admin. Local"

    # Specify the column name and the value to match
    $columnEstado = "Estado Permisos"
    $columnCodUser = "Codigo Usuario"
    $desiredValue = "ESPERA JUSTIFICACION"
    $desiredPattern = "[S|s][P|p][0-9]{5}"

    # Load the Excel file
    $data = Import-Excel -Path $adminControlPath -WorksheetName $sheetName
    
    # Filter rows based on the condition
    $usersSP = $data | Where-Object { $_.$columnEstado -eq $desiredValue -and $_.$columnCodUser -notmatch $desiredPattern } | Select-Object 'Codigo Usuario', 'Maquina/Maquinas'
    $usersPT = $data | Where-Object { $_.$columnEstado -eq $desiredValue -and $_.$columnCodUser -match $desiredPattern } | Select-Object 'Codigo Usuario', 'Maquina/Maquinas'

    # Add 5 new empty columns to users
    $usersSP | Add-Member -MemberType NoteProperty -Name "JUSTIFICACION (SI/NO)" -Value $null
    $usersSP | Add-Member -MemberType NoteProperty -Name "DESCRIPCION" -Value $null
    $usersSP | Add-Member -MemberType NoteProperty -Name "TIPO DE EXCEPCION (TEMPORAL / PERMANENTE)" -Value $null
    $usersSP | Add-Member -MemberType NoteProperty -Name "REPETIDO (LISTADO ANTERIORMENTE)" -Value $null
    $usersSP | Add-Member -MemberType NoteProperty -Name "CADUCADO" -Value $null
    $usersPT | Add-Member -MemberType NoteProperty -Name "JUSTIFICACION (SI/NO)" -Value $null
    $usersPT | Add-Member -MemberType NoteProperty -Name "DESCRIPCION" -Value $null
    $usersPT | Add-Member -MemberType NoteProperty -Name "TIPO DE EXCEPCION (TEMPORAL / PERMANENTE)" -Value $null
    $usersPT | Add-Member -MemberType NoteProperty -Name "REPETIDO (LISTADO ANTERIORMENTE)" -Value $null
    $usersPT | Add-Member -MemberType NoteProperty -Name "CADUCADO" -Value $null

    # Export the filtered data back to Excel
    if ($usersSP){
        $usersSP | Export-Excel -Path $riesgosSP -WorksheetName "Nuevos"  -ClearSheet -AutoSize
    }

    if ($usersPT){
        $usersPT | Export-Excel -Path $riesgosPT -WorksheetName "Nuevos"  -ClearSheet -AutoSize
    }     

}


####### MAIN PROGRAM #######

# Extraer usuario que ejecuta el script y añadir la ruta del fichero origen
$csvFilePath = "Administradores Locales " + "$d$m$y" +".csv"
$sccmFilePath = "\\datos02\9860-ext\CIBERSEGURIDAD\Admin_users\2023\SCCM\Administradores Locales " + "$d$m$y" + ".xlsx"
$evoFilePath = "\\datos02\9860-ext\CIBERSEGURIDAD\Admin_users\2023\EVO\Administradores Locales " + "$d$m$y" + " EVO.xlsx"
$adminControlPath = "\\datos02\9860-ext\CIBERSEGURIDAD\Admin_users\2023\Administradores Locales\Monitorizacion Usuarios Administradores Locales.xlsx"
$riesgosSP = "\\datos02\9860-ext\CIBERSEGURIDAD\Admin_users\2023\RIESGOS\Revision Usuarios Administradores " + "$d$m$y" + ".xlsx"
$riesgosPT = "\\datos02\9860-ext\CIBERSEGURIDAD\Admin_users\2023\RIESGOS\Revision Usuarios Administradores PT " + "$d$m$y" + ".xlsx"
$deletedFilePath = "\\datos02\9860-ext\CIBERSEGURIDAD\BAU\BAU 2023\15. Scripts\Revocacion_Usuarios\Entrada.csv"

## De CSV a XLSX
Convert-CSV-To-XLSX -csvFilePath $csvFilePath -sccmFilePath $sccmFilePath

## CAUFile
Write-Host "Creating CAU File"
CAUFile -sccmFilePath $sccmFilePath
Write-Host "Created CAU File"

## EVOFile
Write-Host "Creating EVO File"
EVOFile -sccmFilePath $sccmFilePath -evoFilePath $evoFilePath
Write-Host "Created EVO File"

## Comparar
Write-Host "Comparing Admins..."
CompareAdmins -sccmFilePath $sccmFilePath -adminControlPath $adminControlPath
Write-Host "Comparing Admins Finished"

## Añadir
Write-Host "Appending Info..."
AutomateValues -sccmFilePath $sccmFilePath -adminControlPath $adminControlPath
Write-Host "Finished"

## DeletedFile
Write-Host "Creating Deleted File"
DeletedFile -deletedFilePath $deletedFilePath -adminControlPath $adminControlPath
Write-Host "Created Deleted File"

## RiesgosFiles
Write-Host "Creating Riesgos/Riscos Files"
RiesgosFiles -riesgosSP $riesgosSP -riesgosPT $riesgosPT -adminControlPath $adminControlPath
Write-Host "Created Riesgos/Riscos Files"

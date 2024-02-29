####################################################################################################
#
#                   Script para Automatizar peticiones de IOCs a API de VT
#
#                       SE EJECUTA DESDE LA RUTA ORIGINAL DEL FICHERO
#                              
#       
#       @Since: 22-01-2024
#       @Version: 1.0
#       @Author: Gabriel Medrano 
#
####################################################################################################


# Funcion analizar IOCs en URL
function analyzeURL {
    param (
        [string[]]$malIOCs
    )

    $lastIOCs = @()

    foreach ($ioc in $malIOCs){
        $typeIOC = checkIOC -ioc $ioc
        switch ($typeIOC) {
            1 {
                # Acciones para una IP
                Write-Host "IP in URL -> $ioc"
                # Realiza acciones específicas para IPs
                # Ejecutar el comando y almacenar la salida en una variable
                $maliciousOrSuspicious = vt ip $ioc --include=last_analysis_stats.malicious,last_analysis_stats.suspicious
                
                if ($maliciousOrSuspicious -match "[1-9]") {
                    <# Action to perform if the condition is true #>
                    $lastIOCs += $ioc
                }

                Write-Host $maliciousOrSuspicious
                break
            }
            2 {
                # Acciones para un dominio
                Write-Host "Domain in URL -> $ioc"
                # Realiza acciones específicas para dominios
                $maliciousOrSuspicious = vt domain $ioc --include=last_analysis_stats.malicious,last_analysis_stats.suspicious
                
                if ($maliciousOrSuspicious -match "[1-9]") {
                    <# Action to perform if the condition is true #>
                    $lastIOCs += $ioc
                }

                Write-Host $maliciousOrSuspicious
                break
            }
        }        
    }

    Write-Host "`n"
    Write-Host "IOCs in URL:"
    Write-Host "----------"
    foreach ($item in $lastIOCs) {
        Write-Host $item
    }
    Write-Host "`n"
}


# Función identificar tipo de IOC:
function checkIOC {
    param (
        [string]$ioc
    )

    # Expresión regular verifica URL
    $urlPattern = "\b(https|http):\/\/.+\b"

    # Expresión regular verifica IP
    $ipPattern = "^\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b$"

    if ($ioc -match $urlPattern) {
        return 0  # Es una URL
    }
    elseif ($ioc -match $ipPattern) {
        return 1  # Es una IP
    }
    else {
        return 2  # Es un dominio
    }
}


# Funcion para enviar una solicitud a VirusTotal
function Send-RequestToVirusTotal {
    param (
        [string]$archivo
    )

    # Lista de IOCs maliciosos
    $malIOCs = @()

    # Leer cada línea del archivo
    Get-Content -Path $archivo | ForEach-Object {
        $ioc = $_.Trim()  
        
        $typeIOC = checkIOC -ioc $ioc
        switch ($typeIOC) {
            0 {
                # Acciones para una URL
                Write-Host "URL -> $ioc"
                $uri = [System.Uri]$ioc
                $iocURL = $uri.Host
                $malIOCs += $iocURL
                break
            }
            1 {
                # Acciones para una IP
                Write-Host "IP -> $ioc"
                # Implementar
                break
            }
            2 {
                # Acciones para un dominio
                Write-Host "Domain -> $ioc"
                # Implementar
                break
            }
        }        
    }
    Write-Host "`n"
    Write-Host "Analyze URLs:"
    Write-Host "----------"
    analyzeURL -malIOCs $malIOCs
}



## MAIN ##

# Configuración de la clave de API de VirusTotal y el archivo con IOCs
$apiKey = "" # Indicar API-Key
$archivo = "iocs.txt"

# Inicio API
vt init -k $apiKey

# Uso de la función
Write-Host "`n"
Write-Host "List IOCs:"
Write-Host "----------"
Send-RequestToVirusTotal -archivo $archivo
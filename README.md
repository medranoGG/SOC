<h3 align="center">- BLUE TEAM : Centro de Operaciones Seguridad Proyectos -</h3>

<p align="center">
  <a href="https://skillicons.dev">
    <img src="./icons/Python-Dark.svg" width="48">
    <img src="./icons/api.svg" width="48">
    <img src="./icons/Powershell-Dark.svg" width="48">
  </a>
</p>

<h3>Proyectos:</h3>

- [x] :anger: Script monitorización usuarios administradores locales con PSExcel.
- [x] :anger: Gestión de IOCs con API de VirusTotal.
- [ ] :anger: Procesos APIs varias.


<h3>Ejecutar:</h3>

  - :one: Admins:

```
Install-module ImportExcel -Repository PSGallery -force
Install-Module -Name PSExcel -Force -Scope CurrentUser

.\Automatizacion_Admin.ps1
```

  - :two: VT-API:

```
install api

.\iocs.py
```

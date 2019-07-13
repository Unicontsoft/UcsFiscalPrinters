## UcsFPHub
Unicontsoft Fiscal Printers Hub

### Description

Unicontsoft Fiscal Printers Hub repo builds to a standalone executable that runs on client workstations as a background service and provides access to local fiscal printers.

The access to fiscal printers is provided by parent UcsFP20 component and supports serial port or TCP/IP connectivity to devices. Fiscal printers can be auto-detected on startup by the service too.

You can use the settings file to configure the available endpoints on which the service is accessible e.g. as a JSON based REST service on local TCP/IP port or as a Microsoft SQL Server designated Service Broker queue.

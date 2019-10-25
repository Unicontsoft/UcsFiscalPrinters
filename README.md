[![Build Status](https://dev.azure.com/wqweto0976/UcsFP20/_apis/build/status/wqweto.UcsFiscalPrinters?branchName=master)](https://dev.azure.com/wqweto0976/UcsFP20/_build/latest?definitionId=1&branchName=master)
[![Download stable UcsFPHub-0.1.25.zip](https://img.shields.io/badge/install-UcsFPHub--0.1.25.zip-brightgreen)](https://github.com/wqweto/UcsFiscalPrinters/releases/download/UcsFPHub-0.1.25/UcsFPHub-0.1.25.zip)
[![Download beta UcsFPHub-latest.zip](https://img.shields.io/badge/beta-UcsFPHub--latest.zip-blue)](https://github.com/wqweto/UcsFiscalPrinters/releases/download/UcsFPHub-latest/UcsFPHub-latest.zip)
[![MIT license](https://img.shields.io/:license-mit-blue.svg)](https://github.com/wqweto/UcsFiscalPrinters/blob/master/LICENSE)

## Unicontsoft Fiscal Printers Component 2.0

`UcsFP20` is a COM component that can be used to configure and operate fiscal printers that are popular in Bulgaria.

`UcsFP20` implements lowest-level protocols that are supported by the fiscal printers, usually sending native commands directly to the COM port the device is attached to.

`UcsFP20` support ESP/POS protocol for "kitchen" order lists. ESC/POS protocol supports serial and TCP/IP (LAN) connectivity.

### Available protocols

Here is a complete list of implemented protocols with corresponding tested and supported models:

Name             | Manufacturer | Tested models  | Other supported models
----             | ------------ | -------------  | ----------------------
`TREMOL ECR`     | Tremol Ltd.  | M20            | M23
`DATECS X`       | Datecs Ltd.  | DP-25X         | All X models
`DATECS FP/ECR`  | Datecs Ltd.  | DP-25, FP-650  | All A and B models
`DAISY FP/ECR`   | Daisy Ltd.   | N/A            |
`INCOTEX FP/ECR` | Incotex Ltd. | 777            |
`ESC/POS`        | Tremol Ltd.  | EP-80250       | All ESC/POS "kitchen" printers

### Sub-projects

`UcsFPHub`: [Unicontsoft Fiscal Printers Hub](contrib/UcsFPHub) -- a REST service to provide remote access to locally attached fiscal devices
`UcsFPHub`: [PROTOCOL.md](contrib/UcsFPHub/PROTOCOL.md) -- REST service protocol description
 
### ToDo

  - [ ] Update documentation 

Enjoy!

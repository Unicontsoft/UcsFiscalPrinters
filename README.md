<div align="center">
<img width="64" height="64" src="contrib/UcsFPHub/res/UcsFPHub3.png">

## Unicontsoft Fiscal Printers Component 2.0

[![Build Status](https://dev.azure.com/wqweto0976/UcsFP20/_apis/build/status/wqweto.UcsFiscalPrinters?branchName=master)](https://dev.azure.com/wqweto0976/UcsFP20/_build?definitionId=1)
[![MIT license](https://img.shields.io/:license-mit-blue.svg)](https://github.com/wqweto/UcsFiscalPrinters/blob/master/LICENSE)
</div>

`UcsFP20` is a COM component that can be used to configure and operate fiscal printers that are popular in Bulgaria.

`UcsFP20` implements lowest-level protocols that are supported by the fiscal printers, usually sending native commands directly to the COM port the device is attached to.

`UcsFP20` supports ESP/POS protocol for "kitchen" order lists. ESC/POS protocol supports serial and TCP/IP (LAN) connectivity.

### Available protocols

Here is a complete list of implemented protocols with corresponding tested and supported models:

Protocol         | Manufacturer | Tested models  | Other supported models
----             | ------------ | -------------  | ----------------------
`TREMOL`         | Tremol Ltd.  | M20            | All
`DATECS`         | Datecs Ltd.  | DP-25, FP-650  | A models, B models
`DATECS/X`       | Datecs Ltd.  | DP-25X         | X models
`DAISY`          | Daisy Ltd.   | CompactM       | All
`INCOTEX`        | Incotex Ltd. | 181, 777       | All
`ELTRADE`        | Eltrade Ltd. | A3             | All
`ESC/POS`        | Tremol Ltd.  | EP-80250       | All ESC/POS "kitchen" printers
`PROXY`          | Unicontsoft  | UcsFPHub       | All

### Sub-projects

`UcsFPHub`: [Unicontsoft Fiscal Printers Hub](contrib/UcsFPHub) -- a REST service to provide remote access to locally attached fiscal devices
`UcsFPHub`: [PROTOCOL.md](contrib/UcsFPHub/PROTOCOL.md) -- REST service protocol description
 
### ToDo

  - [ ] Update documentation 

Enjoy!

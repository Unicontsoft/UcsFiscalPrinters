## Unicontsoft Fiscal Printers Component 2.0


`UcsFP20` is a COM component that can be used to configure and operate fiscal printers that are popular in Bulgaria.

`UcsFP20` implements lowest-level protocols that are supported by the fiscal printers, usually sending native commands directly to the COM port the device is attached to.

`UcsFP20` support ESP/POS protocol for "kitchen" order lists and non-fiscal receipts. ESC/POS protocol supports serial and TCP/IP (LAN) connectivity.

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

[`UcsFPHub`](contrib/UcsFPHub): [Unicontsoft Fiscal Printers Hub](contrib/UcsFPHub) -- a REST service to provide remote access to locally attached fiscal devices
 
### ToDo

  - [ ] Update documentation 

Enjoy!

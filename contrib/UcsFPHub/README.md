## UcsFPHub
Unicontsoft Fiscal Printers Hub -- a REST service to provide remote access to locally attached fiscal devices

### Description

Unicontsoft Fiscal Printers Hub repository builds the standalone `UcsFPHub` service executable that can run as a background process or NT service and provide shared access to some or all fiscal devices that are attached to particular client workstation.

The wire protocols implementation is provided by the parent `UcsFP20` component and supports serial COM port connectivity to locally attached devices or TCP/IP (LAN) connectivity to remote devices. Most locally attached fiscal printers can be auto-detected on startup by the `UcsFPHub` service too.

You can use a settings file to allow and configure fiscal printers sharing, including the available endpoints on which `UcsFPHub` service is accessible as a JSON based REST service (local TCP/IP ports) or as a Service Broker queue (through Microsoft SQL Server connection).

### Command-line options

`UcsFPHub.exe` service executable accepts these command-line options:

Option&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; | Long&nbsp;Option&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; | Description
------         | ---------         | ------------
`-c` `FILE`    | `--config` `FILE` | `FILE` is the full pathname to `UcsFPHub` service configuration file. If no explicit configuration options are used the service tries to find `UcsFPHub.conf` configuration file in the application folder. If still no configuration file is found the service auto-detects printers and starts a local REST service listener on `127.0.0.1:8192` by default.
`-i`           | `--install`       | Installs `UcsFPHub` as NT service. Can be used with `-c` to specify custom configuration file to be used by the NT service.
`-u`           | `--uninstall`     | Stops and removes the `UcsFPHub` NT service.
`-s`           | `--systray`       | Hides the process and only shows the application icon in the system notification area.

### Configuration

The service can be configured via an optional `UcsFPHub.conf` file in JSON format. Not using a config file defaults to sharing all auto-detected devices on local COM ports and starting a REST endpoint on 127.0.0.1 port 8192/tcp that is accessible locally only.

Here is a sample settings file:

```json
{
    "Printers": {
        "AutoDetect": true,
        "PrinterID1": {
            "DeviceString": "Protocol=DATECS FP/ECR;Port=COM2;Speed=115200",
            "Description": "Втори етаж, счетоводството"
        }
    },
    "Endpoints": [
        {
            "Binding": "RestHttp", 
            "Address": "192.168.10.11:8192" 
        },
        { 
            "Binding": "MssqlServiceBroker", 
            "ConnectString": "Provider=SQLNCLI11;DataTypeCompatibility=80;MARS Connection=False;Data Source=SQL-PC;Initial Catalog=Dreem15_Personal;User ID=db_user;Password=%_UCS_SQL_PASSWORD%",
            "SshSettings": "Host=ssh.mycompany.com;User ID=ssh_user;Password=%_UCS_SSH_PASSWORD%",
            "QueueName": "POS-PC/12345",
            "QueueTimeout": 5000,
            "SyncDateTimeAdjustTolerance": 120
        },
    ],
    "Environment": {
        "_UCS_FISCAL_PRINTER_LOG": "C:\\Unicontsoft\\POS\\Logs\\UcsFP.log",
        "_UCS_SSH_PASSWORD": "s3cr3t"
    }
}
```

`%VAR_NAME%` placeholders are expanded with values from current process environment. `Printers` object defines available fiscal devices while `Endpoints` array defines where the service will listen for connections from. `Environment` object can be used to setup values in current services environment.

Currently the `UcsFPHub` service supports these environment variables:

Name                            | Description
----                            | -----------
`_UCS_FP_HUB_LOG`               | Set to `c:\path\to\UcsFPHub.log` to log client connections and requests
`_UCS_FISCAL_PRINTER_LOG`       | Set to `c:\path\to\UcsFP.log` for `UcsFP20` component to log communication with fiscal devices
`_UCS_FISCAL_PRINTER_DATA_DUMP` | Set to `1` to include data transfer dump in `_UCS_FISCAL_PRINTER_LOG`
 
### Device string

The device strings are used to configure the connection used for communication with the fiscal device through a list of `Name=Value` pairs separated by `;` delimiter very similar to database connection strings.

Here is a (short) list of supported `Name` entries, all of which are optional unless marked required:

Name             | Type   | Description
----             | ----   | -----------
`Protocol`       | string | (Required) See [**Available protocols**](#available-protocols) below
`Port`           | string | Serial port the device is attached to (e.g. `COM1`)
`Speed`          | number | Controls serial port speed to use (e.g. `9600`)
`Persistent`     | bool   | Controls if serial port is closed after each operation or not (e.g. `Y` or `N`)
`IP`             | address | IP address on which the device is accessible in LAN (e.g. `192.168.10.200`)
`Port`           | number | TCP port to connect to (e.g. `4999`)
`CodePage`       | number | Code page to use when encoding strings to/from the device (e.g. `1251` or `866`)
`RowChars`       | number | Max number of characters on line (depends on the device model and paper loaded)
`ItemChars`      | number | Max number of characters in a product name (defaults to `RowChars - 5`)
`MinDiscount`    | number | e.g. -99%
`MaxDiscount`    | number | e.g. 99%
`MaxReceiptRows` | number | Max number of rows on the receipt supported
`MaxPaymentLen`  | number | Max number of symbols in a payment name
`PingTimeout`    | number | Tremol only
`DetailedReceipt`| bool   | Tremol only

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

### REST service protocol description

All URLs are case-insensitive i.e. `/printers`, `/Printers` and `/PRINTERS` are the same address. Printer IDs are case-insensitive too. Printers are addressed by `:printer_id` which can either be the serial number as reported by the fiscal device or an alias assigned in the service configuration.

See [PROTOCOL.md](PROTOCOL.md) in root on the repo for each REST service endpoint description and sample usage.

### ToDo

 - [x] Listener on Service Broker queues
 - [x] Hook config editor on systray icon popup menu
 - [x] Impl `IniFile` for `MssqlServiceBroker` binding
 - [x] Impl idempotent/cached `POST` requests w/ `request_id=:unique_token` parameter

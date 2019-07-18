## UcsFPHub
Unicontsoft Fiscal Printers Hub

### Description

Unicontsoft Fiscal Printers Hub repo builds to a standalone executable that runs on client workstations as a background service and provides access to local fiscal printers.

The access to fiscal printers is provided by parent UcsFP20 component and supports serial port or TCP/IP connectivity to devices. Fiscal printers can be auto-detected on startup by the service too.

You can use the settings file to configure the available endpoints on which the service is accessible e.g. as a JSON based REST service on local TCP/IP port or as a Microsoft SQL Server designated Service Broker queue.

### Configuration

The service is configures through `UcsFPHub.conf` file in JSON format. Here is a sample config

```json
{
    "Printers": {
        "Autodetect": true,
        "PrinterID1": {
            "DeviceString": "Protocol=DATECS FP/ECR;Port=COM2;Speed=115200"
        }
    },
    "Endpoints": [
        { 
            "Binding": "MssqlServiceBroker", 
            "ConnectString": "Provider=SQLNCLI10;DataTypeCompatibility=80;MARS Connection=False;Data Source=SQL-PC;Initial Catalog=Dreem15_Personal;User ID=db_user;Password=%_UCS_SQL_PASSWORD%",
            "SshSettings": "1,SSH-PC,22,ssh_user,%_UCS_SSH_PASSWORD%",
            "IniFile": "C:\\Unicontsoft\\Pos\\Pos.ini"
        },
        {
            "Binding": "RestHttp", 
            "Address": "127.0.0.1:8192" 
        }
    ],
    "Environment": {
        "_UCS_FISCAL_PRINTER_LOG": "C:\\Unicontsoft\\POS\\Logs\\UcsFP.log"
    }
}
```

`%VAR_NAME%` placeholders are expanded with values from current service environment. `Printers` defines available fiscal devices while `Endpoints` defines where the service will expect connections from. `Environment` can be used to setup values in current services environment.

Currently the service supports these environment variables:

  - `_UCS_FISCAL_PRINTER_LOG` to specify log file for `UcsFP20.dll` to log communication with fiscal devices
  - `_UCS_FISCAL_PRINTER_DATA_DUMP` when set dumps data transfer too
  - `_UCS_FP_HUB_LOG` to specify client connections log

### Command-line options

| Option&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;   | Long&nbsp;Option&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; | Description                                             |
| -------------- | ----------------- | ------------------------------------------------------- |
| `-c` `FILE`    | `--config` `FILE` | Where `FILE` is full pathname to service config file. If not used the service looks for optional `UcsFPHub.conf` in application folder. If no config file is found the service auto-detects printers a starts local REST service listener on `127.0.0.1:8192` by default. |
| `-i`           | `--install`       | Installs `UcsFPHub` NT service. Can be used with `-c` to specify custom config file to be used by the NT service. |
| `-u`           | `--uninstall`     | Stops and removes `UcsFPHub` NT service.                   |
                                                                |

### REST service protocol description

All URLs are case-insensitive i.e. `/printers`, `/Printers` and `/PRINTERS` are the same address. Printer IDs are case-insensitive too. You can address printers by serial number or by ID (alias) in config file.

These are the REST service endpoints supported. 

#### `GET` `/printers`

Lists currently configured printers.

```
C:> curl -s http://localhost:8192/printers | jq
```
```json
{
  "Count": 2,
  "DT240349": {
    "DeviceSerialNo": "DT240349",
    "FiscalMemoryNo": "02240349",
    "DeviceProtocol": "DATECS FP/ECR",
    "DeviceModel": "FP-3530?",
    "FirmwareVersion": "4.10BG 10MAR08 1130",
    "CharsPerLine": 30,
    "TaxNo": "0000000000",
    "TaxCaption": "БУЛСТАТ",
    "DeviceString": "Protocol=DATECS FP/ECR;Port=COM1;Speed=9600"
  },
  "DT518315": {
    "DeviceSerialNo": "DT518315",
    "FiscalMemoryNo": "02518315",
    "DeviceProtocol": "DATECS FP/ECR",
    "DeviceModel": "DP-25",
    "FirmwareVersion": "263453 08Nov18 1312",
    "CharsPerLine": 30,
    "TaxNo": "НЕЗАДАДЕН",
    "TaxCaption": "ЕИК",
    "DeviceString": "Protocol=DATECS FP/ECR;Port=COM2;Speed=115200"
  },
  "Aliases": {
    "Count": 1,
    "PrinterID1": {
      "DeviceSerialNo": "DT518315"
    }
  }
}
```

#### `GET` `/printers/:printer_id`

Retrieves single printer configuration

```
C:> curl -s http://localhost:8192/printers/DT518315 | jq
```
```json
{
  "DeviceSerialNo": "DT518315",
  "FiscalMemoryNo": "02518315",
  "DeviceProtocol": "DATECS FP/ECR",
  "DeviceModel": "DP-25",
  "FirmwareVersion": "263453 08Nov18 1312",
  "CharsPerLine": 30,
  "TaxNo": "НЕЗАДАДЕН",
  "TaxCaption": "ЕИК",
  "DeviceString": "Protocol=DATECS FP/ECR;Port=COM2;Speed=115200"
}
```

#### `GET` `/printers/:printer_id/status`

```
C:> curl -s http://localhost:8192/printers/DT518315/status | jq
```
```json
{
  "Ok": true,
  "DeviceStatus": "",
  "DeviceDateTime": "2018-07-19 22:55:53"
}
```

#### `POST` `/printers/:printer_id/receipt`

Print fiscal receipt, reversal receipt, invoice or credit note.

```
C:> curl -s http://localhost:8192/printers/DT518315/receipt -d "{ }" | jq
```
```json
{
  "Ok": true,
  "ReceiptNo": "0000048",
  "ReceiptDateTime": "2018-07-19 22:59:34",
  "DeviceSerialNo": "DT518315",
  "FiscalMemoryNo": "02518315"
}
```

#### `POST` `/printers/:printer_id/deposit`

Print service deposit or withdraw

```
C:> curl -s http://localhost:8192/printers/DT518315/deposit -d "{ \"Amount\": 12.34 }" | jq
```
```json
{
  "Ok": true,
  "ReceiptNo": "0000048",
  "ReceiptDateTime": "2018-07-19 23:02:23",
  "Available": 337.34,
  "TotalDeposits": 369,
  "TotalWithdraws": 123
}
```

#### `POST` `/printers/:printer_id/report`

Print reports. Supports daily X or Z and monthly report (by date range).

```
C:> curl -s http://localhost:8192/printers/DT518315/report -d "{ \"ReportType\": 1 }" | jq
```
```json
{
  ...
}
```
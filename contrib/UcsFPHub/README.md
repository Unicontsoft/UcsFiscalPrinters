## UcsFPHub
Unicontsoft Fiscal Printers Hub -- a REST service to provide remote access to locally attached fiscal devices

### Description

Unicontsoft Fiscal Printers Hub repo builds a standalone `UcsFPHub` service executable that runs in the background (or as an NT service) to provide remote access to client-workstation attached fiscal devices.

The access to fiscal devices is provided by parent `UcsFP20` component and supports serial port or TCP/IP connectivity to attached devices. Most fiscal printers can be auto-detected on startup by the `UcsFPHub` service too.

You can use a settings file to configure the available endpoints on which the service is accessible e.g. as a JSON based REST service on local TCP/IP port or as a Microsoft SQL Server designated Service Broker queue.

### Configuration

The service is configured through `UcsFPHub.conf` file in JSON format. Here is a sample configuration file:

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

`%VAR_NAME%` placeholders are expanded with values from current process environment. `Printers` object defines available fiscal devices while `Endpoints` array defines where the service will listen for connections from. `Environment` object can be used to setup values in current services environment.

Currently the `UcsFPHub` service supports these environment variables:

 - `_UCS_FISCAL_PRINTER_LOG` to specify `c:\path\to\UcsFP.log` log file for `UcsFP20` component to log communication with fiscal devices
 - `_UCS_FISCAL_PRINTER_DATA_DUMP` set to `1` to dump data transfer too
 - `_UCS_FP_HUB_LOG` to specify client connections `c:\path\to\UcsFPHub.log` log file

### Command-line options

`UcsFPHub.exe` service executable accepts these command-line options:

| Option&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;   | Long&nbsp;Option&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; | Description                                             |
| -------------- | ----------------- | ------------------------------------------------------- |
| `-c` `FILE`    | `--config` `FILE` | `FILE` is the full pathname to `UcsFPHub` service config file. If no explicit config options are used the service tries to find `UcsFPHub.conf` config file in the application folder. If still no config file is found the service auto-detects printers and starts a local REST service listener on `127.0.0.1:8192` by default. |
| `-i`           | `--install`       | Installs `UcsFPHub` as NT service. Can be used with `-c` to specify custom config file to be used by the NT service. |
| `-u`           | `--uninstall`     | Stops and removes the `UcsFPHub` NT service.                   |

### REST service protocol description

All URLs are case-insensitive i.e. `/printers`, `/Printers` and `/PRINTERS` are the same address. Printer IDs are case-insensitive too. You can address printers by serial number or by ID (alias) in config file.

All responses are in minimized JSON so `curl` sample requests below can use [`jq`](https://stedolan.github.io/jq/) to format JSON results.

These are the REST service endpoints supported:

#### `GET` `/printers`

List currently configured devices.

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

Retrieve device configuration, header texts, footer texts, tax number/caption, last receipt number/datetime and payment names.

```
C:> curl -s http://localhost:8192/printers/DT518315 | jq
```
```json
{
  "Ok": true,
  "DeviceSerialNo": "DT518315",
  "FiscalMemoryNo": "02518315",
  "DeviceProtocol": "DATECS FP/ECR",
  "DeviceModel": "DP-25",
  "FirmwareVersion": "263453 08Nov18 1312",
  "CharsPerLine": 30,
  "Header": [
    "               ИМЕ НА ФИРМА",
    "              АДРЕС НА ФИРМА",
    "               ИМЕ НА ОБЕКТ",
    "              АДРЕС НА ОБЕКТ",
    "",
    ""
  ],
  "Footer": [
    "",
    ""
  ],
  "TaxNo": "НЕЗАДАДЕН",
  "TaxCaption": "ЕИК",
  "ReceiptNo": "0000048",
  "DeviceDateTime": "2019-07-19 11:51:33",
  "PaymentName": [
    "В БРОЙ",
    "С ДЕБИТНА КАРТА",
    "С ЧЕК",
    "ВАУЧЕР",
    "КУПОН",
    "",
    ""
  ]
}
```

#### `POST` `/printers/:printer_id`

Retrieve device configuration only. This will not communicate with the device if config is already retrieved on previous connection.
```
C:> curl -s http://localhost:8192/printers/DT518315 -d "{ }"  | jq
```
```json
{
  "Ok": true,
  "DeviceSerialNo": "DT518315",
  "FiscalMemoryNo": "02518315",
  "DeviceProtocol": "DATECS FP/ECR",
  "DeviceModel": "DP-25",
  "FirmwareVersion": "263453 08Nov18 1312",
  "CharsPerLine": 30
}
```

Retrieve device configuration, operator name and default password.
```
C:> curl -s http://localhost:8192/printers/DT518315 -d "{ \"Operator\": { \"Code\": 1 } }"  | jq
```
```json
{
  "Ok": true,
  "DeviceSerialNo": "DT518315",
  "FiscalMemoryNo": "02518315",
  "DeviceProtocol": "DATECS FP/ECR",
  "DeviceModel": "DP-25",
  "FirmwareVersion": "263453 08Nov18 1312",
  "CharsPerLine": 30,
  "Operator": {
    "Code": 1,
    "Name": "Оператор 1",
    "Password": "****"
  }
}
```

Retrieve device configuration and tax number/caption only
```
C:> curl -s http://localhost:8192/printers/DT518315 -d "{ \"IncludeTaxNo\": true }"  | jq
```
```json
{
  "Ok": true,
  "DeviceSerialNo": "DT518315",
  "FiscalMemoryNo": "02518315",
  "DeviceProtocol": "DATECS FP/ECR",
  "DeviceModel": "DP-25",
  "FirmwareVersion": "263453 08Nov18 1312",
  "CharsPerLine": 30,
  "TaxNo": "НЕЗАДАДЕН",
  "TaxCaption": "ЕИК"
}
```

#### `GET` `/printers/:printer_id/status`

Get device status and current clock.

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

Print fiscal receipt, reversal, invoice or credit note.

```
C:> curl -s http://localhost:8192/printers/DT518315/receipt -d ^"{  ^
    \"ReceiptType\": 1 ^
}^" | jq
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

Supported `ReceiptType` values:

| Name                  | Value | Description                                             |
| --------------        | ----- | ------------------------------------------------------- |
| `ucsFscRcpSale`       | 1     | Prints fiscal receipt |
| `ucsFscRcpReversal`   | 2     | Prints reversal receipt |
| `ucsFscRcpInvoice`    | 3     | Prints extended fiscal receipt |
| `ucsFscRcpCreditNote` | 4     | Prints extended reversal receipt  |
| `ucsFscRcpOrderList`  | 5     | Prints kitchen printers order-list |


#### `GET` `/printers/:printer_id/deposit`

Retrieve service deposit or service withdraw totals.

```
C:> curl -s http://localhost:8192/printers/DT518315/deposit  | jq
```
```json
{
  "Ok": true,
  "Available": 349.68,
  "TotalDeposits": 381.34,
  "TotalWithdraws": 123
}
```

#### `POST` `/printers/:printer_id/deposit`

Print service deposit.

```
C:> curl -s http://localhost:8192/printers/DT518315/deposit -d "{ \"Amount\": 12.34 }" | jq
```
```json
{
  "Ok": true,
  "ReceiptNo": "0000050",
  "ReceiptDateTime": "2019-07-19 12:02:08",
  "Available": 362.02,
  "TotalDeposits": 393.68,
  "TotalWithdraws": 123
}
```

Print service withdraw.

```
C:> curl -s http://localhost:8192/printers/DT518315/deposit -d ^"{ ^
    \"Amount\": -56.78, ^
    \"Operator\": { ^
        \"Code\": \"2\", ^
        \"Password\": \"****\" ^
    } ^
}^" | jq
```
```json
{
  "Ok": true,
  "ReceiptNo": "0000052",
  "ReceiptDateTime": "2019-07-19 12:03:41",
  "Available": 248.46,
  "TotalDeposits": 393.68,
  "TotalWithdraws": 236.56
}
```

#### `POST` `/printers/:printer_id/report`

Print device reports. Supports daily X or Z reports and monthly (by date range) reports.

```
C:> curl -s http://localhost:8192/printers/DT518315/report -d "{ \"ReportType\": 1 }" | jq
```
```json
{
  "Ok": true,
  "ReceiptNo": "",
  "ReceiptDateTime": "00:00:00"
}
```

Supported `ReportType` values:

| Name                | Value | Description                                             |
| --------------      | ----- | ------------------------------------------------------- |
| `ucsFscRptDaily`    | 1     | Prints daily X or Z report. Set `IsClear` for Z report, `IsItems` for report by items, `IsDepartments` for daily report by departments.  |
| `ucsFscRptNumber`   | 2     | Not implemented  |
| `ucsFscRptDate`     | 3     | Prints monthly fiscal report. Use `FromDate` and `ToDate` to specify date range.  |
| `ucsFscRptOperator` | 4     | Not implemented  |


#### `GET` `/printers/:printer_id/datetime`

Get current device date/time.

```
C:> curl -s http://localhost:8192/printers/DT518315/datetime | jq
```
```json
{
  "Ok": true,
  "DeviceStatus": "",
  "DeviceDateTime": "2019-07-19 11:58:31"
}
```

#### `POST` `/printers/:printer_id/datetime`

Set device date/time.

```
C:> curl -s http://localhost:8192/printers/DT518315/datetime -d "{ \"DeviceDateTime\": \"2019-07-19 11:58:31\" }" | jq
```
```json
{
  "Ok": true,
  "PreviousDateTime": "2019-07-19 11:58:39",
  "DeviceStatus": "",
  "DeviceDateTime": "2019-07-19 11:58:31"
}
```

Set device date/time only when device clock is outside specified tolerance (in seconds).

```
C:> curl -s http://localhost:8192/printers/DT518315/datetime -d ^"{ ^
    \"DeviceDateTime\": \"2019-07-19 11:58:31\", ^
    \"AdjustTolerance\": 60 ^
}^" | jq
```
```json
{
  "Ok": true,
  "PreviousDateTime": "2019-07-19 11:59:43",
  "DeviceStatus": "",
  "DeviceDateTime": "2019-07-19 11:58:31"
}
```

### ToDo

 - [ ] Listener on Service Broker queues
    
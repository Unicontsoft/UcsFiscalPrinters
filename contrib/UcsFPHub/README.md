## UcsFPHub
Unicontsoft Fiscal Printers Hub -- a REST service to provide remote access to locally attached fiscal devices

### Description

Unicontsoft Fiscal Printers Hub repository builds the standalone `UcsFPHub` service executable that can run as a background process or NT service and provide shared access to some or all fiscal devices that are attached to particular client workstation.

The wire protocols implementation is provided by the parent `UcsFP20` component and supports serial COM port connectivity to locally attached devices or TCP/IP (LAN) connectivity to remote devices. Most locally attached fiscal printers can be auto-detected on startup by the `UcsFPHub` service too.

You can use a settings file to allow and configure fiscal printers sharing, including the available endpoints on which `UcsFPHub` service is accessible as a JSON based REST service (local TCP/IP ports) or as a Service Broker queue (through Microsoft SQL Server connection).

### Configuration

The service is configured by a `UcsFPHub.conf` file in JSON format. Here is a sample settings file:

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

Option&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; | Long&nbsp;Option&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; | Description
------         | ---------         | ------------
`-c` `FILE`    | `--config` `FILE` | `FILE` is the full pathname to `UcsFPHub` service config file. If no explicit config options are used the service tries to find `UcsFPHub.conf` config file in the application folder. If still no config file is found the service auto-detects printers and starts a local REST service listener on `127.0.0.1:8192` by default.
`-i`           | `--install`       | Installs `UcsFPHub` as NT service. Can be used with `-c` to specify custom config file to be used by the NT service.
`-u`           | `--uninstall`     | Stops and removes the `UcsFPHub` NT service.

### ToDo

 - [ ] Listener on Service Broker queues
    
## REST service protocol description

All URLs are case-insensitive i.e. `/printers`, `/Printers` and `/PRINTERS` are the same address. Printer IDs are case-insensitive too. Printers are addressed by `:printer_id` which can either be the serial number as reported by the fiscal device or an alias assigned in the service configuration.

Both request and response payloads are of `application/json; charset=utf-8` type whether explicitly requested in `Accept` and `Content-Type` headers or not. All endpoints return `"Ok": true` on success or in case of failure include `"ErrorText": "Описание на грешка"` in the JSON response.

The `UcsFPHub` service endpoints return minimized JSON so sample `curl` requests below use [`jq`](https://stedolan.github.io/jq/) (a.k.a. **J**SON **Q**uery) utility to format response in human readable JSON.

These are the REST service endpoints supported:

#### `GET` `/printers`

List currently configured devices.

```
C:> curl http://localhost:8192/printers -sS | jq
```
```json
{
  "Ok": true,
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
C:> curl http://localhost:8192/printers/DT518315 -sS | jq
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
C:> curl http://localhost:8192/printers/DT518315 -d "{ }" -sS | jq
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
C:> curl http://localhost:8192/printers/DT518315 -d "{ \"Operator\": { \"Code\": 1 } }" -sS | jq
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
C:> curl http://localhost:8192/printers/DT518315 -d "{ \"IncludeTaxNo\": true }" -sS | jq
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
C:> curl http://localhost:8192/printers/DT518315/status -sS | jq
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

Following `data-utf8.txt` prints a fiscal receipt (`ReceiptType` is 1, see below) for two products, second one is with discount. The receipt in paid first 10.00 leva with a bank card and the rest in cash. After all receipt totals are printed a free-text line outputs current client loyalty card number used for information.

```json
{
    "ReceiptType": 1,
    "Operator": {
        "Code": "1",
        "Name": "Иван Иванов",
        "Password": "****"
    },
    "UniqueSaleNo": "DT518315-0001-1234567",
    "Rows": [
        {
            "ItemName": "Продукт 1",
            "Price": 12.34,
        },
        {
            "ItemName": "Продукт 2",
            "Price": 5.67,
            "TaxGroup": 2,
            "Quantity": 3.5,
            "Discount": 15
        },
        {
            "Amount": 10,
            "PaymentType": 2
        },
        {
            "PaymentType": 1
        },
        {
            "Text": "Клиентска карта: 12345"
        }
    ]
}
```
```
C:> curl http://localhost:8192/printers/DT518315/receipt --data-binary @data-utf8.txt -sS | jq
```
```json
{
  "Ok": true,
  "ReceiptNo": "0000056",
  "ReceiptDateTime": "2019-07-19 14:05:18",
  "DeviceSerialNo": "DT518315",
  "FiscalMemoryNo": "02518315"
}
```

Following `data-utf8.txt` prints a reversal receipt (`ReceiptType` is 2, see below) for the first products of the previous sale `0000056`.

```
{
    "ReceiptType": 2,
    "Operator": {
        "Code": "1",
        "Name": "Иван Иванов",
        "Password": "****"
    },
    "Reversal: {
        "Type": 1,
        "ReceiptNo": "0000056",
        "ReceiptDateTime": "2019-07-19 14:05:18",
        "FiscalMemoryNo": "02518315",
    },
    "UniqueSaleNo": "DT518315-0001-1234567",
    "Rows": [
        {
            "ItemName": "Продукт 1",
            "Price": 12.34,
        }
    ]
}
```
```
C:> curl http://localhost:8192/printers/DT518315/receipt --data-binary @data-utf8.txt -sS | jq
```
```json
{
  "Ok": true,
  "ReceiptNo": "...",
  "ReceiptDateTime": "...",
  "DeviceSerialNo": "DT518315",
  "FiscalMemoryNo": "02518315"
}
```

Duplicate last receipt. Can be executed only once immediately after printing a receipt (or the command fails).

```
C:> curl http://localhost:8192/printers/DT518315/receipt -d "{ \"PrintDuplicate\": true }" -sS | jq
```
```json
{
  "Ok": false,
  "ErrorText": "Непозволена команда"
}
```

Print duplicate receipt (by receipt number) from device Electronic Journal.

```
C:> curl http://localhost:8192/printers/DT518315/receipt -d "{ \"PrintDuplicate\": true, \"Invoice\": { \"DocNo\": 57 } }" -sS | jq
```
```json
{
  "Ok": false,
  "ErrorText": "Непозволена команда"
}
```

Supported `ReceiptType` values:

Name                  | Value   | Description
----                  | -----   | -----------
`ucsFscRcpSale`       | 1       | Prints fiscal receipt
`ucsFscRcpReversal`   | 2       | Prints reversal receipt
`ucsFscRcpInvoice`    | 3       | Prints extended fiscal receipt
`ucsFscRcpCreditNote` | 4       | Prints extended reversal receipt
`ucsFscRcpOrderList`  | 5       | Prints kitchen printers order-list

Supported `PaymentType` values:

Name                  | Value | Alt  | Description
----                  | ----- | ---- | -----------
`ucsFscPmtCash`       | 1     |      | Payment in cash
`ucsFscPmtCard`       | 2     |      | Payment with debit/credit card
`ucsFscPmtCheque`     | 3     |      | Bank payment (if available)
`ucsFscPmtCustom1`    | -1    | 5    | First custom payment (Талони)
`ucsFscPmtCustom2`    | -2    | 6    | Second custom payment (В.Талони)
`ucsFscPmtCustom3`    | -3    | 7    | Third custom payment (Резерв.1)
`ucsFscPmtCustom4`    | -4    | 8    | Fourth custom payment (Резерв.2)

Supported `ReversalType` values:

Name                        | Value | Description
----                        | ----- | -----------
`ucsFscRevOperatorError`    | 0     | Error by the operator (default)
`ucsFscRevRefund`           | 1     | Refund defect/returned items
`ucsFscRevTaxBaseReduction` | 2     | Reduction of price/quantity of items in an invoice. Use for credit notes only

#### `GET` `/printers/:printer_id/deposit`

Retrieve service deposit and service withdraw totals.

```
C:> curl http://localhost:8192/printers/DT518315/deposit -sS | jq
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
C:> curl http://localhost:8192/printers/DT518315/deposit -d "{ \"Amount\": 12.34 }" -sS | jq
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
C:> curl http://localhost:8192/printers/DT518315/deposit -d ^"{ ^
    \"Amount\": -56.78, ^
    \"Operator\": { ^
        \"Code\": \"2\", ^
        \"Password\": \"****\" ^
    } ^
}^" -sS | jq
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
C:> curl http://localhost:8192/printers/DT518315/report -d "{ \"ReportType\": 1 }" -sS | jq
```
```json
{
  "Ok": true,
  "ReceiptNo": "",
  "ReceiptDateTime": "00:00:00"
}
```

Supported `ReportType` values:

Name                | Value | Description
----                | ----- | -----------
`ucsFscRptDaily`    | 1     | Prints daily X or Z report. Set `IsClear` for Z report, `IsItems` for report by items, `IsDepartments` for daily report by departments.
`ucsFscRptNumber`   | 2     | Not implemented
`ucsFscRptDate`     | 3     | Prints monthly fiscal report. Use `FromDate` and `ToDate` to specify date range.
`ucsFscRptOperator` | 4     | Not implemented


#### `GET` `/printers/:printer_id/datetime`

Get current device date/time.

```
C:> curl http://localhost:8192/printers/DT518315/datetime -sS | jq
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
C:> curl http://localhost:8192/printers/DT518315/datetime -d "{ \"DeviceDateTime\": \"2019-07-19 11:58:31\" }" -sS | jq
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
C:> curl http://localhost:8192/printers/DT518315/datetime -d ^"{ ^
    \"DeviceDateTime\": \"2019-07-19 11:58:31\", ^
    \"AdjustTolerance\": 60 ^
}^" -sS | jq
```
```json
{
  "Ok": true,
  "PreviousDateTime": "2019-07-19 11:59:43",
  "DeviceStatus": "",
  "DeviceDateTime": "2019-07-19 11:58:31"
}
```

#### `GET` `/printers/:printer_id/totals`

Get device totals since last Z report grouped by payment types and tax groups.

```
C:> curl http://localhost:8192/printers/DT518315/totals -sS | jq
```
```json
{
    "Ok": true,
    "NumReceipts": 24,
    "TotalsByPayments": [
        { "PaymentType": 1, "PaymentName": "В БРОЙ", "Amount": 178.18 },
        { "PaymentType": 2, "PaymentName": "С КАРТА", "Amount": 52.96 },
        { "PaymentType": 3, "PaymentName": "НЗОК", "Amount": 0 },
        { "PaymentType": 4, "PaymentName": "КРЕДИТ", "Amount": 0 },
        { "PaymentType": 5, "PaymentName": "ВАУЧЕР", "Amount": 0 },
        { "PaymentType": 6, "PaymentName": "КУПОН", "Amount": 0 },
        { "PaymentType": 7, "PaymentName": "", "Amount": 0 },
        { "PaymentType": 8, "PaymentName": "", "Amount": 0 }
    ],
    "TotalsByTaxGroups": [
        { "TaxGroup": 1, "VatPercent": 0, "Amount": 231.14 },
        { "TaxGroup": 2, "VatPercent": 20, "Amount": 0 },
        { "TaxGroup": 3, "VatPercent": 20, "Amount": 0 },
        { "TaxGroup": 4, "VatPercent": 9, "Amount": 0 },
        { "TaxGroup": 5, "VatPercent": 0, "Amount": 0 },
        { "TaxGroup": 6, "VatPercent": 0, "Amount": 0 },
        { "TaxGroup": 7, "VatPercent": 0, "Amount": 0 },
        { "TaxGroup": 8, "VatPercent": 0, "Amount": 0 }
    ]
}
```

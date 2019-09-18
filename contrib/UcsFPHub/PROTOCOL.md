## REST service protocol description

All URLs are case-insensitive i.e. `/printers`, `/Printers` and `/PRINTERS` are the same address. Printer IDs are case-insensitive too. Printers are addressed by `:printer_id` which can either be the serial number as reported by the fiscal device or an alias assigned in the service configuration.

Both request and response payloads by default are of `application/json; charset=utf-8` type whether explicitly requested in `Accept` and `Content-Type` headers or not. Use `format=xml` as URL query string parameter to change results format to XML with content-type of `text/xml; charset=utf-8`.

Use `request_id=N2qbikc5lUU` as URL query string parameter for idempotent `POST` requests (and general cache control) in order to prevent duplicating fiscal transactions when repeating requests because of a timeout or connectivity issues. If a request is repeated with the same payload and `request_id` then the results would be fetched directly from service cache without communicating with the fiscal device, provided that the previous execution of the same `request_id` succeeded.

All endpoints return `"Ok": true` on success and in case of failure include `"ErrorText": "Описание на грешка"` localized error text in the response.

All endpoints support includes middleware. This allows for instance to set `"IncludePaymentNames": true` in `POST /printers/:printer_id/deposit` request to return available payment names along with standard results. Supported includes are `IncludeHeaders`, `IncludeFooters`, `IncludeTaxNo`, `IncludeReceiptNo`, `IncludePaymentNames` and `IncludeAll` which activates all previous includes.

The `UcsFPHub` service endpoints return minimized JSON so sample `curl` requests below use [`jq`](https://stedolan.github.io/jq/) (a.k.a. **J**SON **Q**uery) utility to format response in human readable JSON.

These are the REST service endpoints supported:

#### `GET` `/printers`

List currently configured devices.

```shell
C:> curl -X GET http://localhost:8192/printers -sS | jq
```
```json
{
  "Ok": true,
  "Count": 2,
  "Aliases": {
    "Count": 2,
    "PrinterID1": {
      "DeviceSerialNo": "DT518315"
    },
    "PrinterID2": {
      "DeviceSerialNo": "ZK133759"
    }
  },
  "DT518315": {
    "DeviceSerialNo": "DT518315",
    "FiscalMemoryNo": "02518315",
    "DeviceProtocol": "DATECS FP/ECR",
    "DeviceModel": "DP-25",
    "FirmwareVersion": "263453 08Nov18 1312",
    "CommentTextMaxLength": 36,
    "TaxNo": "НЕЗАДАДЕН",
    "TaxCaption": "ЕИК",
    "DeviceString": "Protocol=DATECS FP/ECR;Port=COM2;Speed=115200",
    "Host": "WQW-PC",
    "Description": "Втори етаж, счетоводството"
  },
  "ZK133759": {
    "DeviceSerialNo": "ZK133759",
    "FiscalMemoryNo": "50170895",
    "DeviceProtocol": "TREMOL ECR",
    "DeviceModel": "TREMOL M20",
    "FirmwareVersion": "Ver. 1.01 TRA20 C.S. 2541",
    "CommentTextMaxLength": 30,
    "TaxNo": "",
    "TaxCaption": "ЕИК",
    "DeviceString": "Protocol=TREMOL ECR;Port=COM1;Speed=115200",
    "Host": "WQW-PC"
  }
}
```

Same with results in XML.

```shell
C:> curl -X GET http://localhost:8192/printers?format=xml -sS
```
```xml
<Root>
   <Ok __json__bool="1">1</Ok>
   <Count>2</Count>
   <Aliases>
      <Count>2</Count>
      <PrinterID1>
         <DeviceSerialNo>DT518315</DeviceSerialNo>
      </PrinterID1>
      <PrinterID2>
         <DeviceSerialNo>ZK133759</DeviceSerialNo>
      </PrinterID2>
   </Aliases>
   <DT518315>
      <DeviceSerialNo>DT518315</DeviceSerialNo>
      <FiscalMemoryNo>02518315</FiscalMemoryNo>
      <DeviceProtocol>DATECS FP/ECR</DeviceProtocol>
      <DeviceModel>DP-25</DeviceModel>
      <FirmwareVersion>263453 08Nov18 1312</FirmwareVersion>
      <CommentTextMaxLength>36</CommentTextMaxLength>
      <TaxNo>НЕЗАДАДЕН</TaxNo>
      <TaxCaption>ЕИК</TaxCaption>
      <DeviceString>Protocol=DATECS FP/ECR;Port=COM2;Speed=115200</DeviceString>
      <Host>WQW-PC</Host>
      <Description>Втори етаж, счетоводството</Description>
   </DT518315>
   <ZK133759>
      <DeviceSerialNo>ZK133759</DeviceSerialNo>
      <FiscalMemoryNo>50170895</FiscalMemoryNo>
      <DeviceProtocol>TREMOL ECR</DeviceProtocol>
      <DeviceModel>TREMOL M20</DeviceModel>
      <FirmwareVersion>Ver. 1.01 TRA20 C.S. 2541</FirmwareVersion>
      <CommentTextMaxLength>30</CommentTextMaxLength>
      <TaxNo />
      <TaxCaption>ЕИК</TaxCaption>
      <DeviceString>Protocol=TREMOL ECR;Port=COM1;Speed=115200</DeviceString>
      <Host>WQW-PC</Host>
   </ZK133759>
</Root>
```

#### `GET` `/printers/:printer_id`

Retrieve device configuration, header texts, footer texts, tax number/caption, last receipt number/datetime and payment names.

```shell
C:> curl -X GET http://localhost:8192/printers/DT518315 -sS | jq
```
```json
{
  "Ok": true,
  "DeviceSerialNo": "DT518315",
  "FiscalMemoryNo": "02518315",
  "DeviceProtocol": "DATECS FP/ECR",
  "DeviceModel": "DP-25",
  "FirmwareVersion": "263453 08Nov18 1312",
  "CommentTextMaxLength": 28,
  "Headers": [
    "               ИМЕ НА ФИРМА",
    "              АДРЕС НА ФИРМА",
    "               ИМЕ НА ОБЕКТ",
    "              АДРЕС НА ОБЕКТ",
    "",
    ""
  ],
  "Footers": [
    "",
    ""
  ],
  "TaxNo": "НЕЗАДАДЕН",
  "TaxCaption": "ЕИК",
  "ReceiptNo": "0000048",
  "DeviceDateTime": "2019-07-19 11:51:33",
  "PaymentNames": [
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

Same with results in XML.

```shell
C:> curl -X GET http://localhost:8192/printers/DT518315?format=xml -sS
```
```xml
<Root>
   <Ok __json__bool="1">1</Ok>
   <DeviceSerialNo>DT518315</DeviceSerialNo>
   <FiscalMemoryNo>02518315</FiscalMemoryNo>
   <DeviceProtocol>DATECS FP/ECR</DeviceProtocol>
   <DeviceModel>DP-25</DeviceModel>
   <FirmwareVersion>263453 08Nov18 1312</FirmwareVersion>
   <CommentTextMaxLength>30</CommentTextMaxLength>
   <Headers>ИМЕ НА ФИРМА</Headers>
   <Headers>АДРЕС НА ФИРМА</Headers>
   <Headers>ИМЕ НА ОБЕКТ</Headers>
   <Headers>АДРЕС НА ОБЕКТ</Headers>
   <Headers />
   <Headers />
   <Footers />
   <Footers />
   <TaxNo>НЕЗАДАДЕН</TaxNo>
   <TaxCaption>ЕИК</TaxCaption>
   <ReceiptNo>0000081</ReceiptNo>
   <DeviceDateTime>2019-07-23 18:07:01</DeviceDateTime>
   <PaymentNames>В БРОЙ</PaymentNames>
   <PaymentNames>С КАРТА</PaymentNames>
   <PaymentNames>НЗОК</PaymentNames>
   <PaymentNames>ВАУЧЕР</PaymentNames>
   <PaymentNames>КУПОН</PaymentNames>
   <PaymentNames />
   <PaymentNames />
</Root>
```

#### `POST` `/printers/:printer_id`

Retrieve device configuration only. This will not communicate with the device if all needed data was retrieved on previous request.

```shell
C:> curl -X POST http://localhost:8192/printers/DT518315 ^
         --data "{ }" -sS | jq
```
```json
{
  "Ok": true,
  "DeviceSerialNo": "DT518315",
  "FiscalMemoryNo": "02518315",
  "DeviceProtocol": "DATECS FP/ECR",
  "DeviceModel": "DP-25",
  "FirmwareVersion": "263453 08Nov18 1312",
  "CommentTextMaxLength": 28
}
```

Retrieve device configuration, operator name and default password.

```shell
C:> curl -X POST http://localhost:8192/printers/DT518315 ^
         --data "{ \"Operator\": { \"Code\": 1 } }" -sS | jq
```
```json
{
  "Ok": true,
  "DeviceSerialNo": "DT518315",
  "FiscalMemoryNo": "02518315",
  "DeviceProtocol": "DATECS FP/ECR",
  "DeviceModel": "DP-25",
  "FirmwareVersion": "263453 08Nov18 1312",
  "CommentTextMaxLength": 28,
  "Operator": {
    "Code": 1,
    "Name": "Оператор 1",
    "DefaultPassword": "1"
  }
}
```

Retrieve device configuration and tax number/caption only

```shell
C:> curl -X POST http://localhost:8192/printers/DT518315 ^
         --data "{ \"IncludeTaxNo\": true }" -sS | jq
```
```json
{
  "Ok": true,
  "DeviceSerialNo": "DT518315",
  "FiscalMemoryNo": "02518315",
  "DeviceProtocol": "DATECS FP/ECR",
  "DeviceModel": "DP-25",
  "FirmwareVersion": "263453 08Nov18 1312",
  "CommentTextMaxLength": 28,
  "TaxNo": "НЕЗАДАДЕН",
  "TaxCaption": "ЕИК"
}
```

#### `GET` `/printers/:printer_id/status`

Get device status and current clock.

```shell
C:> curl -X GET http://localhost:8192/printers/DT518315/status -sS | jq
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
```shell
C:> curl -X POST http://localhost:8192/printers/DT518315/receipt ^
         --data-binary @data-utf8.txt -sS | jq
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

```json
{
    "ReceiptType": 2,
    "Operator": {
        "Code": "1",
        "Name": "Иван Иванов",
        "Password": "****"
    },
    "Reversal": {
        "ReversalType": 1,
        "ReceiptNo": "0000056",
        "ReceiptDate": "2019-07-19 14:05:18",
        "FiscalMemoryNo": "02518315",
    },
    "UniqueSaleNo": "DT518315-0001-1234567",
    "Rows": [
        {
            "ItemName": "Продукт 1",
            "Price": 12.34,
            "Quantity": -1
        }
    ]
}
```
```shell
C:> curl -X POST http://localhost:8192/printers/DT518315/receipt ^
         --data-binary @data-utf8.txt -sS | jq
```
```json
{
  "Ok": true,
  "ReceiptNo": "0000061",
  "ReceiptDateTime": "2019-07-22 11:46:23",
  "DeviceSerialNo": "DT518315",
  "FiscalMemoryNo": "02518315"
}
```

Following `data-utf8.txt` prints an extended receipt for an invoice (`ReceiptType` is 3, see below) paid in total by bank card. In `Rows` the line for "Продукт 3" is specified in short array form skipping field names altogether. For "Продукт 4" only name and price are specified while tax group defaults to `2` (or `"Б"`), quantity defaults to `1` and discount defaults to `0`.

```json
{
    "ReceiptType": 3,
    "Operator": {
        "Code": "1",
        "Name": "Иван Иванов",
        "Password": "****"
    },
    "Invoice": {
        "DocNo": "1237",
        "CgTaxNo": "130395814",
        "CgTaxNoType": 0,
        "CgVatNo": "BG130395814",
        "CgName": "Униконт Софт ООД",
        "CgAddress": "София, бул. Тотлебен №85-87",
        "CgPrsReceive": "В. Висулчев",
    },
    "UniqueSaleNo": "DT518315-0001-0001234",
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
        [ "Продукт 3", 5.67, "Б", 3.5, 15 ],
        [ "Продукт 4", 2.00  ],
        { "PaymentType": 2 },
    ]
}
```
```shell
C:> curl -X POST http://localhost:8192/printers/DT518315/receipt ^
         --data-binary @data-utf8.txt -sS | jq
```
```json
{
  "Ok": true,
  "ReceiptNo": "0000065",
  "ReceiptDateTime": "2019-07-22 12:05:55",
  "DeviceSerialNo": "DT518315",
  "FiscalMemoryNo": "02518315"
}
```

Duplicate last receipt. Can be executed only once immediately after printing a receipt (or the command fails).

```shell
C:> curl -X POST http://localhost:8192/printers/DT518315/receipt ^
         --data "{ \"PrintDuplicate\": true }" -sS | jq
```
```json
{
  "Ok": false,
  "ErrorText": "Време за достъп изтече в очакване на отговор"
}
```

Print duplicate receipt (by receipt number) from device Electronic Journal.

```shell
C:> curl -X POST http://localhost:8192/printers/DT518315/receipt ^
         --data "{ \"PrintDuplicate\": true, \"Invoice\": { \"DocNo\": 57 } }" -sS | jq
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
`ucsFscRcpSale`       | 1       | Print fiscal receipt
`ucsFscRcpReversal`   | 2       | Print reversal receipt
`ucsFscRcpInvoice`    | 3       | Print extended fiscal receipt for invoice
`ucsFscRcpCreditNote` | 4       | Print extended reversal receipt for credit note
`ucsFscRcpOrderList`  | 5       | Print order-list on kitchen printer

Supported `PaymentType` values:

Name                  | Value | Alt  | Description                       | Device text              | XML code
----                  | ----- | ---- | -----------                       | ----                     | ----
`ucsFscPmtCash`       | 1     |      | Payment in cash                   | "В БРОЙ", "Лева"         | `SCash`
`ucsFscPmtCard`       | 2     |      | Payment with debit/credit card    | "ДЕБ.КАРТА", "Карта"     | `SCards`
`ucsFscPmtBank`       | 3     |      | Wire transfer (for invoices only) | "КРЕДИТ", "Банка"        | `SW`
`ucsFscPmtCheque`     | 4     |      | Payment by cheque                 | "ЧЕК", "Чек"             | `SChecks`
`ucsFscPmtCustom1`    | -1    | 5    | First custom payment              | "КУПОН", "Талон"         | `ST`
`ucsFscPmtCustom2`    | -2    | 6    | Second custom payment             | "ВАУЧЕР", "В.Талон"      | `SOT`
`ucsFscPmtCustom3`    | -3    | 7    | Third custom payment              | "НЗОК", "Резерв 1"       | `SR1`
`ucsFscPmtCustom4`    | -4    | 8    | Fourth custom payment             | "Резерв 2"               | `SR2`

Supported `ReversalType` values:

Name                        | Value | Description
----                        | ----- | -----------
`ucsFscRevOperatorError`    | 0     | Operator entry error (default)
`ucsFscRevRefund`           | 1     | Refund defective/returned goods
`ucsFscRevTaxBaseReduction` | 2     | Reduction of price/quantity of items in an invoice. Use for credit notes only

Supported `TaxNoType` values

Name                        | Value | Description
----                        | ----- | -----------
`ucsFscTxnEIC`              | 0     | ЕИК a.k.a Bulstat (default)
`ucsFscTxnCitizenNo`        | 1     | ЕГН
`ucsFscTxnForeignerNo`      | 2     | ЛНЧ
`ucsFscTxnOfficialNo`       | 3     | Служебен номер


#### `GET` `/printers/:printer_id/deposit`

Retrieve service deposit and service withdraw totals.

```shell
C:> curl -X GET http://localhost:8192/printers/DT518315/deposit -sS | jq
```
```json
{
  "Ok": true,
  "TotalAvailable": 349.68,
  "TotalDeposits": 381.34,
  "TotalWithdraws": 123
}
```

#### `POST` `/printers/:printer_id/deposit`

Print service deposit.

```shell
C:> curl -X POST http://localhost:8192/printers/DT518315/deposit ^
         --data "{ \"Amount\": 12.34 }" -sS | jq
```
```json
{
  "Ok": true,
  "ReceiptNo": "0000050",
  "ReceiptDateTime": "2019-07-19 12:02:08",
  "TotalAvailable": 362.02,
  "TotalDeposits": 393.68,
  "TotalWithdraws": 123
}
```

Print service withdraw.

```shell
C:> curl -X POST http://localhost:8192/printers/DT518315/deposit ^
         --data "{ \"Amount\": -56.78, \"Operator\": { \"Code\": \"2\", \"Password\": \"****\" } }" -sS | jq
```
```json
{
  "Ok": true,
  "ReceiptNo": "0000052",
  "ReceiptDateTime": "2019-07-19 12:03:41",
  "TotalAvailable": 248.46,
  "TotalDeposits": 393.68,
  "TotalWithdraws": 236.56
}
```

#### `POST` `/printers/:printer_id/report`

Print device reports. Supports daily X or Z reports and monthly (by date range) reports.

```shell
C:> curl -X POST http://localhost:8192/printers/DT518315/report ^
         --data "{ \"ReportType\": 1 }" -sS | jq
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

```shell
C:> curl -X GET http://localhost:8192/printers/DT518315/datetime -sS | jq
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

```shell
C:> curl -X POST http://localhost:8192/printers/DT518315/datetime ^
         --data "{ \"DeviceDateTime\": \"2019-07-19 11:58:31\" }" -sS | jq
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

```shell
C:> curl -X POST http://localhost:8192/printers/DT518315/datetime ^
         --data "{ \"DeviceDateTime\": \"2019-07-19 11:58:31\", \"AdjustTolerance\": 60 }" -sS | jq
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

```shell
C:> curl -X GET http://localhost:8192/printers/DT518315/totals -sS | jq
```
```json
{
    "Ok": true,
    "NumReceipts": 29,
    "TotalAvailable": 361.04,
    "TotalDeposits": 393.68,
    "TotalWithdraws": 236.56,
    "TotalsByTaxGroups": [
        { "TaxGroup": 1, "VatPercent": 0, "Amount": 0 },
        { "TaxGroup": 2, "VatPercent": 20, "Amount": 421.46 },
        { "TaxGroup": 3, "VatPercent": 20, "Amount": 0 },
        { "TaxGroup": 4, "VatPercent": 9, "Amount": 0 },
        { "TaxGroup": 5, "VatPercent": 0, "Amount": 0 },
        { "TaxGroup": 6, "VatPercent": 0, "Amount": 0 },
        { "TaxGroup": 7, "VatPercent": 0, "Amount": 0 },
        { "TaxGroup": 8, "VatPercent": 0, "Amount": 0 }
    ],
    "TotalsByPaymentTypes": [
        { "PaymentType": 1, "PaymentName": "В БРОЙ", "Amount": 226.26 },
        { "PaymentType": 2, "PaymentName": "С КАРТА", "Amount": 195.2 },
        { "PaymentType": 3, "PaymentName": "НЗОК", "Amount": 0 },
        { "PaymentType": 4, "PaymentName": "КРЕДИТ", "Amount": 0 },
        { "PaymentType": 5, "PaymentName": "ВАУЧЕР", "Amount": 0 },
        { "PaymentType": 6, "PaymentName": "КУПОН", "Amount": 0 },
        { "PaymentType": 7, "PaymentName": "", "Amount": 0 },
        { "PaymentType": 8, "PaymentName": "", "Amount": 0 }
    ]
}
```

#### `POST` `/printers/:printer_id/drawer`

Send impulse from fiscal device to open connected drawer.

```shell
C:> curl -X POST http://localhost:8192/printers/DT518315/drawer ^
         --data "{ \"IsOpen\": true }" -sS | jq
```
```json
{
    "Ok": true,
}
```

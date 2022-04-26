## REST service protocol description

All URLs are case-insensitive i.e. `/printers`, `/Printers` and `/PRINTERS` are the same address. Printer IDs are case-insensitive too. Printers are addressed by `:printer_id` which can either be the serial number as reported by the fiscal device or an alias assigned in the service configuration.

Both request and response payloads by default are of `application/json` type (in utf-8 encoding) whether explicitly requested in `Accept` and `Content-Type` headers or not. Use `format=xml` as URL query string parameter to change results format to XML with content-type of `text/xml; charset=utf-8`.

Use `request_id=N2qbikc5lUU` as URL query string parameter for idempotent `POST` requests (and general cache control) in order to prevent duplicating fiscal transactions when repeating requests because of a timeout or connectivity issues. If a request is repeated with the same payload and `request_id` then the results would be fetched directly from service cache without communicating with the fiscal device, provided that the previous execution of the same `request_id` succeeded.

All endpoints return `"Ok": true` on success and in case of failure include `"ErrorText": "Описание на грешка"` localized error text in the response.

All endpoints support "Includes Middleware". This allows for instance to set `"IncludePaymentNames": true` in `POST /printers/:printer_id/deposit` request to return available payment names along with standard results. Supported includes are `IncludeHeaders`, `IncludeFooters`, `IncludeTaxNo`, `IncludeTaxRates`, `IncludeReceiptNo`, `IncludePaymentNames` and `IncludeAll` which activates all previous includes and `IncludeOperators` and `IncludeDepartments` which are separate from `IncludeAll` catch-all.

These are the REST service endpoints supported:

Verb  | Endpoint                                                            | Description
----  | --------                                                            | -----------
`GET` | [`/printers`](#get-printers)                                        | List currently configured devices
`GET` | [`/printers/:printer_id`](#get-printersprinter_id)                  | Retrieve device configuration, header texts, footer texts, tax number/caption, last receipt number/datetime and payment names
`POST`| [`/printers/:printer_id`](#post-printersprinter_id)                 | Retrieve device configuration only. This will not communicate with the device if all needed data was retrieved on previous request
`GET` | [`/printers/:printer_id/status`](#get-printersprinter_idstatus)     | Get device status and current clock
`POST`| [`/printers/:printer_id/receipt`](#post-printersprinter_idreceipt)  | Print fiscal receipt, reversal, invoice or credit note
`GET` | [`/printers/:printer_id/deposit`](#get-printersprinter_iddeposit)   | Retrieve service deposit and service withdraw totals
`POST`| [`/printers/:printer_id/deposit`](#post-printersprinter_iddeposit)  | Print service deposit
`POST`| [`/printers/:printer_id/report`](#post-printersprinter_idreport)    | Print device reports. Supports daily X or Z reports and monthly (by date range) reports
`GET` | [`/printers/:printer_id/datetime`](#get-printersprinter_iddatetime) | Get current device date/time
`POST`| [`/printers/:printer_id/datetime`](#post-printersprinter_iddatetime)| Set device date/time
`GET` | [`/printers/:printer_id/totals`](#get-printersprinter_idtotals)     | Get device totals since last Z report grouped by payment types and tax groups
`POST`| [`/printers/:printer_id/drawer`](#post-printersprinter_iddrawer)    | Send impulse from fiscal device to open connected drawer

The `UcsFPHub` service endpoints return minimized JSON so sample `curl` requests below use [`jq`](https://stedolan.github.io/jq/) (a.k.a. **J**SON **Q**uery) utility to format response in human readable JSON.

#### `GET` `/printers`

List currently configured devices.

```shell
C:> curl -X GET http://localhost:8192/printers -sS | jq
```
```json
{
  "Ok": true,
  "Count": 4,
  "ZK133759": {
    "DeviceSerialNo": "ZK133759",
    "FiscalMemoryNo": "50170895",
    "DeviceProtocol": "TREMOL",
    "DeviceModel": "TREMOL M20",
    "FirmwareVersion": "Ver. 1.01 TRA20 C.S. 2541",
    "CharsPerLine": 32,
    "CommentTextMaxLength": 30,
    "ItemNameMaxLength": 26,
    "TaxNo": "",
    "TaxCaption": "ЕИК",
    "DeviceString": "Protocol=TREMOL;Port=COM1;Speed=115200",
    "DeviceHost": "WQW-PC",
    "DevicePort": "COM1",
    "Autodetected": true
  },
  "DT518315": {
    "DeviceSerialNo": "DT518315",
    "FiscalMemoryNo": "02518315",
    "DeviceProtocol": "DATECS",
    "DeviceModel": "DP-25",
    "FirmwareVersion": "263453 08Nov18 1312",
    "CharsPerLine": 42,
    "CommentTextMaxLength": 36,
    "ItemNameMaxLength": 22,
    "TaxNo": "НЕЗАДАДЕН",
    "TaxCaption": "ЕИК",
    "DeviceString": "Protocol=DATECS;Port=COM2;Speed=115200",
    "DeviceHost": "WQW-PC",
    "DevicePort": "COM2"
  },
  "DY450626": {
    "DeviceSerialNo": "DY450626",
    "FiscalMemoryNo": "36608662",
    "DeviceProtocol": "DAISY",
    "DeviceModel": "CompactM",
    "FirmwareVersion": "ONL02-4.01BG 29-10-2018 11:34 F36A",
    "CharsPerLine": 32,
    "CommentTextMaxLength": 28,
    "ItemNameMaxLength": 20,
    "TaxNo": "---------------",
    "TaxCaption": "ЕИК",
    "DeviceString": "Protocol=DAISY;Port=COM11;Speed=9600",
    "DeviceHost": "WQW-PC",
    "DevicePort": "COM11,9600",
    "Autodetected": true
  },
  "DT577430": {
    "DeviceSerialNo": "DT577430",
    "FiscalMemoryNo": "02577430",
    "DeviceProtocol": "DATECS/X",
    "DeviceModel": "DP-25X",
    "FirmwareVersion": "264205 22Jan19 1629",
    "CharsPerLine": 42,
    "CommentTextMaxLength": 40,
    "ItemNameMaxLength": 72,
    "TaxNo": "НЕЗАДАДЕН",
    "TaxCaption": "ЕИК",
    "DeviceString": "Protocol=DATECS/X;IP=192.168.0.20",
    "DeviceHost": "WQW-PC",
    "DevicePort": "192.168.0.20"
  },
  "Aliases": {
    "Count": 2,
    "PrinterID1": {
      "DeviceSerialNo": "DT518315"
    },
    "PrinterID2": {
      "DeviceSerialNo": "DT577430"
    }
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
   <Count>4</Count>
   <ZK133759>
      <DeviceSerialNo>ZK133759</DeviceSerialNo>
      <FiscalMemoryNo>50170895</FiscalMemoryNo>
      <DeviceProtocol>TREMOL</DeviceProtocol>
      <DeviceModel>TREMOL M20</DeviceModel>
      <FirmwareVersion>Ver. 1.01 TRA20 C.S. 2541</FirmwareVersion>
      <CharsPerLine>32</CharsPerLine>
      <CommentTextMaxLength>30</CommentTextMaxLength>
      <ItemNameMaxLength>26</ItemNameMaxLength>
      <TaxNo />
      <TaxCaption>ЕИК</TaxCaption>
      <DeviceString>Protocol=TREMOL;Port=COM1;Speed=115200</DeviceString>
      <DeviceHost>WQW-PC</DeviceHost>
      <DevicePort>COM1</DevicePort>
      <Autodetected __json__bool="1">1</Autodetected>
   </ZK133759>
   <DT518315>
      <DeviceSerialNo>DT518315</DeviceSerialNo>
      <FiscalMemoryNo>02518315</FiscalMemoryNo>
      <DeviceProtocol>DATECS</DeviceProtocol>
      <DeviceModel>DP-25</DeviceModel>
      <FirmwareVersion>263453 08Nov18 1312</FirmwareVersion>
      <CharsPerLine>42</CharsPerLine>
      <CommentTextMaxLength>36</CommentTextMaxLength>
      <ItemNameMaxLength>22</ItemNameMaxLength>
      <TaxNo>НЕЗАДАДЕН</TaxNo>
      <TaxCaption>ЕИК</TaxCaption>
      <DeviceString>Protocol=DATECS;Port=COM2;Speed=115200</DeviceString>
      <DeviceHost>WQW-PC</DeviceHost>
      <DevicePort>COM2</DevicePort>
   </DT518315>
   <DY450626>
      <DeviceSerialNo>DY450626</DeviceSerialNo>
      <FiscalMemoryNo>36608662</FiscalMemoryNo>
      <DeviceProtocol>DAISY</DeviceProtocol>
      <DeviceModel>CompactM</DeviceModel>
      <FirmwareVersion>ONL02-4.01BG 29-10-2018 11:34 F36A</FirmwareVersion>
      <CharsPerLine>32</CharsPerLine>
      <CommentTextMaxLength>28</CommentTextMaxLength>
      <ItemNameMaxLength>20</ItemNameMaxLength>
      <TaxNo>---------------</TaxNo>
      <TaxCaption>ЕИК</TaxCaption>
      <DeviceString>Protocol=DAISY;Port=COM11;Speed=9600</DeviceString>
      <DeviceHost>WQW-PC</DeviceHost>
      <DevicePort>COM11,9600</DevicePort>
      <Autodetected __json__bool="1">1</Autodetected>
   </DY450626>
   <DT577430>
      <DeviceSerialNo>DT577430</DeviceSerialNo>
      <FiscalMemoryNo>02577430</FiscalMemoryNo>
      <DeviceProtocol>DATECS/X</DeviceProtocol>
      <DeviceModel>DP-25X</DeviceModel>
      <FirmwareVersion>264205 22Jan19 1629</FirmwareVersion>
      <CharsPerLine>42</CharsPerLine>
      <CommentTextMaxLength>40</CommentTextMaxLength>
      <ItemNameMaxLength>72</ItemNameMaxLength>
      <TaxNo>НЕЗАДАДЕН</TaxNo>
      <TaxCaption>ЕИК</TaxCaption>
      <DeviceString>Protocol=DATECS/X;IP=192.168.0.20</DeviceString>
      <DeviceHost>WQW-PC</DeviceHost>
      <DevicePort>192.168.0.20</DevicePort>
   </DT577430>
   <Aliases>
      <Count>2</Count>
      <PrinterID1>
         <DeviceSerialNo>DT518315</DeviceSerialNo>
      </PrinterID1>
      <PrinterID2>
         <DeviceSerialNo>DT577430</DeviceSerialNo>
      </PrinterID2>
   </Aliases>
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
  "DeviceProtocol": "DATECS",
  "DeviceModel": "DP-25",
  "FirmwareVersion": "263453 08Nov18 1312",
  "CharsPerLine": 42,
  "CommentTextMaxLength": 36,
  "ItemNameMaxLength": 22,
  "TaxNo": "НЕЗАДАДЕН",
  "TaxCaption": "ЕИК",
  "Headers": [
    "ИМЕ НА ФИРМА",
    "АДРЕС НА ФИРМА",
    "ИМЕ НА ОБЕКТ",
    "АДРЕС НА ОБЕКТ",
    "",
    ""
  ],
  "Footers": [
    "",
    ""
  ],
  "LastReceiptNo": "179",
  "LastReceiptDateTime": "2019-11-12 14:01:08",
  "PaymentNames": [
    "В БРОЙ",
    "С КАРТА",
    "ПО БАНКА",
    "",
    "КУПОН",
    "ВАУЧЕР",
    "НЗОК",
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
   <DeviceProtocol>DATECS</DeviceProtocol>
   <DeviceModel>DP-25</DeviceModel>
   <FirmwareVersion>263453 08Nov18 1312</FirmwareVersion>
   <CharsPerLine>42</CharsPerLine>
   <CommentTextMaxLength>36</CommentTextMaxLength>
   <ItemNameMaxLength>22</ItemNameMaxLength>
   <TaxNo>НЕЗАДАДЕН</TaxNo>
   <TaxCaption>ЕИК</TaxCaption>
   <Headers>ИМЕ НА ФИРМА</Headers>
   <Headers>АДРЕС НА ФИРМА</Headers>
   <Headers>ИМЕ НА ОБЕКТ</Headers>
   <Headers>АДРЕС НА ОБЕКТ</Headers>
   <Headers />
   <Headers />
   <Footers />
   <Footers />
   <LastReceiptNo>179</LastReceiptNo>
   <LastReceiptDateTime>2019-11-12 14:01:36</LastReceiptDateTime>
   <PaymentNames>В БРОЙ</PaymentNames>
   <PaymentNames>С КАРТА</PaymentNames>
   <PaymentNames>ПО БАНКА</PaymentNames>
   <PaymentNames />
   <PaymentNames>КУПОН</PaymentNames>
   <PaymentNames>ВАУЧЕР</PaymentNames>
   <PaymentNames>НЗОК</PaymentNames>
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
  "DeviceProtocol": "DATECS",
  "DeviceModel": "DP-25",
  "FirmwareVersion": "263453 08Nov18 1312",
  "CharsPerLine": 42,
  "CommentTextMaxLength": 36,
  "ItemNameMaxLength": 22
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
  "DeviceProtocol": "DATECS",
  "DeviceModel": "DP-25",
  "FirmwareVersion": "263453 08Nov18 1312",
  "CharsPerLine": 42,
  "CommentTextMaxLength": 36,
  "ItemNameMaxLength": 22,
  "Operator": {
    "Code": 1,
    "Name": "Иван Иванов",
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
  "DeviceProtocol": "DATECS",
  "DeviceModel": "DP-25",
  "FirmwareVersion": "263453 08Nov18 1312",
  "CharsPerLine": 42,
  "CommentTextMaxLength": 36,
  "ItemNameMaxLength": 22,
  "TaxNo": "НЕЗАДАДЕН",
  "TaxCaption": "ЕИК"
}
```

#### `GET` `/printers/:printer_id/status`

Get device status.

```shell
C:> curl -X GET http://localhost:8192/printers/DT518315/status -sS | jq
```
```json
{
  "Ok": true,
  "DeviceStatusCode": "busy",
  "DeviceStatus": "Устройството е заето"
}
```

Supported `DeviceStatusCode` values:

Name            | Value | Description
----            | ----- | -----------
`Ready`         | 0     | The device is ready to accept commands
`Busy`          | 1     | The device is busy printing
`Failed`        | 2     | An error occurred while accessing/operating the device


#### `POST` `/printers/:printer_id/receipt`

Print fiscal receipt, reversal, invoice or credit note.

Following `data-utf8.txt` prints a fiscal receipt (`ReceiptType` is 1, see below) for two products, second one is with discount. The receipt in paid first 10.00 leva with a bank card and the rest in cash. After all receipt totals are printed a free-text line outputs current client loyalty card number used for information.

```json
{
    "ReceiptType": "Sale",
    "Operator": {
        "Code": "1",
        "Name": "Иван Петров",
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
            "PaymentType": "Card"
        },
        {
            "PaymentType": "Cash"
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
  "ReceiptNo": "180",
  "ReceiptDateTime": "2019-11-12 14:05:17",
  "DeviceSerialNo": "DT518315",
  "FiscalMemoryNo": "02518315"
}
```

Following `data-utf8.txt` prints a reversal receipt (`ReceiptType` is 2, see below) for the first products of the previous sale `180` and a bar-code with info about the reversed receipt.

```json
{
    "ReceiptType": 2,
    "Operator": {
        "Code": "1",
        "Name": "Иван Иванов",
        "Password": "****"
    },
    "Reversal": {
        "ReversalType": "Refund",
        "ReceiptNo": "180",
        "ReceiptDateTime": "2019-11-12 14:05:17",
        "FiscalMemoryNo": "02518315",
    },
    "UniqueSaleNo": "DT518315-0001-1234567",
    "Rows": [
        {
            "ItemName": "Продукт 1",
            "Price": 12.34,
            "Quantity": 1
        },
        {
            "BarcodeType": "Code128",
            "Text": "180/2019-11-12"
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
  "ReceiptNo": "181",
  "ReceiptDateTime": "2019-11-12 14:08:24",
  "DeviceSerialNo": "DT518315",
  "FiscalMemoryNo": "02518315"
}
```

Following `data-utf8.txt` prints an extended receipt for an invoice (`ReceiptType` is 3, see below) paid in total by bank card. In `Rows` the line for "Продукт 3" is specified in short array form skipping field names altogether. For "Продукт 4" only name and price are specified while tax group defaults to `2` (or `"Б"`), quantity defaults to `1` and discount defaults to `0`.

```json
{
    "ReceiptType": "Invoice",
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
  "ReceiptNo": "65",
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

Name            | Value | Description
----            | ----- | -----------
`Sale`          | 1     | Print fiscal receipt
`Reversal`      | 2     | Print reversal receipt
`Invoice`       | 3     | Print extended fiscal receipt for invoice
`CreditNote`    | 4     | Print extended reversal receipt for credit note
`OrderList`     | 5     | Print order-list on kitchen printer

Supported `PaymentType` values:

Name            | Value | Description                       | Device text           | XML code
----            | ----- | -----------                       | -----------           | --------
`Cash`          | 1     | Payment in cash                   | "В БРОЙ", "Лева"      | `SCash`
`Cheque`        | 2     | Payment by cheque                 | "ЧЕК", "Чек"          | `SChecks`
`Coupon`        | 3     | Payment w/ coupons                | "КУПОН", "Талон"      | `ST`
`Voucher`       | 4     | Payment w/ external/food vouchers | "ВАУЧЕР", "В.Талон"   | `SOT`
`Packaging`     | 5     | Returned packing deducted         | "Амбалаж"             | `SP`
`Maintenance`   | 6     | N/A                               | "Вътрешно обслужване", "Обслужване" | `SSelf`
`Damage`        | 7     | N/A                               | "Повреди"             | `SDmg`
`Card`          | 8     | Payment with debit/credit card    | "ДЕБ.КАРТА", "Карта"  | `SCards`
`Bank`          | 9     | Wire transfer (for invoices)      | "Банка"               | `SW`
`Custom1`       | 10    | Custom payment 1                  | "Резерв 1", "НЗОК", "Отложено плащане"    | `SR1`
`Custom2`       | 11    | Custom payment 2                  | "Резерв 2", "Вътрешно потребление"        | `SR2`
`EUR`           | 12    | Payment in Euro                   | "EURO"                | N/A

Supported `ReversalType` values:

Name                | Value | Description
----                | ----- | -----------
`OperatorError`     | 0     | Operator entry error (default)
`Refund`            | 1     | Refund defective/returned goods or reduction of quantity of items in an invoice
`TaxBaseReduction`  | 2     | Reduction of price of items in an invoice

Supported `TaxNoType` values

Name            | Value | Description
----            | ----- | -----------
`EIC`           | 0     | ЕИК a.k.a Bulstat (default)
`CitizenNo`     | 1     | ЕГН
`ForeignerNo`   | 2     | ЛНЧ
`OfficialNo`    | 3     | Служебен номер

Supported `BarcodeType` values:

Name            | Value | Description
----            | ----- | -----------
`Ean8`          | 1     | EAN 8, Exactly 7 digits
`Ean13`         | 2     | EAN 13, Exactly 12 digits
`Code128`       | 3     | CODE 128, Up to 20 latin letters and digits
`QRcode`        | 4     | QR Code, Up to 45 latin letters and digits


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
  "ReceiptNo": "50",
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
  "ReceiptNo": "52",
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
         --data "{ \"ReportType\": \"Daily\" }" -sS | jq
```
```json
{
  "Ok": true,
  "ReceiptNo": "182",
  "ReceiptDateTime": "2019-11-12 18:00:16",
  "DeviceSerialNo": "DT518315",
  "FiscalMemoryNo": "02518315"
}
```

Supported `ReportType` values:

Name                 | Value | Description
----                 | ----- | -----------
`Daily`              | 1     | Prints daily X or Z report. Set `IsClear` for Z report, `IsItems` for report by items, `IsDepartments` for daily report by departments.
`MonthlyByReceiptNo` | 2     | Not implemented
`MonthlyByDate`      | 3     | Prints monthly fiscal report. Use `FromDate` and `ToDate` to specify date range.
`DailyByOperators`   | 4     | Prints daily operators report. Set `IsClear` to reset operators daily counters.


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
C:> curl -X GET http://localhost:8192/printers/ZK133759/totals -sS | jq
```
```json
{
    "Ok": true,
    "NumReceipts": 54,
    "LastZReportDateTime": "2018-01-01 00:00:00",
    "TotalAmount": 1029.15,
    "TotalReversal": 636.26,
    "TotalAvailable": 603.28,
    "TotalDeposits": 10,
    "TotalWithdraws": 0,
    "TotalsByTaxGroups": [
        { "TaxGroup": 1, "VatRate": 0, "Amount": 123, "Reversal": 21 },
        { "TaxGroup": 2, "VatRate": 20, "Amount": 906.15, "Reversal": 615.26 },
        { "TaxGroup": 3, "VatRate": 20, "Amount": 0, "Reversal": 0 },
        { "TaxGroup": 4, "VatRate": 9, "Amount": 0, "Reversal": 0 },
        { "TaxGroup": 5, "Amount": 0, "Reversal": 0 },
        { "TaxGroup": 6, "Amount": 0, "Reversal": 0 },
        { "TaxGroup": 7, "Amount": 0, "Reversal": 0 },
        { "TaxGroup": 8, "Amount": 0, "Reversal": 0 }
    ],
    "TotalsByPaymentTypes": [
        { "PaymentType": 1, "PaymentName": "Лева", "Amount": 879 },
        { "PaymentType": 2, "PaymentName": "Чек", "Amount": 85.15 },
        { "PaymentType": 3, "PaymentName": "Талон", "Amount": 2 },
        { "PaymentType": 4, "PaymentName": "В.Талон", "Amount": 0 },
        { "PaymentType": 5, "PaymentName": "Амбалаж", "Amount": 0 },
        { "PaymentType": 6, "PaymentName": "Обслужване", "Amount": 0 },
        { "PaymentType": 7, "PaymentName": "Повреди", "Amount": 0 },
        { "PaymentType": 8, "PaymentName": "Карта", "Amount": 63 },
        { "PaymentType": 9, "PaymentName": "Банка", "Amount": 0 },
        { "PaymentType": 10, "PaymentName": "Резерв 1", "Amount": 0 },
        { "PaymentType": 11, "PaymentName": "Резерв 2", "Amount": 0 }
    ]
}
```

#### `POST` `/printers/:printer_id/drawer`

Send impulse from fiscal device to open the connected drawer. Use to manually open the drawer only as most fiscal devices automatically open the drawer already after a receipt is successfully printed.

```shell
C:> curl -X POST http://localhost:8192/printers/DT518315/drawer ^
         --data "{ \"IsOpen\": true }" -sS | jq
```
```json
{
    "Ok": true,
}
```

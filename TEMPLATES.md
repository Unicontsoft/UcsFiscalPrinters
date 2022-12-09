## Label Templates configuration

Here is a sample `LabelTemplates.conf` file

```json
{
    "L01": {
        "Controls": {
            "Name": {
                "Width": 28,
                "Height": 2,
                "HorAlign": "center"
            },
            "QRCode": {
                "Width": 63,
                "Height": 2,
                "Wrap": "barcode"
            },
            "Content": {
                "Width": 21,
                "Height": 6
            }
        }
    }
}
```

Root object contains list of named templates and each template has a `Controls` subkey which contains list of named controls.

A printer label named `L01` must be prepared and uploaded to the device beforehand using Datecs Label Editor or a similar tool. The label has to bind its text fields and barcodes to variables `V00` to `Vxx` so that these are replaced on form submission done by the LABEL protocol.

In the sample configuration above `Name` control should be implemented by two text fields bound to `V00` and `V01` respectively, `QRCode` should be a 2D-barcode field bound to both `V02V03` (because printer variables are limited to 63 symbols only) and finally `Content` should be implemented by 6 text fields bound to `V04` to `V09` variables respectively.

When printing a receipt to a LABEL protocol device the text for each control is supplied as free-text line (`ucsRwtText`) in `Name=Text` format i.e.

```json
{
    "ReceiptType": "Sale",
    "Rows": [
        [ "Item=Име на продукт" ],
        [ "QRCode=12345678$PCS" ],
        [ "Content=Партида: L123^pПроизход: Bulgaria" ]
    ]
}
```

Order of control texts in the sales receipt does not affect output form variables. It is only the template configuration which determine which texts go to which form variables.

All texts are wrapped and aligned according to control's width while `^p` symbol is replaced with a new line before printing so that consecutive lines go to separate form variables on the label.

using System.Globalization;
using UcsFiscalPrinters;

namespace Demo1
{
    class Program
    {
        static void Main(string[] args)
        {
            // will use Datecs fiscal printer
            var prot = new cICLProtocol();

            // setup localized commands for Serbia
            if (CultureInfo.CurrentCulture.Name.Substring(0, 5) == "sr-SP")
            {
                // open fiscal receipt command uses ';' for first separator. apparently Datecs made this protocol 
                //   incompatibiliy for Serbia on purpose to prevent interoperability with BG software.
                prot.SetLocalizedCommand("pvPrintReceipt", "FiscalOpen2", Param: "%1;%2,%3");
                prot.SetLocalizedCommand("pvPrintReceipt", "FiscalOpen3", Param: "%1;%2,%3");
            }
            
            // cast to printer independent IDeviceProtocol interface
            var fp = (IDeviceProtocol)prot;

            // format of device string: port[,speed][,data,parity,stop]
            fp.Init("COM3,9600");

            // queue new fiscal reciept
            fp.StartReceipt(UcsFiscalReceiptTypeEnum.ucsFscRetFiscal, "1", "Operator 1", "0000");

            // queue sale of 5 items of "Product 1" for 1.23 each in second VAT group (20%)
            fp.AddPLU("Product 1", 1.23, 5, 2);

            // queue payment of 10.00 in cash
            fp.AddPayment(UcsFiscalPaymentTypeEnum.ucsFscPmtCash, "Cash", 10.0);

            // print all queue in transaction here
            fp.EndReceipt(null);
        }
    }
}

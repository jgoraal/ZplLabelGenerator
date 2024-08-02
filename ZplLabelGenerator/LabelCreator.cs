using KeePass.Plugins;
using KeePassLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.NetworkInformation;

namespace ZplLabelGenerator
{
    internal class LabelCreator
    {
        private readonly IPluginHost _host;

        public LabelCreator(IPluginHost host)
        {
            _host = host;
        }

        public void GenerateLabel()
        {
            PwEntry selectedEntry = _host.MainWindow.GetSelectedEntry(false);
            if (selectedEntry == null)
            {
                MessageCreator.CreateErrorMessage(ErrorMessages.EntriesNotFoundError);
                return;
            }

            Dictionary<string, string> entriesStrings = selectedEntry.Strings
                .Where(pair => !pair.Key.Equals("Notes") && !pair.Key.Equals("Password") && !pair.Key.Equals("URL"))
                .ToDictionary(k => k.Key, v => v.Value.ReadString());

            if (!entriesStrings.Any())
            {
                throw new ApplicationException(ErrorMessages.NoDataToExport);
            }

            string hostname = "Hostname: "+ selectedEntry.Strings.ReadSafe(PwDefs.TitleField);
            string purchaseDate = "Data zakupu: " + DateTime.Now.ToString("dd/MM/yyyy");
            string warrantyDate = "Data gwarnacji: " + DateTime.Now.AddYears(1).ToString("dd/MM/yyyy");
            string labelDate = "Data wydruku etykiety: " + DateTime.Now.ToString("dd/MM/yyyy");
            string serialNumber = "Numer seryjny: " + (entriesStrings.ContainsKey("S/N") ? entriesStrings["S/N"] : "Nieznane");
            string model = "Model: " + (entriesStrings.ContainsKey("Model") ? entriesStrings["Model"] : "Nieznane");
            string macAddress = "MAC: " + GetMACAddress();

            string qrCodeData = $"{hostname}\n{purchaseDate}\n{warrantyDate}\n{labelDate}";
            string zpl = GenerateZPL(hostname, purchaseDate, warrantyDate, labelDate, serialNumber, model, macAddress, qrCodeData);

            System.IO.File.WriteAllText("label.zpl", zpl);

            MessageCreator.CreateInfoMessage("Etykieta wygenerowana", "Plik label.zpl został wygenerowany.");
        }

        private string GetMACAddress()
        {
            var macAddr =
            (
                from nic in NetworkInterface.GetAllNetworkInterfaces()
                where nic.OperationalStatus == OperationalStatus.Up
                select nic.GetPhysicalAddress().ToString()
            ).FirstOrDefault();

            return macAddr;
        }

        private string GenerateZPL(string hostname, string purchaseDate, string warrantyDate, string labelDate, string serialNumber, string model, string macAddress, string qrCodeData)
        {
            return $@"
^XA
^FO30,25^BQN,2,5^FDQA,{qrCodeData}^FS
^FO315,35^A0N,40,30^FD{serialNumber}^FS
^FO315,75^A0N,40,30^FD{model}^FS
^FO315,115^A0N,40,40^FD{macAddress}^FS
^FO315,155^A0N,40,40^FD{hostname}^FS
^FO315,195^A0N,40,40^FD{purchaseDate}^FS
^FO315,235^A0N,40,40^FD{warrantyDate}^FS
^FO315,275^A0N,40,40^FD{labelDate}^FS
^XZ";
        }
    }
}

using KeePass.Plugins;
using KeePassLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.NetworkInformation;
using System.Windows.Forms;

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
            string labelDate = DateTime.Now.ToString("dd/MM/yyyy");
            string serialNumber = "Numer seryjny: " + (entriesStrings.ContainsKey("S/N") ? entriesStrings["S/N"] : "Nieznane");
            string model = "Model: " + (entriesStrings.ContainsKey("Model") ? entriesStrings["Model"] : "Nieznane");
            string macAddress = "MAC: " + GetMACAddress();

            string qrCodeData = $"{hostname}_0D_0A{purchaseDate}_0D_0A{warrantyDate}_0D_0AData wygenerowania etykiety: {labelDate}";
            string zpl = GenerateZPL(hostname, purchaseDate, warrantyDate, labelDate, serialNumber, model, macAddress, qrCodeData);

            using (var file = new SaveFileDialog())
            {
                file.FileName = $"{_host.Database.Name}_{selectedEntry.Strings.Get("Title").ReadString()}_etykieta.zpl";

                if (file.ShowDialog() == DialogResult.OK)
                {
                    System.IO.File.WriteAllText(file.FileName, zpl);
                }
            }

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

            var formattedMacAddr = string.Join(":", Enumerable.Range(0, macAddr.Length / 2)
                                          .Select(i => macAddr.Substring(i * 2, 2)));

            return formattedMacAddr;
        }

        private string GenerateZPL(string hostname, string purchaseDate, string warrantyDate, string labelDate, string serialNumber, string model, string macAddress, string qrCodeData)
        {
            return $@"
^XA
^CI28
^PW700
^LL300
^MD30
^FO0,15^BQN,2,4^FH^FDQA,{qrCodeData}^FS
^FO0,245^A0N,25,25^FD{labelDate}^FS
^FO235,25^A0N,25,25^FD{serialNumber}^FS
^FO235,65^A0N,25,25^FD{model}^FS
^FO235,105^A0N,25,25^FD{macAddress}^FS
^FO235,145^A0N,25,25^FD{hostname}^FS
^FO235,185^A0N,25,25^FD{purchaseDate}^FS
^FO235,225^A0N,25,25^FD{warrantyDate}^FS
^XZ";
        }
    }
}

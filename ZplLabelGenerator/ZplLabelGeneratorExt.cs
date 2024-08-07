using KeePass.Plugins;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ZplLabelGenerator
{
    public sealed class ZplLabelGeneratorExt : Plugin
    {
        private IPluginHost Host { get; set; }

        public override bool Initialize(IPluginHost host)
        {
            if (host == null)
            {
                MessageCreator.CreateErrorMessage(ErrorMessages.HostError);
                return false;
            }

            Host = host;

            var tsMenuItem = GetMenuItem(PluginMenuType.Entry);
            if (tsMenuItem == null)
            {
                MessageCreator.CreateErrorMessage(ErrorMessages.MenuCreationError);
                return false;
            }

            return true;
        }

        public override void Terminate()
        {
        }


        /// <summary>
        /// Tworzy element menu dla pluginu.
        /// </summary>
        /// <returns>ToolStripMenuItem dla pluginu.</returns>
        public override ToolStripMenuItem GetMenuItem(PluginMenuType t)
        {
            if (t == PluginMenuType.Entry)
            {
                var tsMenuItem = new ToolStripMenuItem();
                tsMenuItem.Text = "Wygeneruj etykietę ZPL";
                tsMenuItem.Click += ToolsMenuItemClick;

                return tsMenuItem;
            }

            return null;
        }

        private void ToolsMenuItemClick(object sender, EventArgs e)
        {
            try
            {
                var generator = new LabelCreator(Host);
                generator.GenerateLabel();
            }
            catch (ArgumentNullException nullException)
            {
                MessageCreator.CreateErrorMessage(nullException.Message);
            }
            catch (ApplicationException applicationException)
            {
                MessageCreator.CreateErrorMessage(applicationException.Message);
            }
            catch (Exception others)
            {
                MessageCreator.CreateErrorMessage(others.Message);
            }
        }
    }
}

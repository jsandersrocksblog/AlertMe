using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;

namespace VSTOAlert
{
    public partial class AlertRibbon
    {
        private RibbonDropDownItem CreateRibbonDropDownItem()
        {
            return this.Factory.CreateRibbonDropDownItem();
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            populateAddresses();    
        }

        private void refresh()
        {
            Properties.Settings.Default.Save();
            populateAddresses();
            editBox1.Text = "";
        }

        private void populateAddresses()
        {
            this.gallery1.Items.Clear();
            if (Properties.Settings.Default.ToSearchStrings.Count != 0)
            {
                var SearchStrings = Properties.Settings.Default.ToSearchStrings;
                foreach (string toAddr in SearchStrings)
                {
                    var aItem = CreateRibbonDropDownItem();
                    aItem.Label = toAddr;
                    this.gallery1.Items.Add(aItem);
                }
            }
        }

        private void gallery1_Click(object sender, RibbonControlEventArgs e)
        {
            int index = gallery1.SelectedItemIndex;
            if (index != -1)
            {
                var SearchStrings = Properties.Settings.Default.ToSearchStrings;
                SearchStrings.RemoveAt(index);
                refresh();
            }

        }
                
        private void clearAllClick(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.ToSearchStrings.Clear();
            refresh();
        }

        private void addClick(object sender, RibbonControlEventArgs e)
        {
            if (editBox1.Text.Length > 0)
            {
                Properties.Settings.Default.ToSearchStrings.Add(editBox1.Text);
                refresh();
            }
        }
    }
}

﻿using System;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using fstonge.ExcelAddIn.Models;

namespace fstonge.ExcelAddIn.Controls
{
    public partial class SettingsUserControl : UserControl
    {
        public SettingsUserControl()
        {
            InitializeComponent();

            SuperCopyGroupBox.Text = Properties.Resources.Settings_SuperCopyLabel;
            MergeGroupBox.Text = Properties.Resources.Settings_MergeLabel;
            DelimiterLabel.Text = Properties.Resources.Settings_DelimiterLabel;
            SuperCopyDelimiter.Text = Properties.Settings.Default.SuperCopyDelimiter;

            ModeLabel.Text = Properties.Resources.Settings_ModeLabel;
            SuperCopyMode.DataSource = Enum.GetValues(typeof(SuperCopyMode))
                .OfType<SuperCopyMode>()
                .Select(s => new
                {
                    Key = s, 
                    Value = Properties.Resources.ResourceManager.GetString($"SuperCopyMode_{s:F}")
                })
                .ToList();
            SuperCopyMode.DisplayMember = "Value";
            SuperCopyMode.ValueMember = "Key";

            SuperCopyMode.Text = Properties.Settings.Default.SuperCopyMode.ToString("F");

            SkipCellsLabel.Text = Properties.Resources.Settings_SkipCellsLabel;
            SuperCopySkipCells.Value = Properties.Settings.Default.SuperCopySkipCells;

            MergeSkipCellsLabel.Text = Properties.Resources.Settings_SkipCellsLabel;
            MergeSkipCells.Value = Properties.Settings.Default.MergeSkipCells;

            LoadDefault.Text = Properties.Resources.Settings_LoadDefaultLabel;

            SuperCopyDelimiter.TextChanged += Delimiter_TextChanged;
            SuperCopyMode.SelectedValueChanged += Mode_SelectedValueChanged;
            SuperCopySkipCells.ValueChanged += SkipCells_ValueChanged;
            MergeSkipCells.ValueChanged += MergeSkipCells_ValueChanged;
            Properties.Settings.Default.PropertyChanged += Default_PropertyChanged;
            LoadDefault.Click += LoadDefault_Click;
        }

        private void LoadDefault_Click(object sender, EventArgs e)
        {
            SuperCopyDelimiter.Text = ";";
            SuperCopyMode.Text = Models.SuperCopyMode.Column.ToString("F");
            SuperCopySkipCells.Value = 0;
            MergeSkipCells.Value = 0;
        }

        private void SkipCells_ValueChanged(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.SuperCopySkipCells != SuperCopySkipCells.Value)
            {
                Properties.Settings.Default.SuperCopySkipCells = Convert.ToInt32(SuperCopySkipCells.Value);
            }

            Properties.Settings.Default.Save();
        }

        private void Mode_SelectedValueChanged(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.SuperCopyMode != (SuperCopyMode)SuperCopyMode.SelectedValue)
            {
                Properties.Settings.Default.SuperCopyMode = (SuperCopyMode)SuperCopyMode.SelectedValue;
            }

            Properties.Settings.Default.Save();
        }

        private void Delimiter_TextChanged(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.SuperCopyDelimiter != SuperCopyDelimiter.Text)
            {
                Properties.Settings.Default.SuperCopyDelimiter = SuperCopyDelimiter.Text;
            }

            Properties.Settings.Default.Save();
        }

        private void MergeSkipCells_ValueChanged(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.MergeSkipCells != MergeSkipCells.Value)
            {
                Properties.Settings.Default.MergeSkipCells = Convert.ToInt32(MergeSkipCells.Value);
            }

            Properties.Settings.Default.Save();
        }

        private void Default_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            var value = Properties.Settings.Default[e.PropertyName];
            var control = Controls.Find(e.PropertyName, true).FirstOrDefault();
            if (control is TextBox textBox)
            {
                if (textBox.Text != value.ToString())
                {
                    textBox.Text = value.ToString();
                }
            }

            if (control is ComboBox comboBox)
            {
                if (comboBox.Text != value.ToString())
                {
                    comboBox.Text = value.ToString();
                }
            }

            if (control is NumericUpDown numericUpDown)
            {
                if (numericUpDown.Value != Convert.ToInt32(value))
                {
                    numericUpDown.Value = Convert.ToInt32(value);
                }
            }
        }
    }
}

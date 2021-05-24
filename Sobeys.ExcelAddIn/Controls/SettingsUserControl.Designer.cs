
namespace Sobeys.ExcelAddIn.Controls
{
    partial class SettingsUserControl
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.SuperCopyDelimiter = new System.Windows.Forms.TextBox();
            this.DelimiterLabel = new System.Windows.Forms.Label();
            this.SuperCopyTableLayout = new System.Windows.Forms.TableLayoutPanel();
            this.ModeLabel = new System.Windows.Forms.Label();
            this.SkipCellsLabel = new System.Windows.Forms.Label();
            this.SuperCopyMode = new System.Windows.Forms.ComboBox();
            this.SuperCopySkipCells = new System.Windows.Forms.NumericUpDown();
            this.SuperCopyGroupBox = new System.Windows.Forms.GroupBox();
            this.LoadDefault = new System.Windows.Forms.Button();
            this.SuperCopyTableLayout.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.SuperCopySkipCells)).BeginInit();
            this.SuperCopyGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // SuperCopyDelimiter
            // 
            this.SuperCopyDelimiter.Dock = System.Windows.Forms.DockStyle.Fill;
            this.SuperCopyDelimiter.Location = new System.Drawing.Point(97, 3);
            this.SuperCopyDelimiter.Name = "SuperCopyDelimiter";
            this.SuperCopyDelimiter.Size = new System.Drawing.Size(88, 22);
            this.SuperCopyDelimiter.TabIndex = 1;
            // 
            // DelimiterLabel
            // 
            this.DelimiterLabel.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.DelimiterLabel.AutoSize = true;
            this.DelimiterLabel.Location = new System.Drawing.Point(3, 6);
            this.DelimiterLabel.Name = "DelimiterLabel";
            this.DelimiterLabel.Size = new System.Drawing.Size(61, 16);
            this.DelimiterLabel.TabIndex = 0;
            this.DelimiterLabel.Text = "Delimiter";
            this.DelimiterLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // SuperCopyTableLayout
            // 
            this.SuperCopyTableLayout.AutoSize = true;
            this.SuperCopyTableLayout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.SuperCopyTableLayout.ColumnCount = 2;
            this.SuperCopyTableLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.SuperCopyTableLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.SuperCopyTableLayout.Controls.Add(this.SuperCopyDelimiter, 1, 0);
            this.SuperCopyTableLayout.Controls.Add(this.DelimiterLabel, 0, 0);
            this.SuperCopyTableLayout.Controls.Add(this.ModeLabel, 0, 1);
            this.SuperCopyTableLayout.Controls.Add(this.SkipCellsLabel, 0, 2);
            this.SuperCopyTableLayout.Controls.Add(this.SuperCopyMode, 1, 1);
            this.SuperCopyTableLayout.Controls.Add(this.SuperCopySkipCells, 1, 2);
            this.SuperCopyTableLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.SuperCopyTableLayout.Location = new System.Drawing.Point(3, 18);
            this.SuperCopyTableLayout.Name = "SuperCopyTableLayout";
            this.SuperCopyTableLayout.RowCount = 3;
            this.SuperCopyTableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.SuperCopyTableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.SuperCopyTableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.SuperCopyTableLayout.Size = new System.Drawing.Size(188, 86);
            this.SuperCopyTableLayout.TabIndex = 1;
            // 
            // ModeLabel
            // 
            this.ModeLabel.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.ModeLabel.AutoSize = true;
            this.ModeLabel.Location = new System.Drawing.Point(3, 35);
            this.ModeLabel.Name = "ModeLabel";
            this.ModeLabel.Size = new System.Drawing.Size(43, 16);
            this.ModeLabel.TabIndex = 2;
            this.ModeLabel.Text = "Mode";
            // 
            // SkipCellsLabel
            // 
            this.SkipCellsLabel.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.SkipCellsLabel.AutoSize = true;
            this.SkipCellsLabel.Location = new System.Drawing.Point(3, 64);
            this.SkipCellsLabel.Name = "SkipCellsLabel";
            this.SkipCellsLabel.Size = new System.Drawing.Size(66, 16);
            this.SkipCellsLabel.TabIndex = 3;
            this.SkipCellsLabel.Text = "Skip cells";
            // 
            // SuperCopyMode
            // 
            this.SuperCopyMode.Dock = System.Windows.Forms.DockStyle.Fill;
            this.SuperCopyMode.FormattingEnabled = true;
            this.SuperCopyMode.Location = new System.Drawing.Point(97, 31);
            this.SuperCopyMode.Name = "SuperCopyMode";
            this.SuperCopyMode.Size = new System.Drawing.Size(88, 24);
            this.SuperCopyMode.TabIndex = 4;
            // 
            // SuperCopySkipCells
            // 
            this.SuperCopySkipCells.Dock = System.Windows.Forms.DockStyle.Fill;
            this.SuperCopySkipCells.Location = new System.Drawing.Point(97, 61);
            this.SuperCopySkipCells.Name = "SuperCopySkipCells";
            this.SuperCopySkipCells.Size = new System.Drawing.Size(88, 22);
            this.SuperCopySkipCells.TabIndex = 5;
            // 
            // SuperCopyGroupBox
            // 
            this.SuperCopyGroupBox.AutoSize = true;
            this.SuperCopyGroupBox.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.SuperCopyGroupBox.Controls.Add(this.SuperCopyTableLayout);
            this.SuperCopyGroupBox.Dock = System.Windows.Forms.DockStyle.Top;
            this.SuperCopyGroupBox.Location = new System.Drawing.Point(3, 3);
            this.SuperCopyGroupBox.Name = "SuperCopyGroupBox";
            this.SuperCopyGroupBox.Size = new System.Drawing.Size(194, 107);
            this.SuperCopyGroupBox.TabIndex = 2;
            this.SuperCopyGroupBox.TabStop = false;
            this.SuperCopyGroupBox.Text = "SuperCopy";
            // 
            // LoadDefault
            // 
            this.LoadDefault.AutoSize = true;
            this.LoadDefault.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.LoadDefault.Dock = System.Windows.Forms.DockStyle.Top;
            this.LoadDefault.Location = new System.Drawing.Point(3, 110);
            this.LoadDefault.Name = "LoadDefault";
            this.LoadDefault.Size = new System.Drawing.Size(194, 26);
            this.LoadDefault.TabIndex = 2;
            this.LoadDefault.Text = "Load Default";
            this.LoadDefault.UseVisualStyleBackColor = true;
            // 
            // SettingsUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Controls.Add(this.LoadDefault);
            this.Controls.Add(this.SuperCopyGroupBox);
            this.MinimumSize = new System.Drawing.Size(200, 0);
            this.Name = "SettingsUserControl";
            this.Padding = new System.Windows.Forms.Padding(3);
            this.Size = new System.Drawing.Size(200, 139);
            this.SuperCopyTableLayout.ResumeLayout(false);
            this.SuperCopyTableLayout.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.SuperCopySkipCells)).EndInit();
            this.SuperCopyGroupBox.ResumeLayout(false);
            this.SuperCopyGroupBox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox SuperCopyDelimiter;
        private System.Windows.Forms.Label DelimiterLabel;
        private System.Windows.Forms.TableLayoutPanel SuperCopyTableLayout;
        private System.Windows.Forms.GroupBox SuperCopyGroupBox;
        private System.Windows.Forms.Label ModeLabel;
        private System.Windows.Forms.Label SkipCellsLabel;
        private System.Windows.Forms.ComboBox SuperCopyMode;
        private System.Windows.Forms.NumericUpDown SuperCopySkipCells;
        private System.Windows.Forms.Button LoadDefault;
    }
}

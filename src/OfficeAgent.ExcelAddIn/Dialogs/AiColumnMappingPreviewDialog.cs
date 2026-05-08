using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OfficeAgent.Core.Models;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal sealed class AiColumnMappingPreviewDialog : Form
    {
        private readonly HostLocalizedStrings strings;

        private AiColumnMappingPreviewDialog(AiColumnMappingPreview preview, HostLocalizedStrings strings)
        {
            this.strings = strings ?? HostLocalizedStrings.ForLocale("en");

            Text = this.strings.AiColumnMappingPreviewDialogTitle;
            StartPosition = FormStartPosition.CenterParent;
            AutoScaleMode = AutoScaleMode.Font;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = false;
            ShowInTaskbar = false;
            MinimumSize = new Size(860, 420);
            ClientSize = new Size(960, 520);
            Padding = new Padding(16);

            var instructionLabel = new Label
            {
                AutoSize = true,
                Dock = DockStyle.Top,
                MaximumSize = new Size(900, 0),
                Text = this.strings.AiColumnMappingPreviewInstructionText,
            };

            var grid = CreateGrid(preview);

            var okButton = new Button
            {
                Text = this.strings.OkButtonText,
                DialogResult = DialogResult.OK,
                AutoSize = true,
                Padding = new Padding(12, 4, 12, 4),
                Margin = new Padding(8, 0, 0, 0),
            };
            var cancelButton = new Button
            {
                Text = this.strings.CancelButtonText,
                DialogResult = DialogResult.Cancel,
                AutoSize = true,
                Padding = new Padding(12, 4, 12, 4),
                Margin = new Padding(8, 0, 0, 0),
            };
            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                AutoSize = true,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(0, 12, 0, 0),
            };
            buttonPanel.Controls.Add(cancelButton);
            buttonPanel.Controls.Add(okButton);

            AcceptButton = okButton;
            CancelButton = cancelButton;

            Controls.Add(grid);
            Controls.Add(buttonPanel);
            Controls.Add(instructionLabel);
        }

        public static bool Confirm(AiColumnMappingPreview preview)
        {
            var strings = Globals.ThisAddIn?.HostLocalizedStrings ?? HostLocalizedStrings.ForLocale("en");
            using (var dialog = new AiColumnMappingPreviewDialog(preview, strings))
            {
                return dialog.ShowDialog() == DialogResult.OK;
            }
        }

        private static DataGridView CreateGrid(AiColumnMappingPreview preview)
        {
            var grid = new DataGridView
            {
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AutoGenerateColumns = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 12, 0, 0),
                MultiSelect = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            };

            AddTextColumn(grid, "ExcelColumn", "Excel column", 70, DataGridViewContentAlignment.MiddleRight);
            AddTextColumn(grid, "ActualHeader", "Actual header", 150, DataGridViewContentAlignment.MiddleLeft);
            AddTextColumn(grid, "IsdpHeader", "ISDP header", 150, DataGridViewContentAlignment.MiddleLeft);
            AddTextColumn(grid, "SuggestedHeader", "Suggested Excel header", 170, DataGridViewContentAlignment.MiddleLeft);
            AddTextColumn(grid, "Confidence", "Confidence", 80, DataGridViewContentAlignment.MiddleRight);
            AddTextColumn(grid, "Status", "Status", 90, DataGridViewContentAlignment.MiddleLeft);
            AddTextColumn(grid, "Reason", "Reason", 190, DataGridViewContentAlignment.MiddleLeft);

            foreach (var item in preview?.Items ?? Array.Empty<AiColumnMappingPreviewItem>())
            {
                if (item == null)
                {
                    continue;
                }

                grid.Rows.Add(
                    item.ExcelColumn.ToString(),
                    FormatHeader(item.SuggestedExcelL1, item.SuggestedExcelL2),
                    FormatHeader(item.TargetIsdpL1, item.TargetIsdpL2),
                    FormatHeader(item.SuggestedExcelL1, item.SuggestedExcelL2),
                    item.Confidence <= 0 ? string.Empty : item.Confidence.ToString("0.00"),
                    item.Status ?? string.Empty,
                    item.Reason ?? string.Empty);
            }

            if (grid.Rows.Count > 0)
            {
                grid.Rows[0].Selected = true;
            }

            return grid;
        }

        private static void AddTextColumn(
            DataGridView grid,
            string name,
            string headerText,
            int minimumWidth,
            DataGridViewContentAlignment alignment)
        {
            var column = new DataGridViewTextBoxColumn
            {
                Name = name,
                HeaderText = headerText,
                MinimumWidth = minimumWidth,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };
            column.DefaultCellStyle.Alignment = alignment;
            grid.Columns.Add(column);
        }

        private static string FormatHeader(string l1, string l2)
        {
            var parts = new[] { l1, l2 }
                .Where(part => !string.IsNullOrWhiteSpace(part))
                .Select(part => part.Trim())
                .ToArray();
            return string.Join(" / ", parts);
        }
    }
}

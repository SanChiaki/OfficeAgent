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
        private readonly AiColumnMappingPreview preview;
        private readonly DataGridView grid;
        private readonly AiColumnMappingPreviewItem[] displayedItems;

        private AiColumnMappingPreviewDialog(AiColumnMappingPreview preview, HostLocalizedStrings strings)
        {
            this.strings = strings ?? HostLocalizedStrings.ForLocale("en");
            this.preview = preview ?? new AiColumnMappingPreview();
            displayedItems = (this.preview.Items ?? Array.Empty<AiColumnMappingPreviewItem>())
                .Where(IsActionableItem)
                .ToArray();

            Text = this.strings.AiColumnMappingPreviewDialogTitle;
            StartPosition = FormStartPosition.CenterParent;
            AutoScaleMode = AutoScaleMode.Font;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = false;
            ShowInTaskbar = false;
            MinimumSize = new Size(720, 360);
            ClientSize = new Size(820, 460);
            Padding = new Padding(16);

            var instructionLabel = new Label
            {
                AutoSize = true,
                Dock = DockStyle.Top,
                MaximumSize = new Size(900, 0),
                Text = this.strings.AiColumnMappingPreviewInstructionText,
            };

            grid = CreateGrid();

            var okButton = new Button
            {
                Text = this.strings.OkButtonText,
                DialogResult = DialogResult.OK,
                AutoSize = true,
                Padding = new Padding(12, 4, 12, 4),
                Margin = new Padding(8, 0, 0, 0),
            };
            okButton.Click += (sender, args) => ApplySelectionToPreview();
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
            return Confirm(preview, null);
        }

        public static bool Confirm(AiColumnMappingPreview preview, IWin32Window owner)
        {
            var strings = Globals.ThisAddIn?.HostLocalizedStrings ?? HostLocalizedStrings.ForLocale("en");
            using (var dialog = new AiColumnMappingPreviewDialog(preview, strings))
            {
                return owner == null
                    ? dialog.ShowDialog() == DialogResult.OK
                    : dialog.ShowDialog(owner) == DialogResult.OK;
            }
        }

        private DataGridView CreateGrid()
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
                ReadOnly = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            };

            AddApplyColumn(grid);
            AddTextColumn(grid, "ExcelColumn", strings.AiColumnMappingExcelColumnHeader, 70, DataGridViewContentAlignment.MiddleLeft);
            AddTextColumn(grid, "ActualHeader", strings.AiColumnMappingActualHeaderColumnHeader, 220, DataGridViewContentAlignment.MiddleLeft);
            AddTextColumn(grid, "IsdpHeader", strings.AiColumnMappingMatchedHeaderColumnHeader, 220, DataGridViewContentAlignment.MiddleLeft);

            foreach (var item in displayedItems)
            {
                grid.Rows.Add(
                    item.ShouldApply,
                    FormatExcelColumnName(item.ExcelColumn),
                    FormatHeader(item.SuggestedExcelL1, item.SuggestedExcelL2),
                    FormatHeader(item.TargetIsdpL1, item.TargetIsdpL2));
            }

            if (grid.Rows.Count > 0)
            {
                grid.Rows[0].Selected = true;
            }

            return grid;
        }

        private void ApplySelectionToPreview()
        {
            grid.EndEdit();

            for (var index = 0; index < displayedItems.Length && index < grid.Rows.Count; index++)
            {
                var value = grid.Rows[index].Cells["ShouldApply"].Value;
                displayedItems[index].ShouldApply = value is bool shouldApply && shouldApply;
            }
        }

        private static bool IsActionableItem(AiColumnMappingPreviewItem item)
        {
            return item != null &&
                   string.Equals(item.Status, AiColumnMappingPreviewStatuses.Accepted, StringComparison.Ordinal);
        }

        private void AddApplyColumn(DataGridView grid)
        {
            grid.Columns.Add(new DataGridViewCheckBoxColumn
            {
                Name = "ShouldApply",
                HeaderText = strings.AiColumnMappingApplyColumnHeader,
                MinimumWidth = 80,
                FillWeight = 70,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            });
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
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };
            column.DefaultCellStyle.Alignment = alignment;
            grid.Columns.Add(column);
        }

        private static string FormatExcelColumnName(int columnNumber)
        {
            if (columnNumber <= 0)
            {
                return string.Empty;
            }

            var value = columnNumber;
            var result = string.Empty;
            while (value > 0)
            {
                value--;
                result = (char)('A' + (value % 26)) + result;
                value /= 26;
            }

            return result;
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

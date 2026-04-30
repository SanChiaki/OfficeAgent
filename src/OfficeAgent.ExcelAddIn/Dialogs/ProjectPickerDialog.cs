using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OfficeAgent.Core.Models;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal sealed class ProjectPickerDialog : Form
    {
        private readonly TextBox searchTextBox;
        private readonly ListBox projectListBox;
        private readonly IReadOnlyList<ProjectPickerItem> allItems;
        private readonly HostLocalizedStrings strings;

        public ProjectPickerDialog(IReadOnlyList<ProjectPickerItem> items, HostLocalizedStrings strings = null)
        {
            allItems = items ?? Array.Empty<ProjectPickerItem>();
            this.strings = strings ?? Globals.ThisAddIn?.HostLocalizedStrings ?? HostLocalizedStrings.ForLocale("en");

            Text = this.strings.ProjectPickerDialogTitle;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(620, 420);

            var root = new TableLayoutPanel
            {
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Padding = new Padding(16),
                RowCount = 4,
            };
            root.RowStyles.Add(new RowStyle());
            root.RowStyles.Add(new RowStyle());
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle());

            var instructionLabel = new Label
            {
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 8),
                MaximumSize = new Size(560, 0),
                Text = this.strings.ProjectPickerInstructionText,
            };

            searchTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 10),
            };
            searchTextBox.KeyDown += SearchTextBox_KeyDown;
            searchTextBox.TextChanged += SearchTextBox_TextChanged;

            projectListBox = new ListBox
            {
                Dock = DockStyle.Fill,
                HorizontalScrollbar = true,
                IntegralHeight = false,
                Margin = new Padding(0, 0, 0, 12),
            };
            projectListBox.DoubleClick += ProjectListBox_DoubleClick;
            projectListBox.KeyDown += ProjectListBox_KeyDown;

            var okButton = new Button { Text = this.strings.OkButtonText, Width = 88, Height = 30 };
            okButton.Click += OkButton_Click;
            var cancelButton = new Button { Text = this.strings.CancelButtonText, Width = 88, Height = 30, DialogResult = DialogResult.Cancel };

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 46,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(12, 4, 12, 4),
            };
            buttons.Controls.Add(cancelButton);
            buttons.Controls.Add(okButton);

            root.Controls.Add(instructionLabel, 0, 0);
            root.Controls.Add(searchTextBox, 0, 1);
            root.Controls.Add(projectListBox, 0, 2);
            root.Controls.Add(buttons, 0, 3);

            Controls.Add(root);
            AcceptButton = okButton;
            CancelButton = cancelButton;
            Shown += (sender, args) => searchTextBox.Focus();

            RebuildProjectList();
        }

        public ProjectOption SelectedProject { get; private set; }

        private void SearchTextBox_TextChanged(object sender, EventArgs e)
        {
            RebuildProjectList();
        }

        private void SearchTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Down || !projectListBox.Enabled || projectListBox.Items.Count == 0)
            {
                return;
            }

            projectListBox.Focus();
            e.Handled = true;
            e.SuppressKeyPress = true;
        }

        private void ProjectListBox_DoubleClick(object sender, EventArgs e)
        {
            ConfirmSelection();
        }

        private void ProjectListBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter)
            {
                return;
            }

            ConfirmSelection();
            e.Handled = true;
            e.SuppressKeyPress = true;
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            ConfirmSelection();
        }

        private void RebuildProjectList()
        {
            var selectedLabel = (projectListBox.SelectedItem as ProjectPickerListItem)?.Label ?? string.Empty;
            projectListBox.Items.Clear();

            foreach (var item in allItems.Where(item => ProjectSearchMatcher.IsMatch(item.Label, searchTextBox.Text)))
            {
                projectListBox.Items.Add(new ProjectPickerListItem(item));
            }

            if (projectListBox.Items.Count == 0)
            {
                projectListBox.Items.Add(new ProjectPickerListItem(new ProjectPickerItem(strings.ProjectPickerNoMatchesText, null)));
                projectListBox.Enabled = false;
                return;
            }

            projectListBox.Enabled = true;
            var selectedIndex = projectListBox.Items
                .Cast<ProjectPickerListItem>()
                .Select((item, index) => new { item, index })
                .FirstOrDefault(candidate => string.Equals(candidate.item.Label, selectedLabel, StringComparison.Ordinal))
                ?.index ?? 0;
            projectListBox.SelectedIndex = selectedIndex;
        }

        private void ConfirmSelection()
        {
            var selected = projectListBox.SelectedItem as ProjectPickerListItem;
            if (selected?.Project == null)
            {
                OperationResultDialog.ShowWarning(strings.ProjectPickerSelectionRequiredMessage);
                return;
            }

            SelectedProject = selected.Project;
            DialogResult = DialogResult.OK;
            Close();
        }

        internal sealed class ProjectPickerItem
        {
            public ProjectPickerItem(string label, ProjectOption project)
            {
                Label = label ?? string.Empty;
                Project = project;
            }

            public string Label { get; }

            public ProjectOption Project { get; }
        }

        private sealed class ProjectPickerListItem
        {
            private readonly ProjectPickerItem item;

            public ProjectPickerListItem(ProjectPickerItem item)
            {
                this.item = item;
            }

            public string Label => item.Label;

            public ProjectOption Project => item.Project;

            public override string ToString()
            {
                return Label;
            }
        }
    }
}

using System;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeAgent.Core.Models;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal sealed class InitializeSheetDialog : Form
    {
        private readonly InitializeSheetDialogRequest request;
        private readonly Func<InitializeSheetTemplateLoadResult> loadTemplates;
        private readonly HostLocalizedStrings strings;
        private readonly RadioButton templateModeRadioButton;
        private readonly RadioButton configOnlyModeRadioButton;
        private readonly ComboBox templateComboBox;
        private readonly Label templateStatusLabel;
        private readonly Label overwriteRiskLabel;
        private readonly Button confirmButton;
        private bool templatesLoaded;
        private bool templateModeAvailable;

        public InitializeSheetDialog(
            InitializeSheetDialogRequest request,
            Func<InitializeSheetTemplateLoadResult> loadTemplates,
            HostLocalizedStrings strings = null)
        {
            this.request = request ?? throw new ArgumentNullException(nameof(request));
            this.loadTemplates = loadTemplates ?? throw new ArgumentNullException(nameof(loadTemplates));
            this.strings = strings ?? Globals.ThisAddIn?.HostLocalizedStrings ?? HostLocalizedStrings.ForLocale("en");

            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Text = this.strings.InitializeSheetDialogTitle;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            Padding = new Padding(16);

            var instructionLabel = CreateWrappedLabel(this.strings.InitializeSheetInstructionText, new Padding(0));
            var projectLabel = CreateWrappedLabel(
                this.strings.InitializeSheetCurrentProjectText(this.request.ProjectDisplayName),
                new Padding(0, 8, 0, 0));

            templateModeRadioButton = new RadioButton
            {
                AutoSize = true,
                Margin = new Padding(0),
                Text = this.strings.InitializeSheetTemplateImportModeText,
            };
            templateModeRadioButton.CheckedChanged += (sender, args) => RefreshState();

            var templateDescriptionLabel = CreateWrappedLabel(
                this.strings.InitializeSheetTemplateImportDescription,
                new Padding(24, 4, 0, 0));

            var templateEditorPanel = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                Margin = new Padding(24, 10, 0, 0),
                Padding = new Padding(0),
            };
            templateEditorPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            templateEditorPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 338));

            var templateLabel = new Label
            {
                AutoSize = true,
                Margin = new Padding(0, 4, 8, 0),
                Text = this.strings.InitializeSheetTemplateLabel,
            };

            templateComboBox = new ComboBox
            {
                DisplayMember = nameof(TemplateListItem.DisplayText),
                ValueMember = nameof(TemplateListItem.TemplateId),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Enabled = false,
                FormattingEnabled = true,
                Margin = new Padding(0),
                Width = 338,
            };
            templateEditorPanel.Controls.Add(templateLabel, 0, 0);
            templateEditorPanel.Controls.Add(templateComboBox, 1, 0);

            templateStatusLabel = CreateWrappedLabel(
                this.strings.InitializeSheetTemplateLoadingText,
                new Padding(24, 8, 0, 0));

            overwriteRiskLabel = CreateWrappedLabel(
                this.strings.InitializeSheetOverwriteRiskMessage,
                new Padding(24, 10, 0, 0));
            overwriteRiskLabel.ForeColor = Color.FromArgb(156, 73, 0);

            configOnlyModeRadioButton = new RadioButton
            {
                AutoSize = true,
                Margin = new Padding(0, 18, 0, 0),
                Text = this.strings.InitializeSheetConfigOnlyModeText,
            };
            configOnlyModeRadioButton.CheckedChanged += (sender, args) => RefreshState();

            var configOnlyDescriptionLabel = CreateWrappedLabel(
                this.strings.InitializeSheetConfigOnlyDescription,
                new Padding(24, 4, 0, 0));

            confirmButton = new Button
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                DialogResult = DialogResult.None,
                Margin = new Padding(8, 0, 0, 0),
                Padding = new Padding(12, 4, 12, 4),
                Text = this.strings.InitializeSheetConfirmButtonText,
            };
            confirmButton.Click += HandleConfirmClick;

            var cancelButton = new Button
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                DialogResult = DialogResult.Cancel,
                Margin = new Padding(8, 0, 0, 0),
                Padding = new Padding(12, 4, 12, 4),
                Text = this.strings.CancelButtonText,
            };

            var buttonsPanel = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                FlowDirection = FlowDirection.RightToLeft,
                Margin = new Padding(0, 18, 0, 0),
                Padding = new Padding(0),
                WrapContents = false,
            };
            buttonsPanel.Controls.Add(cancelButton);
            buttonsPanel.Controls.Add(confirmButton);

            var contentPanel = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                Padding = new Padding(0),
            };
            contentPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            contentPanel.Controls.Add(instructionLabel);
            contentPanel.Controls.Add(projectLabel);
            contentPanel.Controls.Add(templateModeRadioButton);
            contentPanel.Controls.Add(templateDescriptionLabel);
            contentPanel.Controls.Add(templateEditorPanel);
            contentPanel.Controls.Add(templateStatusLabel);
            contentPanel.Controls.Add(overwriteRiskLabel);
            contentPanel.Controls.Add(configOnlyModeRadioButton);
            contentPanel.Controls.Add(configOnlyDescriptionLabel);
            contentPanel.Controls.Add(buttonsPanel);

            AcceptButton = confirmButton;
            CancelButton = cancelButton;
            Controls.Add(contentPanel);

            templateStatusLabel.Text = this.strings.InitializeSheetTemplateLoadingText;
            confirmButton.Enabled = false;
            configOnlyModeRadioButton.Checked = true;
            RefreshState();
        }

        public InitializeSheetDialogResult Result { get; private set; }

        internal static InitializeSheetMode ResolveDefaultMode(bool isBlankSheet, bool canImportTemplate)
        {
            return isBlankSheet && canImportTemplate
                ? InitializeSheetMode.TemplateImport
                : InitializeSheetMode.ConfigOnly;
        }

        protected override async void OnShown(EventArgs e)
        {
            base.OnShown(e);
            await LoadTemplatesAsync();
        }

        private async Task LoadTemplatesAsync()
        {
            templateStatusLabel.Text = strings.InitializeSheetTemplateLoadingText;
            confirmButton.Enabled = false;
            Refresh();

            InitializeSheetTemplateLoadResult loadResult;
            try
            {
                loadResult = request.SupportsTemplateImport
                    ? await Task.Run(loadTemplates)
                    : InitializeSheetTemplateLoadResult.Unsupported(strings.InitializeSheetTemplateUnsupportedMessage);
            }
            catch
            {
                loadResult = InitializeSheetTemplateLoadResult.Failed(strings.InitializeSheetTemplateLoadFailedMessage);
            }

            if (IsDisposed)
            {
                return;
            }

            ApplyTemplateLoadResult(loadResult);
        }

        private void ApplyTemplateLoadResult(InitializeSheetTemplateLoadResult loadResult)
        {
            var normalizedResult = loadResult ?? InitializeSheetTemplateLoadResult.Failed(strings.InitializeSheetTemplateLoadFailedMessage);
            var templates = normalizedResult.Templates
                .Where(template => template != null && !string.IsNullOrWhiteSpace(template.TemplateId))
                .Select(template => new TemplateListItem(template))
                .ToArray();

            templateComboBox.Items.Clear();
            foreach (var template in templates)
            {
                templateComboBox.Items.Add(template);
            }

            if (templateComboBox.Items.Count > 0)
            {
                templateComboBox.SelectedIndex = 0;
            }

            templateModeAvailable = normalizedResult.IsSupported &&
                normalizedResult.IsSuccess &&
                templateComboBox.Items.Count > 0;
            templatesLoaded = true;

            if (templateModeAvailable)
            {
                templateStatusLabel.Text = string.Empty;
            }
            else if (!string.IsNullOrWhiteSpace(normalizedResult.DisabledReason))
            {
                templateStatusLabel.Text = normalizedResult.DisabledReason;
            }
            else
            {
                templateStatusLabel.Text = strings.InitializeSheetTemplateEmptyMessage;
            }

            var defaultMode = ResolveDefaultMode(request.IsBlankSheet, templateModeAvailable);
            templateModeRadioButton.Checked = defaultMode == InitializeSheetMode.TemplateImport;
            configOnlyModeRadioButton.Checked = defaultMode == InitializeSheetMode.ConfigOnly;
            RefreshState();
        }

        private void RefreshState()
        {
            templateModeRadioButton.Enabled = templatesLoaded && templateModeAvailable;
            templateComboBox.Enabled = templateModeRadioButton.Checked && templateModeAvailable;
            templateStatusLabel.Visible = !templateModeAvailable || !string.IsNullOrWhiteSpace(templateStatusLabel.Text);
            overwriteRiskLabel.Visible = templateModeRadioButton.Checked && !request.IsBlankSheet;
            confirmButton.Enabled = templatesLoaded &&
                (configOnlyModeRadioButton.Checked || (templateModeRadioButton.Checked && templateModeAvailable && templateComboBox.SelectedItem != null));
            confirmButton.Text = overwriteRiskLabel.Visible
                ? strings.InitializeSheetOverwriteButtonText
                : strings.InitializeSheetConfirmButtonText;
        }

        private void HandleConfirmClick(object sender, EventArgs e)
        {
            var mode = templateModeRadioButton.Checked
                ? InitializeSheetMode.TemplateImport
                : InitializeSheetMode.ConfigOnly;
            var selectedTemplate = mode == InitializeSheetMode.TemplateImport
                ? (templateComboBox.SelectedItem as TemplateListItem)?.Template
                : null;

            Result = new InitializeSheetDialogResult
            {
                Mode = mode,
                SelectedTemplate = selectedTemplate,
            };
            DialogResult = DialogResult.OK;
            Close();
        }

        private static Label CreateWrappedLabel(string text, Padding margin)
        {
            return new Label
            {
                AutoSize = true,
                Margin = margin,
                MaximumSize = new Size(520, 0),
                Text = text ?? string.Empty,
            };
        }

        private sealed class TemplateListItem
        {
            public TemplateListItem(BusinessExportTemplateOption template)
            {
                Template = template ?? throw new ArgumentNullException(nameof(template));
                DisplayText = string.IsNullOrWhiteSpace(template.TemplateName)
                    ? template.TemplateId
                    : template.TemplateName;
            }

            public BusinessExportTemplateOption Template { get; }

            public string TemplateId => Template.TemplateId;

            public string DisplayText { get; }

            public override string ToString()
            {
                return DisplayText;
            }
        }
    }
}

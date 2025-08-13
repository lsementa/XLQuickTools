namespace XLQuickTools
{
    partial class XLQuickTools : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public XLQuickTools()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(XLQuickTools));
            this.XLQuickTools_Tab = this.Factory.CreateRibbonTab();
            this.Group_Formatting = this.Factory.CreateRibbonGroup();
            this.Separator_Formatting = this.Factory.CreateRibbonSeparator();
            this.Group_Data = this.Factory.CreateRibbonGroup();
            this.Separator_Data = this.Factory.CreateRibbonSeparator();
            this.Group_Delimiter = this.Factory.CreateRibbonGroup();
            this.Separator_Delimiter = this.Factory.CreateRibbonSeparator();
            this.Group_Hyperlinks = this.Factory.CreateRibbonGroup();
            this.Group_About = this.Factory.CreateRibbonGroup();
            this.BtnRemoveExcess = this.Factory.CreateRibbonSplitButton();
            this.BtnRemoveExcessWS = this.Factory.CreateRibbonButton();
            this.BtnRemoveExcessWB = this.Factory.CreateRibbonButton();
            this.SBtnTrimClean = this.Factory.CreateRibbonSplitButton();
            this.BtnTrimClean = this.Factory.CreateRibbonButton();
            this.BtnTrimCleanWorksheet = this.Factory.CreateRibbonButton();
            this.BtnTrimCleanWorkbook = this.Factory.CreateRibbonButton();
            this.BtnTrimCleanSettings = this.Factory.CreateRibbonButton();
            this.BtnQuickSettings = this.Factory.CreateRibbonButton();
            this.BtnQuickFormat = this.Factory.CreateRibbonSplitButton();
            this.BtnQuickFormatSub = this.Factory.CreateRibbonButton();
            this.BtnQuickFormatAll = this.Factory.CreateRibbonButton();
            this.BtnRemoveFormatting = this.Factory.CreateRibbonSplitButton();
            this.BtnRemoveFormattingSub = this.Factory.CreateRibbonButton();
            this.BtnRemoveFormattingAll = this.Factory.CreateRibbonButton();
            this.separator8 = this.Factory.CreateRibbonSeparator();
            this.BtnConvertTableRemove = this.Factory.CreateRibbonButton();
            this.BtnRemoveObjects = this.Factory.CreateRibbonButton();
            this.BtnDateText = this.Factory.CreateRibbonButton();
            this.BtnUndo = this.Factory.CreateRibbonButton();
            this.CharacterMenu = this.Factory.CreateRibbonMenu();
            this.BtnUppercase = this.Factory.CreateRibbonButton();
            this.BtnLowercase = this.Factory.CreateRibbonButton();
            this.BtnPropercase = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.BtnAddLeadingTrailing = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.BtnRemoveSpaces = this.Factory.CreateRibbonButton();
            this.BtnRemoveLetters = this.Factory.CreateRibbonButton();
            this.BtnRemoveNumbers = this.Factory.CreateRibbonButton();
            this.BtnRemoveSpecial = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.BtnNormalizeText = this.Factory.CreateRibbonButton();
            this.BtnRemoveNonASCII = this.Factory.CreateRibbonButton();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.BtnSubscript = this.Factory.CreateRibbonButton();
            this.AdditionalMenu = this.Factory.CreateRibbonMenu();
            this.BtnDeleteRows = this.Factory.CreateRibbonButton();
            this.BtnDeleteColumns = this.Factory.CreateRibbonButton();
            this.BtnFillDown = this.Factory.CreateRibbonButton();
            this.BtnResetColumn = this.Factory.CreateRibbonButton();
            this.separator5 = this.Factory.CreateRibbonSeparator();
            this.BtnCopyFormatting = this.Factory.CreateRibbonButton();
            this.BtnCopyHighlightedValues = this.Factory.CreateRibbonButton();
            this.BtnCopyHighlightedRows = this.Factory.CreateRibbonButton();
            this.separator6 = this.Factory.CreateRibbonSeparator();
            this.BtnDisplayLength = this.Factory.CreateRibbonButton();
            this.BtnSheetNames = this.Factory.CreateRibbonButton();
            this.BtnFileList = this.Factory.CreateRibbonButton();
            this.BtnMissing = this.Factory.CreateRibbonButton();
            this.BtnDuplicates = this.Factory.CreateRibbonButton();
            this.BtnCompare = this.Factory.CreateRibbonButton();
            this.BtnFilter = this.Factory.CreateRibbonButton();
            this.SBtnUniqueClipboard = this.Factory.CreateRibbonSplitButton();
            this.BtnUniqueClipboard = this.Factory.CreateRibbonButton();
            this.BtnCopyToSheets = this.Factory.CreateRibbonButton();
            this.BtnColumnInfo = this.Factory.CreateRibbonButton();
            this.BtnCommaSelection = this.Factory.CreateRibbonButton();
            this.BtnDelimSelection = this.Factory.CreateRibbonButton();
            this.BtnSheetToFile = this.Factory.CreateRibbonButton();
            this.BtnSplitToRows = this.Factory.CreateRibbonButton();
            this.BtnHyperlinkSettings = this.Factory.CreateRibbonButton();
            this.BtnHyperlinks = this.Factory.CreateRibbonSplitButton();
            this.BtnAddHyperlinks = this.Factory.CreateRibbonButton();
            this.BtnAddHyperlinksCell = this.Factory.CreateRibbonButton();
            this.BtnRemoveHyperlinks = this.Factory.CreateRibbonButton();
            this.BtnAbout = this.Factory.CreateRibbonButton();
            this.XLQuickTools_Tab.SuspendLayout();
            this.Group_Formatting.SuspendLayout();
            this.Group_Data.SuspendLayout();
            this.Group_Delimiter.SuspendLayout();
            this.Group_Hyperlinks.SuspendLayout();
            this.Group_About.SuspendLayout();
            this.SuspendLayout();
            // 
            // XLQuickTools_Tab
            // 
            this.XLQuickTools_Tab.Groups.Add(this.Group_Formatting);
            this.XLQuickTools_Tab.Groups.Add(this.Group_Data);
            this.XLQuickTools_Tab.Groups.Add(this.Group_Delimiter);
            this.XLQuickTools_Tab.Groups.Add(this.Group_Hyperlinks);
            this.XLQuickTools_Tab.Groups.Add(this.Group_About);
            this.XLQuickTools_Tab.Label = "Quick Tools";
            this.XLQuickTools_Tab.Name = "XLQuickTools_Tab";
            this.XLQuickTools_Tab.Position = this.Factory.RibbonPosition.BeforeOfficeId("HelpTab");
            // 
            // Group_Formatting
            // 
            this.Group_Formatting.Items.Add(this.BtnRemoveExcess);
            this.Group_Formatting.Items.Add(this.SBtnTrimClean);
            this.Group_Formatting.Items.Add(this.BtnQuickSettings);
            this.Group_Formatting.Items.Add(this.Separator_Formatting);
            this.Group_Formatting.Items.Add(this.BtnQuickFormat);
            this.Group_Formatting.Items.Add(this.BtnRemoveFormatting);
            this.Group_Formatting.Items.Add(this.BtnDateText);
            this.Group_Formatting.Items.Add(this.BtnUndo);
            this.Group_Formatting.Items.Add(this.CharacterMenu);
            this.Group_Formatting.Items.Add(this.AdditionalMenu);
            this.Group_Formatting.Label = "Formatting";
            this.Group_Formatting.Name = "Group_Formatting";
            // 
            // Separator_Formatting
            // 
            this.Separator_Formatting.Name = "Separator_Formatting";
            // 
            // Group_Data
            // 
            this.Group_Data.Items.Add(this.BtnMissing);
            this.Group_Data.Items.Add(this.BtnDuplicates);
            this.Group_Data.Items.Add(this.BtnCompare);
            this.Group_Data.Items.Add(this.Separator_Data);
            this.Group_Data.Items.Add(this.BtnFilter);
            this.Group_Data.Items.Add(this.SBtnUniqueClipboard);
            this.Group_Data.Label = "Data";
            this.Group_Data.Name = "Group_Data";
            // 
            // Separator_Data
            // 
            this.Separator_Data.Name = "Separator_Data";
            // 
            // Group_Delimiter
            // 
            this.Group_Delimiter.Items.Add(this.BtnCommaSelection);
            this.Group_Delimiter.Items.Add(this.BtnDelimSelection);
            this.Group_Delimiter.Items.Add(this.BtnSheetToFile);
            this.Group_Delimiter.Items.Add(this.Separator_Delimiter);
            this.Group_Delimiter.Items.Add(this.BtnSplitToRows);
            this.Group_Delimiter.Label = "Delimiter";
            this.Group_Delimiter.Name = "Group_Delimiter";
            // 
            // Separator_Delimiter
            // 
            this.Separator_Delimiter.Name = "Separator_Delimiter";
            // 
            // Group_Hyperlinks
            // 
            this.Group_Hyperlinks.Items.Add(this.BtnHyperlinkSettings);
            this.Group_Hyperlinks.Items.Add(this.BtnHyperlinks);
            this.Group_Hyperlinks.Label = "Hyperlinks";
            this.Group_Hyperlinks.Name = "Group_Hyperlinks";
            // 
            // Group_About
            // 
            this.Group_About.Items.Add(this.BtnAbout);
            this.Group_About.Label = "Help";
            this.Group_About.Name = "Group_About";
            // 
            // BtnRemoveExcess
            // 
            this.BtnRemoveExcess.Items.Add(this.BtnRemoveExcessWS);
            this.BtnRemoveExcess.Items.Add(this.BtnRemoveExcessWB);
            this.BtnRemoveExcess.Label = "Remove Excess";
            this.BtnRemoveExcess.Name = "BtnRemoveExcess";
            this.BtnRemoveExcess.OfficeImageId = "TableStyleClear";
            this.BtnRemoveExcess.SuperTip = "Remove Excess Formatting\n\nRemove any excess formatting from either the active wor" +
    "ksheet or active workbook.";
            this.BtnRemoveExcess.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnRemoveExcess_Click);
            // 
            // BtnRemoveExcessWS
            // 
            this.BtnRemoveExcessWS.Label = "Active Work&sheet";
            this.BtnRemoveExcessWS.Name = "BtnRemoveExcessWS";
            this.BtnRemoveExcessWS.OfficeImageId = "SelectSheet";
            this.BtnRemoveExcessWS.ShowImage = true;
            this.BtnRemoveExcessWS.SuperTip = "Remove excess formatting on the active worksheet.";
            this.BtnRemoveExcessWS.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnRemoveExcessWS_Click);
            // 
            // BtnRemoveExcessWB
            // 
            this.BtnRemoveExcessWB.Label = "Active Work&book";
            this.BtnRemoveExcessWB.Name = "BtnRemoveExcessWB";
            this.BtnRemoveExcessWB.OfficeImageId = "MicrosoftExcel";
            this.BtnRemoveExcessWB.ShowImage = true;
            this.BtnRemoveExcessWB.SuperTip = "Remove excess formatting on all worksheets.";
            this.BtnRemoveExcessWB.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnRemoveExcessWB_Click);
            // 
            // SBtnTrimClean
            // 
            this.SBtnTrimClean.Items.Add(this.BtnTrimClean);
            this.SBtnTrimClean.Items.Add(this.BtnTrimCleanWorksheet);
            this.SBtnTrimClean.Items.Add(this.BtnTrimCleanWorkbook);
            this.SBtnTrimClean.Items.Add(this.BtnTrimCleanSettings);
            this.SBtnTrimClean.Label = "Trim and Clean";
            this.SBtnTrimClean.Name = "SBtnTrimClean";
            this.SBtnTrimClean.OfficeImageId = "TableSelectCell";
            this.SBtnTrimClean.SuperTip = "Trim and Clean\n\nRemoves leading, trailing, multiple spaces and non-printable char" +
    "acters from the selected range.";
            this.SBtnTrimClean.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnTrimClean_Click_1);
            // 
            // BtnTrimClean
            // 
            this.BtnTrimClean.Label = "Selected &Range";
            this.BtnTrimClean.Name = "BtnTrimClean";
            this.BtnTrimClean.OfficeImageId = "TableSelectCell";
            this.BtnTrimClean.ShowImage = true;
            this.BtnTrimClean.SuperTip = "Apply Trim & Clean to the selected range.";
            this.BtnTrimClean.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnTrimClean_Click_2);
            // 
            // BtnTrimCleanWorksheet
            // 
            this.BtnTrimCleanWorksheet.Label = "Active Work&sheet";
            this.BtnTrimCleanWorksheet.Name = "BtnTrimCleanWorksheet";
            this.BtnTrimCleanWorksheet.OfficeImageId = "SelectSheet";
            this.BtnTrimCleanWorksheet.ShowImage = true;
            this.BtnTrimCleanWorksheet.SuperTip = "Apply Trim & Clean to active worksheet.";
            this.BtnTrimCleanWorksheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnTrimCleanWorksheet_Click);
            // 
            // BtnTrimCleanWorkbook
            // 
            this.BtnTrimCleanWorkbook.Label = "Active Work&book";
            this.BtnTrimCleanWorkbook.Name = "BtnTrimCleanWorkbook";
            this.BtnTrimCleanWorkbook.OfficeImageId = "MicrosoftExcel";
            this.BtnTrimCleanWorkbook.ShowImage = true;
            this.BtnTrimCleanWorkbook.SuperTip = "Apply Trim & Clean to all worksheets.";
            this.BtnTrimCleanWorkbook.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnTrimCleanWorkbook_Click);
            // 
            // BtnTrimCleanSettings
            // 
            this.BtnTrimCleanSettings.Label = "&Clean Settings";
            this.BtnTrimCleanSettings.Name = "BtnTrimCleanSettings";
            this.BtnTrimCleanSettings.OfficeImageId = "ColumnSettingsMenu";
            this.BtnTrimCleanSettings.ShowImage = true;
            this.BtnTrimCleanSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnTrimCleanSettings_Click);
            // 
            // BtnQuickSettings
            // 
            this.BtnQuickSettings.Label = "Quick Settings";
            this.BtnQuickSettings.Name = "BtnQuickSettings";
            this.BtnQuickSettings.OfficeImageId = "ManageQuickSteps";
            this.BtnQuickSettings.ShowImage = true;
            this.BtnQuickSettings.SuperTip = "Quick Format Settings\n\nSettings you want applied when using the \"Quick Format\" bu" +
    "tton.";
            this.BtnQuickSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnQuickSettings_Click);
            // 
            // BtnQuickFormat
            // 
            this.BtnQuickFormat.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnQuickFormat.Items.Add(this.BtnQuickFormatSub);
            this.BtnQuickFormat.Items.Add(this.BtnQuickFormatAll);
            this.BtnQuickFormat.Label = "Quick Format";
            this.BtnQuickFormat.Name = "BtnQuickFormat";
            this.BtnQuickFormat.OfficeImageId = "AutoFormatChange";
            this.BtnQuickFormat.SuperTip = "Quick Format\n\nApply your saved formatting.";
            this.BtnQuickFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnQuickFormat_Click);
            // 
            // BtnQuickFormatSub
            // 
            this.BtnQuickFormatSub.Label = "&Quick Format";
            this.BtnQuickFormatSub.Name = "BtnQuickFormatSub";
            this.BtnQuickFormatSub.OfficeImageId = "AutoFormatChange";
            this.BtnQuickFormatSub.ShowImage = true;
            this.BtnQuickFormatSub.SuperTip = "Apply Quick Format to active worksheet.";
            this.BtnQuickFormatSub.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnQuickFormatSub_Click);
            // 
            // BtnQuickFormatAll
            // 
            this.BtnQuickFormatAll.Label = "&All Worksheets";
            this.BtnQuickFormatAll.Name = "BtnQuickFormatAll";
            this.BtnQuickFormatAll.OfficeImageId = "AutoFormatDialog";
            this.BtnQuickFormatAll.ShowImage = true;
            this.BtnQuickFormatAll.SuperTip = "Apply Quick Format to all worksheets.";
            this.BtnQuickFormatAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnQuickFormatAll_Click);
            // 
            // BtnRemoveFormatting
            // 
            this.BtnRemoveFormatting.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnRemoveFormatting.Items.Add(this.BtnRemoveFormattingSub);
            this.BtnRemoveFormatting.Items.Add(this.BtnRemoveFormattingAll);
            this.BtnRemoveFormatting.Items.Add(this.separator8);
            this.BtnRemoveFormatting.Items.Add(this.BtnConvertTableRemove);
            this.BtnRemoveFormatting.Items.Add(this.BtnRemoveObjects);
            this.BtnRemoveFormatting.Label = "Remove";
            this.BtnRemoveFormatting.Name = "BtnRemoveFormatting";
            this.BtnRemoveFormatting.OfficeImageId = "Clear";
            this.BtnRemoveFormatting.SuperTip = "Remove Formatting\n\nRemove formatting from the active worksheet.";
            this.BtnRemoveFormatting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnRemoveFormatting_Click);
            // 
            // BtnRemoveFormattingSub
            // 
            this.BtnRemoveFormattingSub.Label = "&Remove Formatting";
            this.BtnRemoveFormattingSub.Name = "BtnRemoveFormattingSub";
            this.BtnRemoveFormattingSub.OfficeImageId = "Clear";
            this.BtnRemoveFormattingSub.ShowImage = true;
            this.BtnRemoveFormattingSub.SuperTip = "Remove formatting from the active worksheet.";
            this.BtnRemoveFormattingSub.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnRemoveFormattingSub_Click);
            // 
            // BtnRemoveFormattingAll
            // 
            this.BtnRemoveFormattingAll.Label = "&All Worksheets";
            this.BtnRemoveFormattingAll.Name = "BtnRemoveFormattingAll";
            this.BtnRemoveFormattingAll.OfficeImageId = "TableStyleClear";
            this.BtnRemoveFormattingAll.ShowImage = true;
            this.BtnRemoveFormattingAll.SuperTip = "Remove formatting from all worksheets.";
            this.BtnRemoveFormattingAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnRemoveFormattingAll_Click);
            // 
            // separator8
            // 
            this.separator8.Name = "separator8";
            // 
            // BtnConvertTableRemove
            // 
            this.BtnConvertTableRemove.Label = "Remove &Table Formatting";
            this.BtnConvertTableRemove.Name = "BtnConvertTableRemove";
            this.BtnConvertTableRemove.OfficeImageId = "TableEraser";
            this.BtnConvertTableRemove.ShowImage = true;
            this.BtnConvertTableRemove.SuperTip = "Convert tables to ranges and remove formatting.";
            this.BtnConvertTableRemove.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnConvertTableRemove_Click);
            // 
            // BtnRemoveObjects
            // 
            this.BtnRemoveObjects.Label = "Remove &Objects";
            this.BtnRemoveObjects.Name = "BtnRemoveObjects";
            this.BtnRemoveObjects.OfficeImageId = "ShapesMoreShapes";
            this.BtnRemoveObjects.ShowImage = true;
            this.BtnRemoveObjects.SuperTip = "Remove objects (shapes, pictures, charts, activeX, etc.)";
            this.BtnRemoveObjects.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnRemoveObjects_Click);
            // 
            // BtnDateText
            // 
            this.BtnDateText.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnDateText.Label = "Date/Text Converter";
            this.BtnDateText.Name = "BtnDateText";
            this.BtnDateText.OfficeImageId = "DateInsert";
            this.BtnDateText.ShowImage = true;
            this.BtnDateText.SuperTip = "Date/Text Converter\n\nConvert the selected range to a desired text or Excel format" +
    ".";
            this.BtnDateText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnDateText_Click);
            // 
            // BtnUndo
            // 
            this.BtnUndo.Enabled = false;
            this.BtnUndo.Label = "Undo";
            this.BtnUndo.Name = "BtnUndo";
            this.BtnUndo.OfficeImageId = "Undo";
            this.BtnUndo.ShowImage = true;
            this.BtnUndo.ShowLabel = false;
            this.BtnUndo.SuperTip = "Undo Date/Text, Character or Fill Down";
            this.BtnUndo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnUndo_Click);
            // 
            // CharacterMenu
            // 
            this.CharacterMenu.Items.Add(this.BtnUppercase);
            this.CharacterMenu.Items.Add(this.BtnLowercase);
            this.CharacterMenu.Items.Add(this.BtnPropercase);
            this.CharacterMenu.Items.Add(this.separator1);
            this.CharacterMenu.Items.Add(this.BtnAddLeadingTrailing);
            this.CharacterMenu.Items.Add(this.separator2);
            this.CharacterMenu.Items.Add(this.BtnRemoveSpaces);
            this.CharacterMenu.Items.Add(this.BtnRemoveLetters);
            this.CharacterMenu.Items.Add(this.BtnRemoveNumbers);
            this.CharacterMenu.Items.Add(this.BtnRemoveSpecial);
            this.CharacterMenu.Items.Add(this.separator3);
            this.CharacterMenu.Items.Add(this.BtnNormalizeText);
            this.CharacterMenu.Items.Add(this.BtnRemoveNonASCII);
            this.CharacterMenu.Items.Add(this.separator4);
            this.CharacterMenu.Items.Add(this.BtnSubscript);
            this.CharacterMenu.Label = "Character Menu";
            this.CharacterMenu.Name = "CharacterMenu";
            this.CharacterMenu.OfficeImageId = "CharacterBorder";
            this.CharacterMenu.ShowImage = true;
            this.CharacterMenu.ShowLabel = false;
            this.CharacterMenu.SuperTip = "Text Options";
            // 
            // BtnUppercase
            // 
            this.BtnUppercase.Label = "&Upper Case";
            this.BtnUppercase.Name = "BtnUppercase";
            this.BtnUppercase.OfficeImageId = "MessagePrevious";
            this.BtnUppercase.ShowImage = true;
            this.BtnUppercase.SuperTip = "Set the selected range to UPPERCASE.";
            this.BtnUppercase.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnUppercase_Click);
            // 
            // BtnLowercase
            // 
            this.BtnLowercase.Label = "&Lower Case";
            this.BtnLowercase.Name = "BtnLowercase";
            this.BtnLowercase.OfficeImageId = "MessageNext";
            this.BtnLowercase.ShowImage = true;
            this.BtnLowercase.SuperTip = "Set the selected range to lowercase.";
            this.BtnLowercase.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnLowercase_Click);
            // 
            // BtnPropercase
            // 
            this.BtnPropercase.Label = "&Proper Case";
            this.BtnPropercase.Name = "BtnPropercase";
            this.BtnPropercase.OfficeImageId = "SchemeFontsGallery";
            this.BtnPropercase.ShowImage = true;
            this.BtnPropercase.SuperTip = "Set the selected range to Proper Case.";
            this.BtnPropercase.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnPropercase_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // BtnAddLeadingTrailing
            // 
            this.BtnAddLeadingTrailing.Label = "&Add Leading or Trailing";
            this.BtnAddLeadingTrailing.Name = "BtnAddLeadingTrailing";
            this.BtnAddLeadingTrailing.OfficeImageId = "FormControlEditBox";
            this.BtnAddLeadingTrailing.ShowImage = true;
            this.BtnAddLeadingTrailing.SuperTip = "Add leading or trailing characters/text.";
            this.BtnAddLeadingTrailing.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnAddLeadingTrailing_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // BtnRemoveSpaces
            // 
            this.BtnRemoveSpaces.Label = "Remove E&xtra Spaces";
            this.BtnRemoveSpaces.Name = "BtnRemoveSpaces";
            this.BtnRemoveSpaces.OfficeImageId = "EndOfLine";
            this.BtnRemoveSpaces.ShowImage = true;
            this.BtnRemoveSpaces.SuperTip = "Remove two or more consecutive whitespace characters.";
            this.BtnRemoveSpaces.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnRemoveSpaces_Click);
            // 
            // BtnRemoveLetters
            // 
            this.BtnRemoveLetters.Label = "Remove L&etters [A-Z]";
            this.BtnRemoveLetters.Name = "BtnRemoveLetters";
            this.BtnRemoveLetters.OfficeImageId = "TextSmallCaps";
            this.BtnRemoveLetters.ShowImage = true;
            this.BtnRemoveLetters.SuperTip = "Remove letters from a string in the selected range.";
            this.BtnRemoveLetters.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnRemoveLetters_Click);
            // 
            // BtnRemoveNumbers
            // 
            this.BtnRemoveNumbers.Label = "Remove &Numbers [0-9]";
            this.BtnRemoveNumbers.Name = "BtnRemoveNumbers";
            this.BtnRemoveNumbers.OfficeImageId = "NumberStyleGallery";
            this.BtnRemoveNumbers.ShowImage = true;
            this.BtnRemoveNumbers.SuperTip = "Remove numbers from a string in the selected range.";
            this.BtnRemoveNumbers.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnRemoveNumbers_Click);
            // 
            // BtnRemoveSpecial
            // 
            this.BtnRemoveSpecial.Label = "Remove &Special [~!@#$%^]";
            this.BtnRemoveSpecial.Name = "BtnRemoveSpecial";
            this.BtnRemoveSpecial.OfficeImageId = "FieldCodes";
            this.BtnRemoveSpecial.ShowImage = true;
            this.BtnRemoveSpecial.SuperTip = "Remove special characters from a string in the selected range.";
            this.BtnRemoveSpecial.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnRemoveSpecial_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // BtnNormalizeText
            // 
            this.BtnNormalizeText.Label = "Normalize &Text [é > e]";
            this.BtnNormalizeText.Name = "BtnNormalizeText";
            this.BtnNormalizeText.OfficeImageId = "EncodingMenu";
            this.BtnNormalizeText.ShowImage = true;
            this.BtnNormalizeText.SuperTip = "Normalize text/replace diacritics within the selected range.";
            this.BtnNormalizeText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnNormalizeText_Click);
            // 
            // BtnRemoveNonASCII
            // 
            this.BtnRemoveNonASCII.Label = "Replace N&on-ASCII [é > &&#233;]";
            this.BtnRemoveNonASCII.Name = "BtnRemoveNonASCII";
            this.BtnRemoveNonASCII.OfficeImageId = "JapaneseGreetingsInsertMenu";
            this.BtnRemoveNonASCII.ShowImage = true;
            this.BtnRemoveNonASCII.SuperTip = "Replace characters that fall outside the 128-character ASCII set with the HTML en" +
    "tity code.";
            this.BtnRemoveNonASCII.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnRemoveNonASCII_Click);
            // 
            // separator4
            // 
            this.separator4.Name = "separator4";
            // 
            // BtnSubscript
            // 
            this.BtnSubscript.Label = "Subscript &Chemical Formulas";
            this.BtnSubscript.Name = "BtnSubscript";
            this.BtnSubscript.OfficeImageId = "Subscript";
            this.BtnSubscript.ShowImage = true;
            this.BtnSubscript.SuperTip = "Set numbers in a string to subscript in the selected range.";
            this.BtnSubscript.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnSubscript_Click);
            // 
            // AdditionalMenu
            // 
            this.AdditionalMenu.Items.Add(this.BtnDeleteRows);
            this.AdditionalMenu.Items.Add(this.BtnDeleteColumns);
            this.AdditionalMenu.Items.Add(this.BtnFillDown);
            this.AdditionalMenu.Items.Add(this.BtnResetColumn);
            this.AdditionalMenu.Items.Add(this.separator5);
            this.AdditionalMenu.Items.Add(this.BtnCopyFormatting);
            this.AdditionalMenu.Items.Add(this.BtnCopyHighlightedValues);
            this.AdditionalMenu.Items.Add(this.BtnCopyHighlightedRows);
            this.AdditionalMenu.Items.Add(this.separator6);
            this.AdditionalMenu.Items.Add(this.BtnDisplayLength);
            this.AdditionalMenu.Items.Add(this.BtnSheetNames);
            this.AdditionalMenu.Items.Add(this.BtnFileList);
            this.AdditionalMenu.Label = "Additional Options";
            this.AdditionalMenu.Name = "AdditionalMenu";
            this.AdditionalMenu.OfficeImageId = "ControlWizards";
            this.AdditionalMenu.ShowImage = true;
            this.AdditionalMenu.ShowLabel = false;
            this.AdditionalMenu.SuperTip = "Additional Options";
            // 
            // BtnDeleteRows
            // 
            this.BtnDeleteRows.Label = "Delete Empty &Rows";
            this.BtnDeleteRows.Name = "BtnDeleteRows";
            this.BtnDeleteRows.OfficeImageId = "DeleteRows";
            this.BtnDeleteRows.ShowImage = true;
            this.BtnDeleteRows.SuperTip = "Delete any empty rows on the active worksheet.";
            this.BtnDeleteRows.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnDeleteRows_Click);
            // 
            // BtnDeleteColumns
            // 
            this.BtnDeleteColumns.Label = "Delete Empty &Columns";
            this.BtnDeleteColumns.Name = "BtnDeleteColumns";
            this.BtnDeleteColumns.OfficeImageId = "DeleteColumns";
            this.BtnDeleteColumns.ShowImage = true;
            this.BtnDeleteColumns.SuperTip = "Delete any empty columns on the active worksheet.";
            this.BtnDeleteColumns.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnDeleteColumns_Click);
            // 
            // BtnFillDown
            // 
            this.BtnFillDown.Label = "Fill &Down";
            this.BtnFillDown.Name = "BtnFillDown";
            this.BtnFillDown.OfficeImageId = "InsertRowBelow";
            this.BtnFillDown.ShowImage = true;
            this.BtnFillDown.SuperTip = "Fill any blank cells with the value above in the selected range.";
            this.BtnFillDown.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnFillDown_Click);
            // 
            // BtnResetColumn
            // 
            this.BtnResetColumn.Label = "R&eset Column";
            this.BtnResetColumn.Name = "BtnResetColumn";
            this.BtnResetColumn.OfficeImageId = "NumberFormatGallery";
            this.BtnResetColumn.ShowImage = true;
            this.BtnResetColumn.SuperTip = "Reset the selected column. Numbers stored as text will be converted to numbers.";
            this.BtnResetColumn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnResetColumn_Click);
            // 
            // separator5
            // 
            this.separator5.Name = "separator5";
            // 
            // BtnCopyFormatting
            // 
            this.BtnCopyFormatting.Label = "Copy &Formatting to All Sheets";
            this.BtnCopyFormatting.Name = "BtnCopyFormatting";
            this.BtnCopyFormatting.OfficeImageId = "PasteFormatting";
            this.BtnCopyFormatting.ShowImage = true;
            this.BtnCopyFormatting.SuperTip = "Copy all formatting from the active worksheet to all worksheets.";
            this.BtnCopyFormatting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnCopyFormatting_Click);
            // 
            // BtnCopyHighlightedValues
            // 
            this.BtnCopyHighlightedValues.Label = "Copy &Highlighted to Clipboard";
            this.BtnCopyHighlightedValues.Name = "BtnCopyHighlightedValues";
            this.BtnCopyHighlightedValues.OfficeImageId = "HighlighterMode";
            this.BtnCopyHighlightedValues.ShowImage = true;
            this.BtnCopyHighlightedValues.SuperTip = "Copies all highlighted (colored) non-empty cells from your selection to the clipb" +
    "oard.";
            this.BtnCopyHighlightedValues.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnCopyHighlightedValues_Click);
            // 
            // BtnCopyHighlightedRows
            // 
            this.BtnCopyHighlightedRows.Label = "Copy Hi&ghlighted Rows to New Sheet";
            this.BtnCopyHighlightedRows.Name = "BtnCopyHighlightedRows";
            this.BtnCopyHighlightedRows.OfficeImageId = "ConditionalFormattingColorScalesGallery";
            this.BtnCopyHighlightedRows.ShowImage = true;
            this.BtnCopyHighlightedRows.SuperTip = "Creates a new worksheet with all highlighted rows from your active worksheet. Any" +
    " row that contains at least one highlighted cell.";
            this.BtnCopyHighlightedRows.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnCopyHighlightedRows_Click);
            // 
            // separator6
            // 
            this.separator6.Name = "separator6";
            // 
            // BtnDisplayLength
            // 
            this.BtnDisplayLength.Label = "Display &Length [Off]";
            this.BtnDisplayLength.Name = "BtnDisplayLength";
            this.BtnDisplayLength.OfficeImageId = "WordCount";
            this.BtnDisplayLength.ShowImage = true;
            this.BtnDisplayLength.SuperTip = "Displays the length of the cell in the status bar at the bottom left. Toggle [On/" +
    "Off].";
            this.BtnDisplayLength.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnDisplayLength_Click);
            // 
            // BtnSheetNames
            // 
            this.BtnSheetNames.Label = "Sheet &Names to Clipboard";
            this.BtnSheetNames.Name = "BtnSheetNames";
            this.BtnSheetNames.OfficeImageId = "PasteMergeList";
            this.BtnSheetNames.ShowImage = true;
            this.BtnSheetNames.SuperTip = "Create a list of worksheet names using the active workbook.";
            this.BtnSheetNames.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnSheetNames_Click);
            // 
            // BtnFileList
            // 
            this.BtnFileList.Label = "Create F&ile List";
            this.BtnFileList.Name = "BtnFileList";
            this.BtnFileList.OfficeImageId = "CopyToFolder";
            this.BtnFileList.ShowImage = true;
            this.BtnFileList.SuperTip = "Create a list of files using a selected folder.";
            this.BtnFileList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnFileList_Click);
            // 
            // BtnMissing
            // 
            this.BtnMissing.Label = "Find Missing";
            this.BtnMissing.Name = "BtnMissing";
            this.BtnMissing.OfficeImageId = "WhatIfAnalysisMenu";
            this.BtnMissing.ShowImage = true;
            this.BtnMissing.SuperTip = "Find Missing Data\n\nToggle on/off the task pane used for finding missing data betw" +
    "een two ranges.";
            this.BtnMissing.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnMissing_Click);
            // 
            // BtnDuplicates
            // 
            this.BtnDuplicates.Label = "Duplicates";
            this.BtnDuplicates.Name = "BtnDuplicates";
            this.BtnDuplicates.OfficeImageId = "PivotPlusMinusFieldHeadersShowHide";
            this.BtnDuplicates.ShowImage = true;
            this.BtnDuplicates.SuperTip = "Check for Duplicates\n\nCheck if a selected column contains duplicates. Toggle on/o" +
    "ff a count column based on the selected column.";
            this.BtnDuplicates.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnDuplicates_Click);
            // 
            // BtnCompare
            // 
            this.BtnCompare.Label = "Compare Sheets";
            this.BtnCompare.Name = "BtnCompare";
            this.BtnCompare.OfficeImageId = "ReviewCompareTwoVersions";
            this.BtnCompare.ShowImage = true;
            this.BtnCompare.SuperTip = "Compare Worksheets\n\nToggle on/off the task pane to compare two worksheets.";
            this.BtnCompare.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnCompare_Click);
            // 
            // BtnFilter
            // 
            this.BtnFilter.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnFilter.Label = "Filter";
            this.BtnFilter.Name = "BtnFilter";
            this.BtnFilter.OfficeImageId = "DataFilter";
            this.BtnFilter.ShowImage = true;
            this.BtnFilter.SuperTip = "Filter\n\nToggle filtering on/off";
            this.BtnFilter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnFilter_Click);
            // 
            // SBtnUniqueClipboard
            // 
            this.SBtnUniqueClipboard.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SBtnUniqueClipboard.Items.Add(this.BtnUniqueClipboard);
            this.SBtnUniqueClipboard.Items.Add(this.BtnCopyToSheets);
            this.SBtnUniqueClipboard.Items.Add(this.BtnColumnInfo);
            this.SBtnUniqueClipboard.Label = "Unique";
            this.SBtnUniqueClipboard.Name = "SBtnUniqueClipboard";
            this.SBtnUniqueClipboard.OfficeImageId = "TableColumnSelect";
            this.SBtnUniqueClipboard.SuperTip = "Unique Data Options";
            this.SBtnUniqueClipboard.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SBtnUniqueClipboard_Click);
            // 
            // BtnUniqueClipboard
            // 
            this.BtnUniqueClipboard.Label = "Selection to &Clipboard";
            this.BtnUniqueClipboard.Name = "BtnUniqueClipboard";
            this.BtnUniqueClipboard.OfficeImageId = "Copy";
            this.BtnUniqueClipboard.ShowImage = true;
            this.BtnUniqueClipboard.SuperTip = "Copy the selections unique data to clipboard.";
            this.BtnUniqueClipboard.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnUniqueClipboard_Click);
            // 
            // BtnCopyToSheets
            // 
            this.BtnCopyToSheets.Label = "Create &New Sheets";
            this.BtnCopyToSheets.Name = "BtnCopyToSheets";
            this.BtnCopyToSheets.OfficeImageId = "DuplicateThisPage";
            this.BtnCopyToSheets.ShowImage = true;
            this.BtnCopyToSheets.SuperTip = "Create worksheets based on the unique values of a column and copy data to each on" +
    "e.";
            this.BtnCopyToSheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnCopyToSheets_Click);
            // 
            // BtnColumnInfo
            // 
            this.BtnColumnInfo.Label = "Column &Information";
            this.BtnColumnInfo.Name = "BtnColumnInfo";
            this.BtnColumnInfo.OfficeImageId = "InformationDialog";
            this.BtnColumnInfo.ShowImage = true;
            this.BtnColumnInfo.SuperTip = "Displays details on the selected column: unique values, blanks, non-blanks, and t" +
    "otal row count.";
            this.BtnColumnInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnColumnInfo_Click);
            // 
            // BtnCommaSelection
            // 
            this.BtnCommaSelection.Label = "Selection";
            this.BtnCommaSelection.Name = "BtnCommaSelection";
            this.BtnCommaSelection.OfficeImageId = "ApplyCommaFormat";
            this.BtnCommaSelection.ShowImage = true;
            this.BtnCommaSelection.SuperTip = "Selection to Clipboard\n\nComma separate the selection to your clipboard.";
            this.BtnCommaSelection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnCommaSelection_Click);
            // 
            // BtnDelimSelection
            // 
            this.BtnDelimSelection.Label = "Selection+";
            this.BtnDelimSelection.Name = "BtnDelimSelection";
            this.BtnDelimSelection.OfficeImageId = "CommaSign";
            this.BtnDelimSelection.ShowImage = true;
            this.BtnDelimSelection.SuperTip = "Selection+ to Clipboard\n\nSet a delimiter plus any leading and trailing text to yo" +
    "ur clipboard.";
            this.BtnDelimSelection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnDelimSelection_Click);
            // 
            // BtnSheetToFile
            // 
            this.BtnSheetToFile.Label = "Sheet to File";
            this.BtnSheetToFile.Name = "BtnSheetToFile";
            this.BtnSheetToFile.OfficeImageId = "SectionsMerge";
            this.BtnSheetToFile.ShowImage = true;
            this.BtnSheetToFile.SuperTip = "Worksheet to File\n\nCreate a delimited file using the current active worksheet.";
            this.BtnSheetToFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnSheetToFile_Click);
            // 
            // BtnSplitToRows
            // 
            this.BtnSplitToRows.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnSplitToRows.Label = "Split to Rows";
            this.BtnSplitToRows.Name = "BtnSplitToRows";
            this.BtnSplitToRows.OfficeImageId = "TablesGallery";
            this.BtnSplitToRows.ShowImage = true;
            this.BtnSplitToRows.SuperTip = "Split Columns to Rows\n\nSplit delimited column(s) to rows.";
            this.BtnSplitToRows.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnSplitToRows_Click);
            // 
            // BtnHyperlinkSettings
            // 
            this.BtnHyperlinkSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnHyperlinkSettings.Label = "Custom URLs";
            this.BtnHyperlinkSettings.Name = "BtnHyperlinkSettings";
            this.BtnHyperlinkSettings.OfficeImageId = "GetPowerQueryDataSourceSettings";
            this.BtnHyperlinkSettings.ShowImage = true;
            this.BtnHyperlinkSettings.SuperTip = "Custom Links\n\nView and edit your parameterized URLs.";
            this.BtnHyperlinkSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnHyperlinkSettings_Click);
            // 
            // BtnHyperlinks
            // 
            this.BtnHyperlinks.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnHyperlinks.Items.Add(this.BtnAddHyperlinks);
            this.BtnHyperlinks.Items.Add(this.BtnAddHyperlinksCell);
            this.BtnHyperlinks.Items.Add(this.BtnRemoveHyperlinks);
            this.BtnHyperlinks.Label = "Add/Remove";
            this.BtnHyperlinks.Name = "BtnHyperlinks";
            this.BtnHyperlinks.OfficeImageId = "HyperlinkInsert";
            this.BtnHyperlinks.SuperTip = "Add or Remove Custom Hyperlinks";
            this.BtnHyperlinks.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnHyperlinks_Click);
            // 
            // BtnAddHyperlinks
            // 
            this.BtnAddHyperlinks.Label = "Add &Custom Hyperlinks";
            this.BtnAddHyperlinks.Name = "BtnAddHyperlinks";
            this.BtnAddHyperlinks.OfficeImageId = "PersonalizedHyperlinkInsert";
            this.BtnAddHyperlinks.ShowImage = true;
            this.BtnAddHyperlinks.SuperTip = "Add custom hyperlinks on a selected column using your active parameterized URL.";
            this.BtnAddHyperlinks.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnAddHyperlinks_Click);
            // 
            // BtnAddHyperlinksCell
            // 
            this.BtnAddHyperlinksCell.Label = "&Add Hyperlinks";
            this.BtnAddHyperlinksCell.Name = "BtnAddHyperlinksCell";
            this.BtnAddHyperlinksCell.OfficeImageId = "HyperlinkInsert";
            this.BtnAddHyperlinksCell.ShowImage = true;
            this.BtnAddHyperlinksCell.SuperTip = "Creates hyperlinks based on text URL in cell.";
            this.BtnAddHyperlinksCell.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnAddHyperlinksCell_Click);
            // 
            // BtnRemoveHyperlinks
            // 
            this.BtnRemoveHyperlinks.Label = "&Remove [Cell && Formula]";
            this.BtnRemoveHyperlinks.Name = "BtnRemoveHyperlinks";
            this.BtnRemoveHyperlinks.OfficeImageId = "HyperlinksRemove";
            this.BtnRemoveHyperlinks.ShowImage = true;
            this.BtnRemoveHyperlinks.SuperTip = "Remove both cell and formula based hyperlinks.";
            this.BtnRemoveHyperlinks.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnRemoveHyperlinks_Click);
            // 
            // BtnAbout
            // 
            this.BtnAbout.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnAbout.Image = ((System.Drawing.Image)(resources.GetObject("BtnAbout.Image")));
            this.BtnAbout.Label = "Quick Tools";
            this.BtnAbout.Name = "BtnAbout";
            this.BtnAbout.ShowImage = true;
            this.BtnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnAbout_Click);
            // 
            // XLQuickTools
            // 
            this.Name = "XLQuickTools";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.XLQuickTools_Tab);
            this.XLQuickTools_Tab.ResumeLayout(false);
            this.XLQuickTools_Tab.PerformLayout();
            this.Group_Formatting.ResumeLayout(false);
            this.Group_Formatting.PerformLayout();
            this.Group_Data.ResumeLayout(false);
            this.Group_Data.PerformLayout();
            this.Group_Delimiter.ResumeLayout(false);
            this.Group_Delimiter.PerformLayout();
            this.Group_Hyperlinks.ResumeLayout(false);
            this.Group_Hyperlinks.PerformLayout();
            this.Group_About.ResumeLayout(false);
            this.Group_About.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab XLQuickTools_Tab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group_Delimiter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnCommaSelection;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnDelimSelection;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group_Formatting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnSplitToRows;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnFillDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnSheetToFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group_Data;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnQuickSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group_Hyperlinks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnMissing;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnDuplicates;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnCompare;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnHyperlinkSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator Separator_Delimiter;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator Separator_Formatting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnCopyToSheets;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnDateText;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnUndo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnFilter;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator Separator_Data;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnUppercase;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnLowercase;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnPropercase;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnDeleteRows;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnDeleteColumns;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu CharacterMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu AdditionalMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnRemoveSpecial;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnRemoveNumbers;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnRemoveLetters;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnNormalizeText;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnResetColumn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnSubscript;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton SBtnTrimClean;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnTrimCleanWorkbook;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnTrimCleanWorksheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnTrimClean;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnAddLeadingTrailing;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnUniqueClipboard;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnRemoveNonASCII;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnDisplayLength;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton BtnRemoveExcess;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnRemoveExcessWS;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnRemoveExcessWB;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnRemoveSpaces;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnCopyFormatting;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton BtnQuickFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnQuickFormatSub;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnQuickFormatAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton BtnRemoveFormatting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnRemoveFormattingSub;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnRemoveFormattingAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton SBtnUniqueClipboard;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnSheetNames;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnFileList;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator4;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator5;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnConvertTableRemove;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnRemoveObjects;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnCopyHighlightedRows;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnCopyHighlightedValues;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton BtnHyperlinks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnAddHyperlinks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnRemoveHyperlinks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnColumnInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnTrimCleanSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnAddHyperlinksCell;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group_About;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnAbout;
    }

    partial class ThisRibbonCollection
    {
        internal XLQuickTools Ribbon1
        {
            get { return this.GetRibbon<XLQuickTools>(); }
        }
    }
}

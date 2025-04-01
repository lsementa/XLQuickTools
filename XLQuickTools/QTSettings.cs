using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace XLQuickTools
{
    public class QTSettings
    {
        // Folder and file path
        private static readonly string FolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "XLQuickTools");
        private static readonly string FilePath = Path.Combine(FolderPath, "settings.xml");

        // User settings that are saved
        public class UserSettings
        {
            // Formatting properties
            public bool Bold { get; set; }
            public bool Underline { get; set; }
            public bool Freeze { get; set; }
            public bool Border1 { get; set; }
            public bool Border2 { get; set; }
            public bool InteriorColor { get; set; }
            public string Color { get; set; }

            // AutoFit properties
            public bool AutoFit1 { get; set; }
            public bool AutoFit2 { get; set; }
            public bool AutoFilter { get; set; }
            public bool WrapText { get; set; }

            // View properties
            public string ZoomValue { get; set; }
            public bool Zoom { get; set; }
            public bool Gridlines { get; set; }

            // Size properties
            public bool ColWidth { get; set; }
            public bool RowHeight { get; set; }
            public string Height { get; set; }
            public string Width { get; set; }

            // Font properties
            public string FontSizeTxt { get; set; }
            public bool FontSize { get; set; }

            // Horizontal alignment properties
            public bool AllignLeft { get; set; }
            public bool AllignRight { get; set; }
            public bool AllignCenter { get; set; }
            public bool AllignLeft_FR { get; set; }
            public bool AllignRight_FR { get; set; }
            public bool AllignCenter_FR { get; set; }

            // Vertical alignment properties
            public bool AllignTop { get; set; }
            public bool AllignMiddle { get; set; }
            public bool AllignBottom { get; set; }
            public bool AllignTop_FR { get; set; }
            public bool AllignMiddle_FR { get; set; }
            public bool AllignBottom_FR { get; set; }

            // Hyperlink entries
            public List<HyperlinkEntry> HyperlinkEntries { get; set; } = new List<HyperlinkEntry>();
        }

        // User Settings Hyperlinks
        public class HyperlinkEntry
        {
            public string Name { get; set; }
            public string URL { get; set; }
            public bool Use { get; set; }
        }

        // Save user settings to an XML file
        public static void SaveUserSettingsToXml(UserSettings settings)
        {
            // Create XLQuickTools directory if it doesn't exist
            Directory.CreateDirectory(FolderPath);

            // Create an XML serializer for the UserSettings class
            var serializer = new XmlSerializer(typeof(UserSettings));
            using (var writer = new StreamWriter(FilePath))
            {
                serializer.Serialize(writer, settings);
            }
        }

        // Load the user settings XML file
        public static UserSettings LoadUserSettingsFromXml()
        {
            if (File.Exists(FilePath))
            {
                try
                {
                    var serializer = new XmlSerializer(typeof(UserSettings));
                    using (var reader = new StreamReader(FilePath))
                    {
                        return (UserSettings)serializer.Deserialize(reader);
                    }
                }
                catch (Exception)
                {
                    // If deserialization fails, return default settings
                    return new UserSettings();
                }
            }

            // Return default settings if file doesn't exist
            return new UserSettings();
        }
    }
}

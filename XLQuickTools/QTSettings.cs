using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace XLQuickTools
{
    public class QTSettings
    {

        // User settings that are saved
        public class UserSettings
        {
            public bool Bold { get; set; }
            public bool Underline { get; set; }
            public bool Freeze { get; set; }
            public bool Border1 { get; set; }
            public bool Border2 { get; set; }
            public bool InteriorColor { get; set; }
            public string Color { get; set; }
            public bool AutoFit1 { get; set; }
            public bool AutoFit2 { get; set; }
            public bool AutoFilter { get; set; }
            public bool WrapText { get; set; }
            public string ZoomValue { get; set; }
            public bool Zoom { get; set; }
            public bool Gridlines { get; set; }
            public bool ColWidth { get; set; }
            public bool RowHeight { get; set; }
            public string Height { get; set; }
            public string Width { get; set; }
            public string FontSizeTxt { get; set; }
            public bool FontSize { get; set; }
            public bool AllignLeft { get; set; }
            public bool AllignRight { get; set; }
            public bool AllignCenter { get; set; }
            public bool AllignLeft_FR { get; set; }
            public bool AllignRight_FR { get; set; }
            public bool AllignCenter_FR { get; set; }
            public bool AllignTop { get; set; }
            public bool AllignMiddle { get; set; }
            public bool AllignBottom { get; set; }
            public bool AllignTop_FR { get; set; }
            public bool AllignMiddle_FR { get; set; }
            public bool AllignBottom_FR { get; set; }
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
            string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "XLQuickTools.xml");

            // Create an XML serializer for the UserSettings class
            XmlSerializer serializer = new XmlSerializer(typeof(UserSettings));

            using (StreamWriter writer = new StreamWriter(filePath))
            {
                serializer.Serialize(writer, settings); // Serialize the settings object to XML
            }
        }

        // Load the user settings XML file
        public static UserSettings LoadUserSettingsFromXml()
        {
            string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "XLQuickTools.xml");

            if (File.Exists(filePath))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(UserSettings));

                using (StreamReader reader = new StreamReader(filePath))
                {
                    return (UserSettings)serializer.Deserialize(reader); // Deserialize XML back to UserSettings object
                }
            }

            // Return default settings - blank if it doesn't exist
            return new UserSettings();
        }
    }
}

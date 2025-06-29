using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLQuickTools
{
    internal class QTConstants
    {
        // Start up
        public const int MAX_PANE_WIDTH = 500;

        // Quick Format
        public const int DEFAULT_FONT_SIZE = 11;
        public const int MAX_FONT_SIZE = 72;
        public const int MAX_ZOOM = 400;
        public const int MIN_ZOOM = 10;
        public const int DEFAULT_ZOOM = 100;
        public const double MAX_ROW_HEIGHT = 400;
        public const double DEFAULT_ROW_HEIGHT = 15;
        public const double MAX_COLUMN_WIDTH = 250;
        public const double DEFAULT_COLUMN_WIDTH = 8.43;

        // User Control (Task Panes)
        public const int DEFAULT_DATAGRID_ROWS = 25;
        public const int MAX_ROWS = 75;
        public const int MIN_ROWS = 1;

        // Formatting functions
        public const int UPPERCASE = 0;
        public const int LOWERCASE = 1;
        public const int PROPERCASE = 2;
        public const int REMOVE_LETTERS = 3;
        public const int REMOVE_NUMBERS = 4;
        public const int REMOVE_SPECIAL = 5;
        public const int NORMALIZE_TEXT = 6;
        public const int TRIMCLEAN = 7;
        public const int ADD_LEADING_TRAILING = 8;
        public const int REPLACE_NON_ASCII = 9;
        public const int REMOVE_SPACES = 10;

        // Add/Remove hyperlinks processing
        public const int CHUNK_SIZE = 5000;

        // Creating unique worksheets
        public const int MAX_SHEETS = 50;

    }
}

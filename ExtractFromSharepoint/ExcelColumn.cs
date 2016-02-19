namespace ExtractFromSharepoint
{
    /// <summary>
    /// Represents one column object in excel
    /// </summary>
    internal class ExcelColumn
    {
        /// <summary>
        /// Empty column constructor
        /// </summary>
        internal ExcelColumn()
        {
            Name = "";
            Width = 0;
        }

/*
        internal ExcelColumn(string name, int width)
        {
            Name = name;
            Width = width;
        }
*/

        /// <summary>
        /// The name of the column
        /// </summary>
        internal string Name { get; set; }

        /// <summary>
        /// The width to set the column too
        /// </summary>
        internal decimal Width { get; set; }
    }
}
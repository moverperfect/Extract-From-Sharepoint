namespace ExtractFromSharepoint
{
    internal class ExcelColumn
    {
        internal ExcelColumn()
        {
            Name = "";
            Width = 0;
        }

        internal ExcelColumn(string name, int width)
        {
            Name = name;
            Width = width;
        }

        internal string Name { get; set; }

        internal decimal Width { get; set; }
    }
}
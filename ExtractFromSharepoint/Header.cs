namespace ExtractFromSharepoint
{
    internal class Header
    {
        internal Header()
        {
            Height = 0;
            BackgroundColour = "0";
            TextColour = "0";
        }

        internal decimal Height { get; set; }

        internal string BackgroundColour { get; set; }

        internal string TextColour { get; set; }
    }
}
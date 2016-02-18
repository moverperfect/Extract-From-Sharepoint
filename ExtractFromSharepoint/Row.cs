namespace ExtractFromSharepoint
{
    class Row
    {
        internal Row()
        {
            
        }

        internal Row(string colour, int height)
        {
            Colour = colour;
            Height = height;
        }

        internal string Colour { get; set; }

        internal decimal Height { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

using System.Collections.Generic;
using OpenQA.Selenium;

namespace ExtractFromSharepoint
{
    /// <summary>
    /// One sharepoint item
    /// </summary>
    class AppDetail
    {
        /// <summary>
        /// Creates a new AppDetail from the informatino on an ieDriver page
        /// </summary>
        /// <param name="ieDriver">The ieDriver to be used to extract the data from</param>
        public AppDetail(ISearchContext ieDriver)
        {
            // Create empty lists
            ProperyNames = new List<string>();
            Properties = new List<string>();

            // Grab all of the rows from the page and iterate over them
            var rows = ieDriver.FindElement(By.ClassName("ms-formtable")).FindElements(By.TagName("tr"));
            foreach (var row in rows)
            {
                // Grab the 2 cells from each row
                var cells = row.FindElements(By.TagName("td"));

                // If the title is empty then continue
                if(cells[0].Text == "")
                    continue;

                // Grab the property name, property and any links that are within the property
                var propName = cells[0].Text;
                var links = cells[1].FindElements(By.TagName("a"));
                var prop = cells[1].Text;

                // For all of the links that were discovered
                foreach (var link in links)
                {
                    // Get the href of the link
                    var href = link.GetAttribute("href");

                    // If the href is an external site then continue
                    if(!href.Contains(Program.Domian))
                        continue;

                    // Add in a representation of the link and the name of the link
                    prop = prop.Replace(link.Text,
                        "||(|)||" + href + "||()||" + link.Text + "||(|)||");
                }

                // Add the properties to the object
                ProperyNames.Add(propName);
                Properties.Add(prop);
            }
        }

        /// <summary>
        /// Creates a new empty AppDetail
        /// </summary>
        public AppDetail()
        {
            ProperyNames = new List<string>();
            Properties = new List<string>();
        }

        /// <summary>
        /// A list of the property names within the item
        /// </summary>
        internal List<string> ProperyNames { get; }

        /// <summary>
        /// A list of the property values within the item
        /// </summary>
        internal List<string> Properties { get; }
    }
}

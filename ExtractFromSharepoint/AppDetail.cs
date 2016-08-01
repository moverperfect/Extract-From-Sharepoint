using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
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
            if(ieDriver == null)
                throw new ArgumentNullException(nameof(ieDriver));

            // Create empty lists
            ProperyNames = new List<string>();
            Properties = new List<string>();

            // Grab the form tables from the page
            var formtables = ieDriver.FindElements(By.ClassName("ms-formtable"));
            var tabElements = ieDriver.FindElements(By.ClassName("ui-tabs-anchor"));

            if (formtables.Count - 1 != tabElements.Count)
            {
                Console.WriteLine("WARNING! We have detected that you have invisible tables that may mess up the program");
                Console.WriteLine("We will try to detect them and work around them but this may not work");
            }

            // Used to blacklist formtables so that we do not loop over them
            var blacklist = new List<int>();

            for (var i = 0; i < formtables.Count - 1; i++)
            {
                tabElements[i - blacklist.Count < 0 ? 0 : i - blacklist.Count].Click();
                tabElements[i - blacklist.Count < 0 ? 0 : i - blacklist.Count].Click();

                var formtable = formtables[i];
                if (formtable.FindElements(By.ClassName("ms-formtable")).Count > 0)
                {
                    Console.WriteLine("Invisible table has been found inside tab " + (i+1));
                    blacklist.Add(i+1);
                }

                // If this table is in the blacklist then continue over it
                if (blacklist.Contains(i))
                {
                    continue;
                }

                // Grab all of the rows from the form page and iterate over them
                var rows = formtable.FindElements(By.TagName("tr"));
                foreach (var row in rows)
                {
                    // Grab the 2 cells from each row
                    var cells = row.FindElements(By.TagName("td"));

                    // If the title is empty then continue or if the 'title' does not contain 'ms-formlabel' as a class then it is not a title
                    if (cells[0].Text == "" || !cells[0].GetAttribute("class").Contains("ms-formlabel"))
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

                        var unspupportedRegex =
                            new Regex(
                                "(^(PRN|AUX|NUL|CON|COM[1-9]|LPT[1-9]|(\\.+)$)(\\..*)?$)|(([\\x00-\\x1f\\\\?*:\";‌​|/<>‌​])+)",
                                RegexOptions.IgnoreCase);

                        // If the href is an external site then continue
                        if (!href.Contains(Program.Domian) || unspupportedRegex.IsMatch(link.Text))
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

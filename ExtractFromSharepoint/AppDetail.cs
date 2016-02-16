using System.Collections.Generic;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;

namespace ExtractFromSharepoint
{
    class AppDetail
    {
        public AppDetail(InternetExplorerDriver ieDriver)
        {
            ProperyNames = new List<string>();
            Properties = new List<string>();
            var table = ieDriver.FindElement(By.ClassName("ms-formtable"));
            var rows = table.FindElements(By.TagName("tr"));
            foreach (var row in rows)
            {
                var cells = row.FindElements(By.TagName("td"));
                var propName = cells[0].Text;
                var links = cells[1].FindElements(By.TagName("a"));
                var prop = cells[1].Text;
                for (var i = 0; i < links.Count; i++)
                {
                    var href = links[i].GetAttribute("href");
                    if(!href.Contains(Program.Domian))
                        continue;
                    prop = prop.Replace(links[i].Text,
                        "||(|)||" + href + "||()||" + links[i].Text + "||(|)||");
                }
                ProperyNames.Add(propName);
                Properties.Add(prop);
            }
        }

        internal List<string> ProperyNames { get; set; }

        internal List<string> Properties { get; set; }
    }
}

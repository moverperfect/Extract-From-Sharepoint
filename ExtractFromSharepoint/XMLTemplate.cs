using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace ExtractFromSharepoint
{
    /// <summary>
    /// Very powerfull tool to write XML with poperties and sub properties
    /// </summary>
    public class XmlTemplate
    {
        /// <summary>
        /// Assign all of the properies
        /// </summary>
        /// <param name="thisName">The name of this 'Property'</param>
        /// <param name="propertyNames">The names of the string properties</param>
        /// <param name="properties">The string properties</param>
        /// <param name="propertyAttributeNames">Names of the attributes to add</param>
        /// <param name="subProperties">A XmlTemplate filled with its own elements that is part of this element</param>
        /// <param name="propertyAttributes">Attributes to add</param>
        public XmlTemplate(string thisName, List<string> propertyNames, List<string> properties,
            List<string> propertyAttributes, List<string> propertyAttributeNames, List<XmlTemplate> subProperties)
        {
            Properties = properties ?? new List<string>();
            SubProperties = subProperties ?? new List<XmlTemplate>();
            ThisName = thisName ?? "";
            PropertyNames = propertyNames ?? new List<string>();
            PropertyAttributes = propertyAttributes ?? new List<string>();
            PropertyAttributeNames = propertyAttributeNames ?? new List<string>();
        }

        /// <summary>
        /// The names of all of the String Properties that are a part of this element
        /// </summary>
        internal List<string> PropertyNames { private get; set; }

        /// <summary>
        /// The string properties that are a part of this element
        /// </summary>
        internal List<string> Properties { private get; set; }

        /// <summary>
        /// The list of attribute values to be inserted
        /// </summary>
        internal List<string> PropertyAttributes { private get; set; }

        /// <summary>
        /// List of the attribute names
        /// </summary>
        internal List<string> PropertyAttributeNames { private get; set; } 

        /// <summary>
        /// The name of this element
        /// </summary>
        private string ThisName { get; }

        /// <summary>
        /// A list of the elements that are within this element with their own properties e.g tires are all inside the element of a car
        /// </summary>
        internal List<XmlTemplate> SubProperties { get; }

        /// <summary>
        /// Write all of the data to a XmlWriter
        /// </summary>
        /// <param name="w">The XmlWriter to write all of the data to</param>
        /// <returns>The XmlWriter that all of the data has been written to</returns>
        internal XmlWriter GetData(XmlWriter w)
        {
            // Start of this element
            w.WriteStartElement(ThisName);

            // Write all of the strings for this element
            for (var i = 0; i < Properties.Count; i++)
            {
                w.WriteStartElement(PropertyNames[i]);
                w.WriteAttributeString(PropertyAttributeNames[i], PropertyAttributes[i]);
                w.WriteCData(Properties[i]);
                w.WriteEndElement();
            }

            // Write all of the elements that are inside this element
            w = SubProperties.Aggregate(w, (current, t) => t.GetData(current));

            // End this element and return
            w.WriteEndElement();
            return w;
        }
    }
}
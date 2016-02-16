using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ExtractFromSharepoint
{
    static class UserDetails
    {
        /// <summary>
        /// Enrry into the UserDetails main menu
        /// </summary>
        public static void Main()
        {
            while (true)
            {
                Console.Clear();
                Console.WriteLine("Sharepoint settings:");
                Console.WriteLine("AdDomain: " + Program.AdDomain);
                Console.WriteLine("Domain: " + Program.Domian);
                Console.WriteLine("ListUrl: " + Program.ListUrl);
                Console.WriteLine("Username: " + Program.Username);
                Console.WriteLine("1. Change Settings");
                Console.WriteLine("2. Back");

                var option = Console.ReadLine();

                switch (option)
                {
                    case "1":
                        GetUserInfo();
                        SaveUserConfig();
                        break;

                    case "2":
                        return;

                    default:
                        continue;
                }
            }
        }

        /// <summary>
        /// Grabs all of the user information about the website to scrape and the username and password to use
        /// </summary>
        internal static void GetUserInfo()
        {
            GetUserPassword();

            Console.WriteLine("Enter website active directory Domain");
            Program.AdDomain = Console.ReadLine();

            Console.WriteLine("Please enter your active directory username");
            Program.Username = Console.ReadLine();

            Console.WriteLine("Please enter the url of the list you would like to extract from");
            Program.ListUrl = Console.ReadLine();

            Console.WriteLine("Please enter the domain name of your site e.g google.com");
            Program.Domian = Console.ReadLine();
        }

        /// <summary>
        /// Gets the user's password
        /// </summary>
        private static void GetUserPassword()
        {
            ConsoleKeyInfo key;

            // Take in the users password
            Console.Write("Enter active directory password: ");
            do
            {
                key = Console.ReadKey(true);
                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                {
                    Program.Password += key.KeyChar;
                    Console.Write("*");
                }
                else
                {
                    if (key.Key == ConsoleKey.Backspace && Program.Password.Length > 0)
                    {
                        Program.Password = Program.Password.Substring(0, (Program.Password.Length - 1));
                        Console.Write("\b \b");
                    }
                }
            } while (key.Key != ConsoleKey.Enter);
            Console.Clear();
        }

        /// <summary>
        /// Saves the users configuration to a file
        /// </summary>
        internal static void SaveUserConfig()
        {
            var w = new XmlTextWriter("User.config", Encoding.Default);
            
            w.WriteStartDocument();
            w.WriteStartElement("config");
            w.WriteElementString("AdDomain", Program.AdDomain);
            w.WriteElementString("Username", Program.Username);
            w.WriteElementString("ListUrl", Program.ListUrl);
            w.WriteElementString("Domain", Program.Domian);
            w.WriteEndElement();
            w.WriteEndDocument();

            w.Close();
        }
    }
}

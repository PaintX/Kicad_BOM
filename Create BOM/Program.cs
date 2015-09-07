using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;

namespace ConsoleApplication1
{
    class Program
    {
        static Component component = new Component();
        static String xmlFileName = "";

        static void Main(string[] args)
        {
            Boolean isFileEntry = false;
            
            foreach (String arg in args)
            {
                if ( arg == "<" )
                {
                    isFileEntry = true;
                }
                else if ( arg == ">" )
                {
                    isFileEntry = false;
                }
                else if (isFileEntry == true)
                {
                    xmlFileName = arg.Trim();
                }
            }

            XmlDocument xml = new XmlDocument();
            xml.Load(xmlFileName);

            XmlElement root = xml.DocumentElement;
            XmlNodeList xmlNodes = root.SelectNodes("/export/components/comp");

            foreach (XmlNode xn in xmlNodes)
            {
                //-- detection d'un comosant

                //-- on sauvegarde l'acien composant
                component.save();

                //-- recupération du nomage du composant
                component.setRef(xn.Attributes["ref"].Value);

                Console.Write(xn.Attributes["ref"].Value + " => ");

                //-- recupération de la valeur du composant
                XmlNode xm = xn.SelectSingleNode("value");
                component.setValue(xm.InnerText);

                Console.Write(xm.InnerText + " ");

                XmlNodeList xNodes = xn.SelectNodes("fields/field");
                foreach (XmlNode x in xNodes)
                {
                    try
                    {
                        if (x.Attributes["name"].Value.ToLower().Contains("ref") && x.Attributes["name"].Value.ToLower().Contains("digikey"))
                        {
                            //-- recupération de la valeur du composant
                            component.setRefDigikey(x.InnerText);
                            Console.Write(x.InnerText + " ");
                            break;
                        }
                    }
                    catch
                    {
                        Console.Write("no");
                    }
                }

                Console.WriteLine("");
            }
            component.save();
            component.createList();

            component.createBOM();


            return;
        }
    }
}

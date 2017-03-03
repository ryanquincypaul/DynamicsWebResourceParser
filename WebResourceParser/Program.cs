using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Threading;

namespace DynamicsWebResourceParser
{
    class Program
    {
        public const string CUSTOMIZATIONS_FILE_NAME = "customizations.xml";
        public const string WEBRESOURCES_DIR_NAME = "WebResources/";
        public const string WEBRESOURCE_NODE_NAME = "WebResource";
        public const char SOLUTION_FILENAME_DELIMITER = '_';

        //TODO: Break into functions
        static void Main(string[] args)
        {
            String solutionArchiveFileName = args[0];
            
            //does file exist?
            if (File.Exists(solutionArchiveFileName))
            {
                Console.WriteLine(string.Format("Parsing {0} for Web Resource Files.", solutionArchiveFileName));
                string solutionName = solutionArchiveFileName.Split('.')[0];
                using (ZipArchive archive = ZipFile.OpenRead(solutionArchiveFileName))
                {
                    List<WebResource> webResources = GetWebResourceMetaData(archive);

                    if (webResources.Any())
                    {
                        CollectAndZipWebResourceFiles(webResources, archive, solutionName);
                    }
                    else
                    {
                        Console.WriteLine("No Web Resources to collect. No zip file will be created.");
                    }

                }
            }
            else
            {
                Console.WriteLine("Solution file does not exist, please check the path/filename and try again.");
            }
            Console.WriteLine("Process completed, press any key to end");
            Console.ReadKey();
        }

        /// <summary>
        /// Collect web resource metadata that will be used to get and name the files with the correct file ending.
        /// </summary>
        /// <param name="archive"></param>
        /// <returns>List of WebResource objects.</returns>
        static List<WebResource> GetWebResourceMetaData(ZipArchive archive)
        {
            ZipArchiveEntry customizationsFileEntry = archive.Entries.Where(file => file.Name == CUSTOMIZATIONS_FILE_NAME).FirstOrDefault();

            List<WebResource> webResources = new List<WebResource>();

            if (customizationsFileEntry != null)
            {
                using (Stream stream = customizationsFileEntry.Open())
                {
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        XDocument webResourcesXml = XDocument.Parse(reader.ReadToEnd());
                        foreach (XElement webResourceNode in webResourcesXml.Descendants(WEBRESOURCE_NODE_NAME))
                        {
                            XmlSerializer serializer = new XmlSerializer(typeof(WebResource));
                            webResources.Add((WebResource)serializer.Deserialize(webResourceNode.CreateReader()));
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("Required file 'customizations.xml' does not exist. Please check the contects of the solution file and try again.");
            }

            return webResources;
        }

        /// <summary>
        /// Collects files based on the web resource ids found in the resource metadata, determines the file type for them, and correctly names them.
        /// The extracted and named files will then be zipped into a new archive.
        /// </summary>
        /// <param name="webResources"></param>
        /// <param name="archive"></param>
        /// <param name="solutionName"></param>
        static void CollectAndZipWebResourceFiles(List<WebResource> webResources, ZipArchive archive, string solutionName)
        {
            //CreateDirectory that will be zipped
            System.IO.Directory.CreateDirectory(solutionName);

            try
            {
                foreach (WebResource resource in webResources)
                {
                    //Find zip entry for each web resource mentioned in customizations.xml
                    ZipArchiveEntry webResourceEntry = archive.Entries.Where(entry => entry.Name.Contains(resource.WebResourceId.ToString().ToUpper())).FirstOrDefault();
                    if (webResourceEntry != null)
                    {
                        string fileEnding = GetFileEnding(resource.WebResourceType);
                        if (!String.IsNullOrWhiteSpace(fileEnding))
                        {

                            webResourceEntry.ExtractToFile(solutionName + "\\" + resource.Name + fileEnding);

                        }
                        else //web resource is an invalid type and will require a manual download from the solution dialog in crm
                        {
                            Console.WriteLine(string.Format("Skipping file {0}", webResourceEntry.FullName));
                        }

                    }
                }
                ZipWebResources(solutionName);
                Console.WriteLine(string.Format("Zip file {0}.zip successfully created.", solutionName));
            }
            catch (IOException ioe) //thrown when the zip file already exists
            {
                Console.WriteLine(string.Format("Error creating Zip file {0}.zip. Check to see if it already exists.", solutionName));
            }

            try
            {
                System.IO.Directory.Delete(solutionName, true);
            }
            catch (IOException ioe) //sometimes Directory delete throws an exception if it isn't empty despite giving it the true parameter.
            {
                Console.WriteLine("Error when deleting temp directory. Sleeping a half second and then will try again.");
                Thread.Sleep(500); //HACK: This prevents intermittent issue deletion exception
                System.IO.Directory.Delete(solutionName, true);
            }
        }

        static void ZipWebResources(string directoryName)
        {
            ZipFile.CreateFromDirectory(directoryName, directoryName + "_WebResources.zip");
        }

        static string GetFileEnding(int webResourceType)
        {
            string fileEnding = string.Empty;
            if (Enum.IsDefined(typeof(WebResourceTypes), webResourceType))
            {
                fileEnding = "." + ((WebResourceTypes)webResourceType).ToString();
            }
            else
            {
                Console.WriteLine(string.Format("WebResourceType {0} is invalid.", webResourceType.ToString()));
            }

            return fileEnding;
        }
    }


    /// <summary>
    /// File type is represented with an integer. These are pulled from https://msdn.microsoft.com/en-us/library/mt622399.aspx.
    /// </summary>
    //TODO: Find another way to pull this? Maybe it's stored somewhere in the file itself?
    enum WebResourceTypes
    {
        html = 1, //Webpage (HTML)
        css = 2, //Style Sheet (CSS)
        js = 3, //Script (JScript)
        xml = 4, //Data (XML)
        png = 5, //PNG format
        jpg = 6, //JPG format
        gif = 7, //GIF format
        xap = 8, //Silverlight (XAP)
        xsl = 9, //Style Sheet (XSL)
        ico = 10 //ICO format
    }
}

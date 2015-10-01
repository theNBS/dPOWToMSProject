using System.IO;
using dPowToMSProject;

namespace dPowToMSProjectConsole
{
    class Program
    {
        /// <summary>
        /// Example command line application for performing a conversion
        /// </summary>
        /// <param name="args">command line arguments</param>
        static void Main(string[] args)
        {
            // Path to input dPOW file
            string inputFile = "..\\..\\dPow\\004-Newtown_Country_Park.dpow";

            // Path to write output to
            // Extension should be xml to generate MS Project compatible xml file
            string outputFile = "004-Newtown_Country_Park.xml";

            // Instantiate the convertor and perform the conversion
            Convertor convertor = new Convertor(inputFile);
            convertor.Convert(outputFile);

            // This will automatically open the created file in it's default program
            System.Diagnostics.Process.Start(Path.Combine(Directory.GetCurrentDirectory(), outputFile));
        }
    }
}

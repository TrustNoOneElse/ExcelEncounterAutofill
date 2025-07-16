// See https://aka.ms/new-console-template for more information

using ExcelEncountersAutofill;
using OfficeOpenXml;

ExcelPackage.License.SetNonCommercialPersonal("Florian Pabst");


var mappings = NameMappingReader.Start();
new ExcelDirectoryProcessor(mappings).ProcessDirectory();

Console.WriteLine("Press any key to exit.");
Console.Read();

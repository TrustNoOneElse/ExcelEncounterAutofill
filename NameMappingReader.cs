using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelEncountersAutofill
{
    internal class NameMappingReader
    {
        internal static List<NameMapping> Start()
        {
            try
            {
                List<NameMapping> mappings = ReadNameMappings();
                foreach (var mapping in mappings)
                {
                    Console.WriteLine($"{string.Join(", ", mapping.Names)} => {mapping.SpokenAs}");
                }
                return mappings;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error reading embedded resource: " + ex.Message);
            }
            return new List<NameMapping>();
        }

        private static List<NameMapping> ReadNameMappings()
        {
            var mappingList = new List<NameMapping>();
            // Adjust the resource name to match your actual namespace and file location.
            const string resourceName = "ExcelEncountersAutofill.Resources.nameMapping.txt";
            Assembly assembly = Assembly.GetExecutingAssembly();

            using Stream? resourceStream = assembly.GetManifestResourceStream(resourceName);
            if (resourceStream == null)
            {
                throw new InvalidOperationException($"Embedded resource '{resourceName}' not found.");
            }

            using var reader = new StreamReader(resourceStream);
            string? line;
            while ((line = reader.ReadLine()) != null)
            {
                string[] columns = line.Split('\t');

                // Split the first column by comma.
                List<string> names = columns[0]
                    .Split(['/'], StringSplitOptions.RemoveEmptyEntries)
                    .Select(name => name.Trim())
                    .ToList();

                string spokenAs = columns[1].Trim();

                mappingList.Add(new NameMapping
                {
                    Names = names,
                    SpokenAs = spokenAs
                });
            }

            return mappingList;
        }
    }
}

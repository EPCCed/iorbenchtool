/*
    Copyright [2020] [The University of Edinburgh]

    Licensed under the Apache License, Version 2.0 (the "License");
    you may not use this file except in compliance with the License.
    You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

    Unless required by applicable law or agreed to in writing, software
    distributed under the License is distributed on an "AS IS" BASIS,
    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
    See the License for the specific language governing permissions and
    limitations under the License.

    SPDX-License-Identifier: Apache-2.0
*/

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using CommandLine;
using CommandLine.Text;
using Serilog;
using Serilog.Events;
using System.Collections.Concurrent;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis;
using System.Reflection;
using System.Runtime.Loader;

namespace parser
{
    internal class CliOptions
    {
        // options
        [Option('r', "recursive", Required = false, HelpText = "Recursively search for IOR files under the path provided.")]
        public bool Recursive { get; set; }
        [Option("dump", Required = false, HelpText = "Dump the content of the parsed IOR files.")]
        public bool Dump { get; set; }
        [Option('v', "verbose", Required = false, HelpText = "Display all the detailed log information.")]
        public bool Verbose { get; set; }
        [Option("log", Required = false, HelpText = "Log information into a log file on directory path.")]
        public bool LogFile { get; set; }
        [Option('u', "unite", Required = false, HelpText = "Unite all Excel reports into a single one. Only valid when --recursive is present.")]
        public bool Unite { get; set; }
        [Option('t', "template", Required = false, HelpText = "Template path for the Excel reports.")]
        public string ExcelTemplate { get; set; }

        //values
        [Value(0, Required = true, MetaName = "DirPath", HelpText = "Directory path where IOR output file(s) can be found.")]
        public string DirPath { get; set; }

        //usage examples
        [Usage(ApplicationAlias = "ior-parser")]
        public static IEnumerable<Example> Examples
        {
            get
            {
                return new List<Example>()
                {
                    new Example("Parse a set of files under a directory", new CliOptions { DirPath = "../ior-files/" }),
                    new Example("Parse a set of subdirectories with IOR files each", new CliOptions {Recursive = true , DirPath = "../folder-root/" }),
                    new Example("Parse a set of subdirectories then, summarizes all the reports into a single one", new CliOptions { DirPath = "../folder-root/", Recursive = true, Unite = true }),
                    new Example("Parse a set of IOR files with a custom Excel template", new CliOptions { DirPath = "../ior-files/", ExcelTemplate = "./ExcelTemplate.cs" })
                };
            }
        }
    }

    internal class Program
    {
        public static bool Dump;
        public static ConcurrentBag<(string, string)> ExcelsLocation; //DirName, LocationPath

        private static Type CustomClass;

        static void Main(string[] args)
        {
            var cmd = Parser.Default.ParseArguments<CliOptions>(args)
                .WithParsed(cli =>
                {
                    var parentDir = new DirectoryInfo(cli.DirPath);

                    var logConfig = new LoggerConfiguration()
                        .MinimumLevel.Verbose()
                        .WriteTo.Conditional(f => cli.Verbose, wt => wt.Console())
                        .WriteTo.Conditional(f => !cli.Verbose, wt => wt.Console(
                            restrictedToMinimumLevel: LogEventLevel.Information));

                    if (cli.LogFile)
                    {
                        string path = Path.Combine(parentDir.FullName, $"log-{DateTime.Now.ToString("dd-M-yyyy--HH-mm-ss")}.log");
                        logConfig.WriteTo.File(path);
                    }

                    Log.Logger = logConfig.CreateLogger();

                    Dump = cli.Dump;

                    Log.Debug("Root folder: {FullName}.", parentDir.FullName);

                    CustomClass = LoadCustomTemplate(cli.ExcelTemplate);

                    if (cli.Recursive)
                    {
                        Log.Debug("-r option is present.");

                        if (cli.Unite)
                        {
                            Log.Debug("-u option is present.");
                            ExcelsLocation = new ConcurrentBag<(string, string)>();
                        }

                        parentDir.GetDirectories()
                            .AsParallel()
                            .ForAll(dir => ParseDirectory(dir));

                        if (cli.Unite)
                        {
                            string reportLocation = Path.Combine(parentDir.FullName,
                                $"United-Benchmark-{parentDir.Name}-{DateTime.Now.ToString("dd-M-yyyy--HH-mm-ss")}.xlsx");
                            ExcelParser.CreateExcel(reportLocation);

                            Log.Debug("ExcelsLocation is null: {Null}", ExcelsLocation == null);
                            Log.Verbose("ExcelsLocation content: {Content}", ObjectDumper.Dump(ExcelsLocation));

                            Log.Information("Summarizing all reports into single one.");

                            using (var excel = new XLWorkbook(reportLocation))
                            {
                                Log.Debug("Adding sheets.");
                                foreach (var name in ExcelsLocation)
                                {
                                    excel.AddWorksheet(name.Item1);
                                }

                                excel.Worksheets.First().Delete();

                                excel.Save();
                            }

                            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(reportLocation, true))
                            {
                                Log.Debug("Creating charts.");

                                foreach (var sheet in doc.WorkbookPart.WorksheetParts.Zip(ExcelsLocation))
                                {
                                    // var genClass = new GeneratedCode.GeneratedClass();
                                    var genClass = (IGenerated)Activator.CreateInstance(CustomClass);
                                    genClass.CreateWorksheetPart(sheet.First, sheet.Second.Item1);
                                }

                                doc.Save();
                            }

                            using (var excel = new XLWorkbook(reportLocation))
                            {
                                foreach (var conn in ExcelsLocation)
                                {
                                    Log.Verbose("Coping {Dir} into summary report.", conn.Item1);
                                    using (var tmpExcel = new XLWorkbook(conn.Item2))
                                    {
                                        var cellsUsed = tmpExcel.Worksheets.First().CellsUsed();

                                        foreach (var cell in cellsUsed)
                                        {
                                            Log.Verbose("Coping {Value} from: {OgExcel}{OgCell} to {NExcel}{NCell}", cell.Value, conn.Item1, cell.Address, "Sum", cell.Address);
                                            excel.Worksheet(conn.Item1).Cell(cell.Address).Value = cell.Value;
                                        }
                                    }
                                }

                                excel.Save();
                            }
                        }
                    }
                    else
                    {
                        ParseDirectory(parentDir);
                    }
                });
            Log.Information("Done.");
        }

        private static void ParseDirectory(DirectoryInfo outPath)
        {
            Log.Information("Parsing {Name} directory.", outPath.Name);

            var files = outPath.GetFiles();
            Log.Verbose("{Name} internals: {Files}", outPath.Name, files);

            string reportLocation = Path.Combine(outPath.FullName,
                $"Benchmark-{outPath.Name}-{DateTime.Now.ToString("dd-M-yyyy--HH-mm-ss")}.xlsx");

            // Parallel file parsing.
            var infos = files.AsParallel()
                             .Select(file => FileParser.GenerateInfo(file))
                             .Where(inf => inf != null)
                             .ToArray();

            // Dumps if --dump is present.
            if (Dump)
            {
                Log.Debug("Dumping IOR data.");

                var resultDump = ObjectDumper.Dump(infos);
                string dumpFile = Path.Combine(outPath.FullName,
                     $"TextDump-{outPath.Name}-{DateTime.Now.ToString("dd-M-yyyy--HH-mm-ss")}.txtdmp");

                File.WriteAllText(dumpFile, resultDump);

                Log.Information("Dumped parsed information into {DumpFile}", dumpFile);
            }

            // Creates and Excel file. First, chart is drawn following GeneratedCode/GeneratedClass.cs blueprint.
            ExcelParser.CreateExcel(reportLocation);
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(reportLocation, true))
            {
                // var genClass = new GeneratedCode.GeneratedClass();
                var genClass = (IGenerated)Activator.CreateInstance(CustomClass);
                genClass.CreateWorksheetPart(doc.WorkbookPart.WorksheetParts.First(), "Benchmark Test Report");
            }

            // Excel created above is populated with the IOR files data.
            using (var excel = new XLWorkbook(reportLocation))
            {
                ExcelParser.PortInformation(excel, infos);
                excel.Save();
            }

            ExcelsLocation?.Add((outPath.Name, reportLocation));

            Log.Information("Directory {Name} parsing completed!", outPath.Name);
        }

        private static Type LoadCustomTemplate(string path)
        {
            Log.Debug("Checking for custom template.");

            if (path is null)
            {
                Log.Debug("No custom template.");
                return typeof(GeneratedCode.GeneratedClass);
            }

            Log.Debug("-t option is present. Path: {Path}", path);

            var templateText = File.ReadAllText(path);
            // var interfaceText = File.ReadAllText("./GeneratedCode/IGenerated.cs");

            var assemblyPath = Path.GetDirectoryName(typeof(object).Assembly.Location);
            var assemblies = new[]
            {
                typeof(object).GetTypeInfo().Assembly.Location,
                typeof(System.Runtime.GCSettings).GetTypeInfo().Assembly.Location,
                typeof(DocumentFormat.OpenXml.Packaging.WorksheetPart).GetTypeInfo().Assembly.Location,
                typeof(IGenerated).GetTypeInfo().Assembly.Location
            };

            var refs = assemblies.Distinct()
                                 .Select(a => MetadataReference.CreateFromFile(a))
                                 .ToList();

            refs.Add(MetadataReference.CreateFromFile(Path.Combine(assemblyPath, "mscorlib.dll")));
            refs.Add(MetadataReference.CreateFromFile(Path.Combine(assemblyPath, "System.dll")));
            refs.Add(MetadataReference.CreateFromFile(Path.Combine(assemblyPath, "System.Core.dll")));
            refs.Add(MetadataReference.CreateFromFile(Path.Combine(assemblyPath, "System.Runtime.dll")));

            var compilation = CSharpCompilation.Create("excel-template")
                .WithOptions(new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary,
                    usings: new[] { "DocumentFormat.OpenXml.Packaging" }))
                .AddReferences(refs)
                .AddSyntaxTrees(CSharpSyntaxTree.ParseText(File.ReadAllText(path)));

            var filename = "template.dll";
            var completePath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(path), filename));

            var result = compilation.Emit(completePath);

            if (!result.Success)
            {
                Log.Error("Error during compilation: {Msg}",
                    result.Diagnostics
                            .Select(e => e.GetMessage())
                            .Aggregate((s1, s2) => String.Format("{0}\n\t{1}", s1, s2)));

                throw new Exception("Compilation error!");
            }

            var template = AssemblyLoadContext.Default.LoadFromAssemblyPath(completePath);

            Log.Information("Template compiled!");

            Type templateType = null;
            foreach (var clazz in template.GetTypes())
            {
                foreach (var inhe in clazz.GetInterfaces())
                {
                    if (inhe == typeof(IGenerated))
                    {
                        Log.Debug("Valid template found!");
                        templateType = clazz;
                    }
                }
            }

            if (templateType is null)
            {
                Log.Warning("Invalid template! Using fallback Excel template.");
                templateType = typeof(GeneratedCode.GeneratedClass);
            }

            return templateType;
        }
    }
}
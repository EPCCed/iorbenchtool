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
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using Serilog;

namespace parser
{
    #region ContainerClasses
    /// <summary>
    /// Contains information about IOR and the machine in which the test were run.
    /// </summary>
    internal class Info
    {
        public string IOR_Version { get; set; }
        public string CmdUsed { get; set; }
        public string FileName { get; set; }
        public string Machine { get; set; }
        public DateTime StartTime { get; set; }
        public string FileSystem { get; set; }
        public Options Options { get; set; }
        public Result Result { get; set; }
    }

    /// <summary>
    /// Contains information about the configuration of the test itself.
    /// </summary>
    internal class Options
    {
        public string API { get; set; }
        public string API_Version { get; set; }
        public string AccessType { get; set; }
        public string TestType { get; set; }
        public string OrderInFile { get; set; }
        public ushort Tasks { get; set; }
        public ushort ClientsPerNode { get; set; }
        public ushort Repetitions { get; set; }
        public string XferSize { get; set; }
        public string BlockSize { get; set; }
        public string AggregateFileSize { get; set; }
    }

    /// <summary>
    /// Contains the differents <see cref="ResultInfo"/> about Writes, Reads and Removes.
    /// </summary>
    internal class Result
    {
        public List<ResultInfo> Writes { get; set; }
        public List<ResultInfo> Reads { get; set; }
        public List<float> Removes { get; set; }
    }

    /// <summary>
    /// Contains the information related to one interation of the test.
    /// </summary>
    internal class ResultInfo
    {
        public float Bandwidth_MiBs { get; set; }
        public ulong Block_KiB { get; set; }
        public float XFer_KiB { get; set; }
        public double Open_S { get; set; }
        public double WrRd_S { get; set; }
        public double Close_S { get; set; }
        public double Total { get; set; }
    }
    #endregion

    internal class FileParser
    {
        //Linux and Windows line break separators.
        private static string[] SEPARATORS = new[] { "\n", "\r\n" };

        /// <summary>
        /// Parses a string into a custom type date.
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        private static DateTime DateParser(string date)
        {
            DateTime new_date = DateTime.MinValue;

            try
            {
                new_date = DateTime.ParseExact(date, "ddd MMM dd HH:mm:ss yyyy",
                CultureInfo.CreateSpecificCulture("en-UK"));
            }
            catch (Exception)
            {
                Log.Error("Date {Date} could not be parsed into ddd MMM dd HH:mm:ss yyyy format!", date);
            }

            return new_date;
        }

        /// <summary>
        /// Generates the <see cref="Info"/> and sub-objects related to an IOR file.
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static Info GenerateInfo(FileInfo file)
        {
            Log.Debug("Parsing file {File}", file);

            // Break the file into lines.
            string[] fileStringSplit = file.OpenText().ReadToEnd()
                                       .Split(SEPARATORS, StringSplitOptions.RemoveEmptyEntries);

            int optionIndex = -1, resultIndex = -1;
            Task<Info> infoTask = null;
            Task<Options> optTask = null;
            Task<Result> resTask = null;

            // Creates tasks foreach of the file blocks
            for (int i = 0; i < fileStringSplit.Length; i++)
            {
                if (fileStringSplit[i].StartsWith("Options:"))
                {
                    optionIndex = i;
                    infoTask = Task.Run(() => ParseInformation(fileStringSplit[..optionIndex]));

                    Log.Verbose("{DirectoryName}/{Name} Information block with index 0-{OptionIndex} found!",
                        file.DirectoryName, file.Name, optionIndex);
                }
                else if (fileStringSplit[i].StartsWith("Results:"))
                {
                    resultIndex = i;
                    optTask = Task.Run(() => ParseOptions(fileStringSplit[(optionIndex + 1)..resultIndex]));
                    resTask = Task.Run(() => ParseResults(fileStringSplit[(resultIndex + 1)..], file.Name));

                    Log.Verbose("{DirectoryName}/{Name} Options block with index [{OptionIndex}..{ResultIndex}] found!",
                        file.DirectoryName, file.Name, optionIndex, resultIndex);
                    Log.Verbose("{DirectoryName}/{Name} Results block with index [{ResultIndex}..] found!",
                        file.DirectoryName, file.Name, resultIndex);
                }
            }

            // Invalid file format. Could not find the file blocks.
            if (optionIndex < 0 || resultIndex < 0)
            {
                Log.Verbose("{DirectoryName}/{Name} cannot be parsed! Invalid format.", file.DirectoryName, file.Name);
                return null;
            }

            Info inf = infoTask.Result;
            inf.FileName = file.Name;

            Options opt = optTask.Result;
            Result res = resTask.Result;

            inf.Options = opt;
            inf.Result = res;

            Log.Verbose("{DirectoryName}/{Name} parse is completed!", file.DirectoryName, file.Name);

            return inf;
        }

        /// <summary>
        /// Creates a Dictionary with Key the property in lowercase and Value the value of said property.
        /// </summary>
        /// <param name="chunk"></param>
        /// <returns></returns>
        private static Dictionary<string, string> Tokenize(string[] chunk)
        {
            var dic = new Dictionary<string, string>();
            foreach (var line in chunk)
            {
                var colonIndex = line.IndexOf(':');
                if (colonIndex < 0)
                {
                    continue;
                }

                var key = line[..colonIndex].Trim().ToLowerInvariant();
                var value = line[(colonIndex + 1)..].Trim();

                dic.Add(key, value);
            }

            return dic;
        }

        /// <summary>
        /// Parses the Information block into a <see cref="Info"/>
        /// </summary>
        /// <param name="chunk"></param>
        /// <returns></returns>
        private static Info ParseInformation(string[] chunk)
        {
            Info inf = new Info();
            Dictionary<string, string> dic = Tokenize(chunk);

            Log.Verbose("Tokenized Information chunk: {Dic}", dic);

            foreach (var entry in dic)
            {
                switch (entry.Key)
                {
                    case "command line":
                        inf.CmdUsed = entry.Value;
                        break;
                    case "machine":
                        inf.Machine = entry.Value;
                        break;
                    case "starttime":
                        inf.StartTime = DateParser(entry.Value);
                        break;
                    case "fs":
                        inf.FileSystem = entry.Value;
                        break;
                    default:
                        if (entry.Key.StartsWith("ior"))
                        {
                            inf.IOR_Version = entry.Key[4..];
                        }
                        else
                        {
                            Log.Verbose("Information token {Key} skipped!", entry.Key);
                        }
                        break;
                }
            }

            Log.Debug("Parsing Information block completed!");
            return inf;
        }

        /// <summary>
        /// Parses the Options block into a <see cref="Options"/>
        /// </summary>
        /// <param name="chunk"></param>
        /// <returns></returns>
        private static Options ParseOptions(string[] chunk)
        {
            Options opt = new Options();
            Dictionary<string, string> dic = Tokenize(chunk);

            Log.Verbose("Tokenized Options chunk: {Dic}", dic);

            foreach (var entry in dic)
            {
                switch (entry.Key)
                {
                    case "api":
                        opt.API = entry.Value;
                        break;
                    case "apiversion":
                        opt.API_Version = entry.Value;
                        break;
                    case "access":
                        opt.AccessType = entry.Value;
                        break;
                    case "type":
                        opt.TestType = entry.Value;
                        break;
                    case "ordering in a file":
                        opt.OrderInFile = entry.Value;
                        break;
                    case "tasks":
                        opt.Tasks = ushort.Parse(entry.Value);
                        break;
                    case "clients per node":
                        opt.ClientsPerNode = ushort.Parse(entry.Value);
                        break;
                    case "repetitions":
                        opt.Repetitions = ushort.Parse(entry.Value);
                        break;
                    case "xfersize":
                        opt.XferSize = entry.Value;
                        break;
                    case "blocksize":
                        opt.BlockSize = entry.Value;
                        break;
                    case "aggregate filesize":
                        opt.AggregateFileSize = entry.Value;
                        break;
                    default:
                        Log.Verbose("Options token {Key} skipped!", entry.Key);
                        continue;
                }
            }

            Log.Debug("Parsing Options block completed!");
            return opt;
        }

        /// <summary>
        /// Parses the Result block into <see cref="Result"/>
        /// </summary>
        /// <param name="chunk"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private static Result ParseResults(string[] chunk, string fileName)
        {
            Log.Verbose("Parsing Results block of {FileName}", fileName);

            Result res = new Result();
            res.Writes = new List<ResultInfo>();
            res.Reads = new List<ResultInfo>();
            res.Removes = new List<float>();

            foreach (var line in chunk)
            {
                if (line.StartsWith("Summary")) //No more information to parse.
                {
                    Log.Verbose("No information to parse.");
                    break;
                }

                bool isWrite = line.StartsWith("write");
                bool isRead = line.StartsWith("read");
                bool isRemove = line.StartsWith("remove");

                if (!(isWrite || isRead || isRemove))
                {
                    continue; //Useless information is skipped.
                }

                // Columns of data are split, removing empty rows.
                string[] tokens = line.Split(" ", StringSplitOptions.RemoveEmptyEntries);

                if (!isRemove) //If is a write or read test row.
                {
                    ResultInfo info = new ResultInfo();
                    try
                    {
                        info.Bandwidth_MiBs = float.Parse(tokens[1]);
                        info.Block_KiB = ulong.Parse(tokens[2]);
                        info.XFer_KiB = float.Parse(tokens[3]);
                        info.Open_S = double.Parse(tokens[4]);
                        info.WrRd_S = double.Parse(tokens[5]);
                        info.Close_S = double.Parse(tokens[6]);

                        info.Total = info.Open_S + info.WrRd_S + info.Close_S;
                    }
                    catch (System.Exception)
                    {
                        Log.Warning(String.Format("Cannot parse {0} of {1}\n",
                            String.IsNullOrWhiteSpace(line) ? "_" : line, fileName));
                    }
                    if (isWrite)
                    {
                        res.Writes.Add(info);
                    }
                    else
                    {
                        res.Reads.Add(info);
                    }

                }
                else
                {
                    res.Removes.Add(float.Parse(tokens[7]));
                }
            }

            Log.Verbose("Parsing Results block completed!");
            return res;
        }
    }
}
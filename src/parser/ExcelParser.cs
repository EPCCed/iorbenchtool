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
using System.Linq;
using ClosedXML.Excel;
using Serilog;

namespace parser
{
    internal class ExcelParser
    {
        /// <summary>
        /// Creates an Excel workbook.
        /// </summary>
        internal static void CreateExcel(string reportLocation)
        {
            using (var excel = new XLWorkbook())
            {
                excel.Worksheets.Add("Benchmark Test Report");

                excel.SaveAs(reportLocation);
            }

            Log.Verbose("Excel {ReportLocation} workbook created!", reportLocation);
        }

        /// <summary>
        /// Processes all the Info of IOR, creating statistics and write them into the Excel file.
        /// </summary>
        /// <param name="excel"></param>
        /// <param name="info"></param>
        internal static void PortInformation(XLWorkbook excel, Info[] info)
        {
            var excelInfo = new List<TaskInfo>(info.Length);

            Log.Verbose("Parsing statistics of {Excel}", excel);
            foreach (var inf in info) //Statistics
            {
                var tmpInfo = new TaskInfo();
                tmpInfo.ParticipantTasks = inf.Options.Tasks;

                tmpInfo.WritesMean_MiB = inf.Result.Writes.Average(r => r.Bandwidth_MiBs);
                tmpInfo.WritesStdDev = MathF.Sqrt(inf.Result.Writes.Average(r =>
                    MathF.Pow(r.Bandwidth_MiBs - tmpInfo.WritesMean_MiB, 2)));

                tmpInfo.ReadsMean_MiB = inf.Result.Reads.Average(r => r.Bandwidth_MiBs);
                tmpInfo.ReadsStdDev = MathF.Sqrt(inf.Result.Reads.Average(r =>
                    MathF.Pow(r.Bandwidth_MiBs - tmpInfo.ReadsMean_MiB, 2)));

                excelInfo.Add(tmpInfo);
            }

            excelInfo.Sort((a, b) => a.ParticipantTasks.CompareTo(b.ParticipantTasks));

            Log.Verbose("Populating {Excel} workbook with parsed data!", excel);
            var ws = excel.Worksheets.First();

            // Static part of the Excel file where the data is written.
            ws.Cell("A25").Value = "Participant Tasks";
            ws.Cell("B25").Value = "Writes Bandwidth mean";
            ws.Cell("C25").Value = "Writes Bandwidth StdDev";
            ws.Cell("D25").Value = "Reads Bandwidth mean";
            ws.Cell("E25").Value = "Reads Bandwidth StdDev";

            // Data can have more than 64 participant tasks, but GeneratedCode/GeneratedClass.cs must
            // be modifed.
            int extra = excelInfo.Count < 6 ? (7 - excelInfo.Count) : excelInfo.Count;

            if (excelInfo.Count < 6)
            {
                for (int i = 0; i < extra; i++)
                {
                    excelInfo.Add(new TaskInfo
                    {
                        ParticipantTasks = 0,
                        ReadsMean_MiB = 0,
                        ReadsStdDev = 0,
                        WritesMean_MiB = 0,
                        WritesStdDev = 0
                    });
                }
            }

            for (int i = 26; i < 26 + excelInfo.Count; i++)
            {
                var taskInfo = excelInfo[i - 26];
                Log.Verbose("{Excel} TaskInfo[{I}]: {TaskInfo}", excel, -26, ObjectDumper.Dump(taskInfo));

                ws.Cell($"A{i}").Value = taskInfo.ParticipantTasks;
                ws.Cell($"B{i}").Value = taskInfo.WritesMean_MiB;
                ws.Cell($"C{i}").Value = taskInfo.WritesStdDev;
                ws.Cell($"D{i}").Value = taskInfo.ReadsMean_MiB;
                ws.Cell($"E{i}").Value = taskInfo.ReadsStdDev;
            }

            Log.Verbose("Excel {Excel} populated!", excel);
        }

        /// <summary>
        /// Data that is printed into the Excel. This object must be modified to add new data into the
        /// Excel file. 
        /// </summary>
        private class TaskInfo
        {
            public ushort ParticipantTasks { get; set; }
            public float WritesMean_MiB { get; set; }
            public float WritesStdDev { get; set; }
            public float ReadsMean_MiB { get; set; }
            public float ReadsStdDev { get; set; }
        }
    }
}
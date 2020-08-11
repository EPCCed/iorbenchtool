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

using DocumentFormat.OpenXml.Packaging;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C14 = DocumentFormat.OpenXml.Office2010.Drawing.Charts;
using Cs = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Spreadsheet;

namespace GeneratedCode
{
    internal class GeneratedClass : parser.IGenerated
    {
        private string Sheet;

        // Adds child parts and generates content of the specified part.
        public void CreateWorksheetPart(WorksheetPart part, string sheet)
        {
            Sheet = sheet;

            DrawingsPart drawingsPart1 = part.AddNewPart<DrawingsPart>("rId1");
            GenerateDrawingsPart1Content(drawingsPart1);

            ChartPart chartPart1 = drawingsPart1.AddNewPart<ChartPart>("rId1");
            GenerateChartPart1Content(chartPart1);

            ChartColorStylePart chartColorStylePart1 = chartPart1.AddNewPart<ChartColorStylePart>("rId2");
            GenerateChartColorStylePart1Content(chartColorStylePart1);

            ChartStylePart chartStylePart1 = chartPart1.AddNewPart<ChartStylePart>("rId1");
            GenerateChartStylePart1Content(chartStylePart1);

            GeneratePartContent(part);

        }

        // Generates content of drawingsPart1.
        private void GenerateDrawingsPart1Content(DrawingsPart drawingsPart1)
        {
            Xdr.WorksheetDrawing worksheetDrawing1 = new Xdr.WorksheetDrawing();
            worksheetDrawing1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheetDrawing1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            Xdr.TwoCellAnchor twoCellAnchor1 = new Xdr.TwoCellAnchor();

            Xdr.FromMarker fromMarker1 = new Xdr.FromMarker();
            Xdr.ColumnId columnId1 = new Xdr.ColumnId();
            columnId1.Text = "0";
            Xdr.ColumnOffset columnOffset1 = new Xdr.ColumnOffset();
            columnOffset1.Text = "0";
            Xdr.RowId rowId1 = new Xdr.RowId();
            rowId1.Text = "0";
            Xdr.RowOffset rowOffset1 = new Xdr.RowOffset();
            rowOffset1.Text = "0";

            fromMarker1.Append(columnId1);
            fromMarker1.Append(columnOffset1);
            fromMarker1.Append(rowId1);
            fromMarker1.Append(rowOffset1);

            Xdr.ToMarker toMarker1 = new Xdr.ToMarker();
            Xdr.ColumnId columnId2 = new Xdr.ColumnId();
            columnId2.Text = "9";
            Xdr.ColumnOffset columnOffset2 = new Xdr.ColumnOffset();
            columnOffset2.Text = "653142";
            Xdr.RowId rowId2 = new Xdr.RowId();
            rowId2.Text = "21";
            Xdr.RowOffset rowOffset2 = new Xdr.RowOffset();
            rowOffset2.Text = "185056";

            toMarker1.Append(columnId2);
            toMarker1.Append(columnOffset2);
            toMarker1.Append(rowId2);
            toMarker1.Append(rowOffset2);

            Xdr.GraphicFrame graphicFrame1 = new Xdr.GraphicFrame() { Macro = "" };

            Xdr.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties1 = new Xdr.NonVisualGraphicFrameProperties();

            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Chart 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{81C387B4-4DBD-4247-ABCF-0D35610C4CDD}\" />");

            nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);

            nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

            nonVisualDrawingProperties1.Append(nonVisualDrawingPropertiesExtensionList1);

            Xdr.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Xdr.NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            nonVisualGraphicFrameProperties1.Append(nonVisualDrawingProperties1);
            nonVisualGraphicFrameProperties1.Append(nonVisualGraphicFrameDrawingProperties1);

            Xdr.Transform transform1 = new Xdr.Transform();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 0L, Cy = 0L };

            transform1.Append(offset1);
            transform1.Append(extents1);

            A.Graphic graphic1 = new A.Graphic();

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };

            C.ChartReference chartReference1 = new C.ChartReference() { Id = "rId1" };
            chartReference1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartReference1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            graphicData1.Append(chartReference1);

            graphic1.Append(graphicData1);

            graphicFrame1.Append(nonVisualGraphicFrameProperties1);
            graphicFrame1.Append(transform1);
            graphicFrame1.Append(graphic1);
            Xdr.ClientData clientData1 = new Xdr.ClientData();

            twoCellAnchor1.Append(fromMarker1);
            twoCellAnchor1.Append(toMarker1);
            twoCellAnchor1.Append(graphicFrame1);
            twoCellAnchor1.Append(clientData1);

            worksheetDrawing1.Append(twoCellAnchor1);

            drawingsPart1.WorksheetDrawing = worksheetDrawing1;
        }

        // Generates content of chartPart1.
        private void GenerateChartPart1Content(ChartPart chartPart1)
        {
            C.ChartSpace chartSpace1 = new C.ChartSpace();
            chartSpace1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartSpace1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            chartSpace1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            chartSpace1.AddNamespaceDeclaration("c16r2", "http://schemas.microsoft.com/office/drawing/2015/06/chart");
            C.Date1904 date19041 = new C.Date1904() { Val = false };
            C.EditingLanguage editingLanguage1 = new C.EditingLanguage() { Val = "en-US" };
            C.RoundedCorners roundedCorners1 = new C.RoundedCorners() { Val = false };

            AlternateContent alternateContent1 = new AlternateContent();
            alternateContent1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

            AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "c14" };
            alternateContentChoice1.AddNamespaceDeclaration("c14", "http://schemas.microsoft.com/office/drawing/2007/8/2/chart");
            C14.Style style1 = new C14.Style() { Val = 102 };

            alternateContentChoice1.Append(style1);

            AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();
            C.Style style2 = new C.Style() { Val = 2 };

            alternateContentFallback1.Append(style2);

            alternateContent1.Append(alternateContentChoice1);
            alternateContent1.Append(alternateContentFallback1);

            C.Chart chart1 = new C.Chart();

            C.Title title1 = new C.Title();

            C.ChartText chartText1 = new C.ChartText();

            C.RichText richText1 = new C.RichText();
            A.BodyProperties bodyProperties1 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle1 = new A.ListStyle();

            A.Paragraph paragraph1 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties1 = new A.DefaultRunProperties() { FontSize = 1400, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Spacing = 0, Baseline = 0 };

            A.SolidFill solidFill1 = new A.SolidFill();

            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset1 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor1.Append(luminanceModulation1);
            schemeColor1.Append(luminanceOffset1);

            solidFill1.Append(schemeColor1);
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties1.Append(solidFill1);
            defaultRunProperties1.Append(latinFont1);
            defaultRunProperties1.Append(eastAsianFont1);
            defaultRunProperties1.Append(complexScriptFont1);

            paragraphProperties1.Append(defaultRunProperties1);

            A.Run run1 = new A.Run();
            A.RunProperties runProperties1 = new A.RunProperties() { Language = "en-GB" };
            A.Text text1 = new A.Text();
            text1.Text = "Bandwidth mean vs";

            run1.Append(runProperties1);
            run1.Append(text1);

            A.Run run2 = new A.Run();
            A.RunProperties runProperties2 = new A.RunProperties() { Language = "en-GB", Baseline = 0 };
            A.Text text2 = new A.Text();
            text2.Text = " Participating tasks";

            run2.Append(runProperties2);
            run2.Append(text2);
            A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties() { Language = "en-GB" };

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);
            paragraph1.Append(run2);
            paragraph1.Append(endParagraphRunProperties1);

            richText1.Append(bodyProperties1);
            richText1.Append(listStyle1);
            richText1.Append(paragraph1);

            chartText1.Append(richText1);
            C.Overlay overlay1 = new C.Overlay() { Val = false };

            C.ChartShapeProperties chartShapeProperties1 = new C.ChartShapeProperties();
            A.NoFill noFill1 = new A.NoFill();

            A.Outline outline1 = new A.Outline();
            A.NoFill noFill2 = new A.NoFill();

            outline1.Append(noFill2);
            A.EffectList effectList1 = new A.EffectList();

            chartShapeProperties1.Append(noFill1);
            chartShapeProperties1.Append(outline1);
            chartShapeProperties1.Append(effectList1);

            C.TextProperties textProperties1 = new C.TextProperties();
            A.BodyProperties bodyProperties2 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle2 = new A.ListStyle();

            A.Paragraph paragraph2 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties2 = new A.DefaultRunProperties() { FontSize = 1400, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Spacing = 0, Baseline = 0 };

            A.SolidFill solidFill2 = new A.SolidFill();

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset2 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor2.Append(luminanceModulation2);
            schemeColor2.Append(luminanceOffset2);

            solidFill2.Append(schemeColor2);
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties2.Append(solidFill2);
            defaultRunProperties2.Append(latinFont2);
            defaultRunProperties2.Append(eastAsianFont2);
            defaultRunProperties2.Append(complexScriptFont2);

            paragraphProperties2.Append(defaultRunProperties2);
            A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(endParagraphRunProperties2);

            textProperties1.Append(bodyProperties2);
            textProperties1.Append(listStyle2);
            textProperties1.Append(paragraph2);

            title1.Append(chartText1);
            title1.Append(overlay1);
            title1.Append(chartShapeProperties1);
            title1.Append(textProperties1);
            C.AutoTitleDeleted autoTitleDeleted1 = new C.AutoTitleDeleted() { Val = false };

            C.PlotArea plotArea1 = new C.PlotArea();
            C.Layout layout1 = new C.Layout();

            C.BarChart barChart1 = new C.BarChart();
            C.BarDirection barDirection1 = new C.BarDirection() { Val = C.BarDirectionValues.Column };
            C.BarGrouping barGrouping1 = new C.BarGrouping() { Val = C.BarGroupingValues.Clustered };
            C.VaryColors varyColors1 = new C.VaryColors() { Val = false };

            C.BarChartSeries barChartSeries1 = new C.BarChartSeries();
            C.Index index1 = new C.Index() { Val = (UInt32Value)0U };
            C.Order order1 = new C.Order() { Val = (UInt32Value)0U };

            C.SeriesText seriesText1 = new C.SeriesText();

            C.StringReference stringReference1 = new C.StringReference();
            C.Formula formula1 = new C.Formula();
            formula1.Text = $"\'{Sheet}\'!$B$25";

            C.StringCache stringCache1 = new C.StringCache();
            C.PointCount pointCount1 = new C.PointCount() { Val = (UInt32Value)1U };

            C.StringPoint stringPoint1 = new C.StringPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue1 = new C.NumericValue();
            numericValue1.Text = "Writes Bandwidth mean";

            stringPoint1.Append(numericValue1);

            stringCache1.Append(pointCount1);
            stringCache1.Append(stringPoint1);

            stringReference1.Append(formula1);
            stringReference1.Append(stringCache1);

            seriesText1.Append(stringReference1);

            C.ChartShapeProperties chartShapeProperties2 = new C.ChartShapeProperties();

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            solidFill3.Append(schemeColor3);

            A.Outline outline2 = new A.Outline();
            A.NoFill noFill3 = new A.NoFill();

            outline2.Append(noFill3);
            A.EffectList effectList2 = new A.EffectList();

            chartShapeProperties2.Append(solidFill3);
            chartShapeProperties2.Append(outline2);
            chartShapeProperties2.Append(effectList2);
            C.InvertIfNegative invertIfNegative1 = new C.InvertIfNegative() { Val = false };

            C.ErrorBars errorBars1 = new C.ErrorBars();
            C.ErrorBarType errorBarType1 = new C.ErrorBarType() { Val = C.ErrorBarValues.Both };
            C.ErrorBarValueType errorBarValueType1 = new C.ErrorBarValueType() { Val = C.ErrorValues.Custom };
            C.NoEndCap noEndCap1 = new C.NoEndCap() { Val = false };

            C.Plus plus1 = new C.Plus();

            C.NumberReference numberReference1 = new C.NumberReference();
            C.Formula formula2 = new C.Formula();
            formula2.Text = $"\'{Sheet}\'!$C$26:$C$32";

            C.NumberingCache numberingCache1 = new C.NumberingCache();
            C.FormatCode formatCode1 = new C.FormatCode();
            formatCode1.Text = "General";
            C.PointCount pointCount2 = new C.PointCount() { Val = (UInt32Value)7U };

            C.NumericPoint numericPoint1 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue2 = new C.NumericValue();
            numericValue2.Text = "18.26923";

            numericPoint1.Append(numericValue2);

            C.NumericPoint numericPoint2 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue3 = new C.NumericValue();
            numericValue3.Text = "161.9091";

            numericPoint2.Append(numericValue3);

            C.NumericPoint numericPoint3 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue4 = new C.NumericValue();
            numericValue4.Text = "173.05860000000001";

            numericPoint3.Append(numericValue4);

            C.NumericPoint numericPoint4 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue5 = new C.NumericValue();
            numericValue5.Text = "122.4585";

            numericPoint4.Append(numericValue5);

            C.NumericPoint numericPoint5 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue6 = new C.NumericValue();
            numericValue6.Text = "118.0836";

            numericPoint5.Append(numericValue6);

            C.NumericPoint numericPoint6 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue7 = new C.NumericValue();
            numericValue7.Text = "9.2654160000000001";

            numericPoint6.Append(numericValue7);

            C.NumericPoint numericPoint7 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue8 = new C.NumericValue();
            numericValue8.Text = "7.8593120000000001";

            numericPoint7.Append(numericValue8);

            numberingCache1.Append(formatCode1);
            numberingCache1.Append(pointCount2);
            numberingCache1.Append(numericPoint1);
            numberingCache1.Append(numericPoint2);
            numberingCache1.Append(numericPoint3);
            numberingCache1.Append(numericPoint4);
            numberingCache1.Append(numericPoint5);
            numberingCache1.Append(numericPoint6);
            numberingCache1.Append(numericPoint7);

            numberReference1.Append(formula2);
            numberReference1.Append(numberingCache1);

            plus1.Append(numberReference1);

            C.Minus minus1 = new C.Minus();

            C.NumberReference numberReference2 = new C.NumberReference();
            C.Formula formula3 = new C.Formula();
            formula3.Text = $"\'{Sheet}\'!$C$26:$C$32";

            C.NumberingCache numberingCache2 = new C.NumberingCache();
            C.FormatCode formatCode2 = new C.FormatCode();
            formatCode2.Text = "General";
            C.PointCount pointCount3 = new C.PointCount() { Val = (UInt32Value)7U };

            C.NumericPoint numericPoint8 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue9 = new C.NumericValue();
            numericValue9.Text = "18.26923";

            numericPoint8.Append(numericValue9);

            C.NumericPoint numericPoint9 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue10 = new C.NumericValue();
            numericValue10.Text = "161.9091";

            numericPoint9.Append(numericValue10);

            C.NumericPoint numericPoint10 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue11 = new C.NumericValue();
            numericValue11.Text = "173.05860000000001";

            numericPoint10.Append(numericValue11);

            C.NumericPoint numericPoint11 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue12 = new C.NumericValue();
            numericValue12.Text = "122.4585";

            numericPoint11.Append(numericValue12);

            C.NumericPoint numericPoint12 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue13 = new C.NumericValue();
            numericValue13.Text = "118.0836";

            numericPoint12.Append(numericValue13);

            C.NumericPoint numericPoint13 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue14 = new C.NumericValue();
            numericValue14.Text = "9.2654160000000001";

            numericPoint13.Append(numericValue14);

            C.NumericPoint numericPoint14 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue15 = new C.NumericValue();
            numericValue15.Text = "7.8593120000000001";

            numericPoint14.Append(numericValue15);

            numberingCache2.Append(formatCode2);
            numberingCache2.Append(pointCount3);
            numberingCache2.Append(numericPoint8);
            numberingCache2.Append(numericPoint9);
            numberingCache2.Append(numericPoint10);
            numberingCache2.Append(numericPoint11);
            numberingCache2.Append(numericPoint12);
            numberingCache2.Append(numericPoint13);
            numberingCache2.Append(numericPoint14);

            numberReference2.Append(formula3);
            numberReference2.Append(numberingCache2);

            minus1.Append(numberReference2);

            C.ChartShapeProperties chartShapeProperties3 = new C.ChartShapeProperties();
            A.NoFill noFill4 = new A.NoFill();

            A.Outline outline3 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset3 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor4.Append(luminanceModulation3);
            schemeColor4.Append(luminanceOffset3);

            solidFill4.Append(schemeColor4);
            A.Round round1 = new A.Round();

            outline3.Append(solidFill4);
            outline3.Append(round1);
            A.EffectList effectList3 = new A.EffectList();

            chartShapeProperties3.Append(noFill4);
            chartShapeProperties3.Append(outline3);
            chartShapeProperties3.Append(effectList3);

            errorBars1.Append(errorBarType1);
            errorBars1.Append(errorBarValueType1);
            errorBars1.Append(noEndCap1);
            errorBars1.Append(plus1);
            errorBars1.Append(minus1);
            errorBars1.Append(chartShapeProperties3);

            C.CategoryAxisData categoryAxisData1 = new C.CategoryAxisData();

            C.NumberReference numberReference3 = new C.NumberReference();
            C.Formula formula4 = new C.Formula();
            formula4.Text = $"\'{Sheet}\'!$A$26:$A$32";

            C.NumberingCache numberingCache3 = new C.NumberingCache();
            C.FormatCode formatCode3 = new C.FormatCode();
            formatCode3.Text = "General";
            C.PointCount pointCount4 = new C.PointCount() { Val = (UInt32Value)7U };

            C.NumericPoint numericPoint15 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue16 = new C.NumericValue();
            numericValue16.Text = "1";

            numericPoint15.Append(numericValue16);

            C.NumericPoint numericPoint16 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue17 = new C.NumericValue();
            numericValue17.Text = "2";

            numericPoint16.Append(numericValue17);

            C.NumericPoint numericPoint17 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue18 = new C.NumericValue();
            numericValue18.Text = "4";

            numericPoint17.Append(numericValue18);

            C.NumericPoint numericPoint18 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue19 = new C.NumericValue();
            numericValue19.Text = "8";

            numericPoint18.Append(numericValue19);

            C.NumericPoint numericPoint19 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue20 = new C.NumericValue();
            numericValue20.Text = "16";

            numericPoint19.Append(numericValue20);

            C.NumericPoint numericPoint20 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue21 = new C.NumericValue();
            numericValue21.Text = "32";

            numericPoint20.Append(numericValue21);

            C.NumericPoint numericPoint21 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue22 = new C.NumericValue();
            numericValue22.Text = "64";

            numericPoint21.Append(numericValue22);

            numberingCache3.Append(formatCode3);
            numberingCache3.Append(pointCount4);
            numberingCache3.Append(numericPoint15);
            numberingCache3.Append(numericPoint16);
            numberingCache3.Append(numericPoint17);
            numberingCache3.Append(numericPoint18);
            numberingCache3.Append(numericPoint19);
            numberingCache3.Append(numericPoint20);
            numberingCache3.Append(numericPoint21);

            numberReference3.Append(formula4);
            numberReference3.Append(numberingCache3);

            categoryAxisData1.Append(numberReference3);

            C.Values values1 = new C.Values();

            C.NumberReference numberReference4 = new C.NumberReference();
            C.Formula formula5 = new C.Formula();
            formula5.Text = $"\'{Sheet}\'!$B$26:$B$32";

            C.NumberingCache numberingCache4 = new C.NumberingCache();
            C.FormatCode formatCode4 = new C.FormatCode();
            formatCode4.Text = "General";
            C.PointCount pointCount5 = new C.PointCount() { Val = (UInt32Value)7U };

            C.NumericPoint numericPoint22 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue23 = new C.NumericValue();
            numericValue23.Text = "750.04499999999996";

            numericPoint22.Append(numericValue23);

            C.NumericPoint numericPoint23 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue24 = new C.NumericValue();
            numericValue24.Text = "686.71";

            numericPoint23.Append(numericValue24);

            C.NumericPoint numericPoint24 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue25 = new C.NumericValue();
            numericValue25.Text = "690.04499999999996";

            numericPoint24.Append(numericValue25);

            C.NumericPoint numericPoint25 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue26 = new C.NumericValue();
            numericValue26.Text = "609.10749999999996";

            numericPoint25.Append(numericValue26);

            C.NumericPoint numericPoint26 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue27 = new C.NumericValue();
            numericValue27.Text = "585.09749999999997";

            numericPoint26.Append(numericValue27);

            C.NumericPoint numericPoint27 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue28 = new C.NumericValue();
            numericValue28.Text = "426.45499999999998";

            numericPoint27.Append(numericValue28);

            C.NumericPoint numericPoint28 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue29 = new C.NumericValue();
            numericValue29.Text = "290.10250000000002";

            numericPoint28.Append(numericValue29);

            numberingCache4.Append(formatCode4);
            numberingCache4.Append(pointCount5);
            numberingCache4.Append(numericPoint22);
            numberingCache4.Append(numericPoint23);
            numberingCache4.Append(numericPoint24);
            numberingCache4.Append(numericPoint25);
            numberingCache4.Append(numericPoint26);
            numberingCache4.Append(numericPoint27);
            numberingCache4.Append(numericPoint28);

            numberReference4.Append(formula5);
            numberReference4.Append(numberingCache4);

            values1.Append(numberReference4);

            C.BarSerExtensionList barSerExtensionList1 = new C.BarSerExtensionList();

            C.BarSerExtension barSerExtension1 = new C.BarSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
            barSerExtension1.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

            OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000000-496A-4610-8F22-94E58C7150D5}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

            barSerExtension1.Append(openXmlUnknownElement2);

            barSerExtensionList1.Append(barSerExtension1);

            barChartSeries1.Append(index1);
            barChartSeries1.Append(order1);
            barChartSeries1.Append(seriesText1);
            barChartSeries1.Append(chartShapeProperties2);
            barChartSeries1.Append(invertIfNegative1);
            barChartSeries1.Append(errorBars1);
            barChartSeries1.Append(categoryAxisData1);
            barChartSeries1.Append(values1);
            barChartSeries1.Append(barSerExtensionList1);

            C.BarChartSeries barChartSeries2 = new C.BarChartSeries();
            C.Index index2 = new C.Index() { Val = (UInt32Value)2U };
            C.Order order2 = new C.Order() { Val = (UInt32Value)1U };

            C.SeriesText seriesText2 = new C.SeriesText();

            C.StringReference stringReference2 = new C.StringReference();
            C.Formula formula6 = new C.Formula();
            formula6.Text = $"\'{Sheet}\'!$D$25";

            C.StringCache stringCache2 = new C.StringCache();
            C.PointCount pointCount6 = new C.PointCount() { Val = (UInt32Value)1U };

            C.StringPoint stringPoint2 = new C.StringPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue30 = new C.NumericValue();
            numericValue30.Text = "Reads Bandwidth mean";

            stringPoint2.Append(numericValue30);

            stringCache2.Append(pointCount6);
            stringCache2.Append(stringPoint2);

            stringReference2.Append(formula6);
            stringReference2.Append(stringCache2);

            seriesText2.Append(stringReference2);

            C.ChartShapeProperties chartShapeProperties4 = new C.ChartShapeProperties();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent3 };

            solidFill5.Append(schemeColor5);

            A.Outline outline4 = new A.Outline();
            A.NoFill noFill5 = new A.NoFill();

            outline4.Append(noFill5);
            A.EffectList effectList4 = new A.EffectList();

            chartShapeProperties4.Append(solidFill5);
            chartShapeProperties4.Append(outline4);
            chartShapeProperties4.Append(effectList4);
            C.InvertIfNegative invertIfNegative2 = new C.InvertIfNegative() { Val = false };

            C.ErrorBars errorBars2 = new C.ErrorBars();
            C.ErrorBarType errorBarType2 = new C.ErrorBarType() { Val = C.ErrorBarValues.Both };
            C.ErrorBarValueType errorBarValueType2 = new C.ErrorBarValueType() { Val = C.ErrorValues.Custom };
            C.NoEndCap noEndCap2 = new C.NoEndCap() { Val = false };

            C.Plus plus2 = new C.Plus();

            C.NumberReference numberReference5 = new C.NumberReference();
            C.Formula formula7 = new C.Formula();
            formula7.Text = $"\'{Sheet}\'!$E$26:$E$32";

            C.NumberingCache numberingCache5 = new C.NumberingCache();
            C.FormatCode formatCode5 = new C.FormatCode();
            formatCode5.Text = "General";
            C.PointCount pointCount7 = new C.PointCount() { Val = (UInt32Value)7U };

            C.NumericPoint numericPoint29 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue31 = new C.NumericValue();
            numericValue31.Text = "4.7214410000000004";

            numericPoint29.Append(numericValue31);

            C.NumericPoint numericPoint30 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue32 = new C.NumericValue();
            numericValue32.Text = "10.775539999999999";

            numericPoint30.Append(numericValue32);

            C.NumericPoint numericPoint31 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue33 = new C.NumericValue();
            numericValue33.Text = "214.67490000000001";

            numericPoint31.Append(numericValue33);

            C.NumericPoint numericPoint32 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue34 = new C.NumericValue();
            numericValue34.Text = "212.03030000000001";

            numericPoint32.Append(numericValue34);

            C.NumericPoint numericPoint33 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue35 = new C.NumericValue();
            numericValue35.Text = "77.603099999999998";

            numericPoint33.Append(numericValue35);

            C.NumericPoint numericPoint34 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue36 = new C.NumericValue();
            numericValue36.Text = "58.671860000000002";

            numericPoint34.Append(numericValue36);

            C.NumericPoint numericPoint35 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue37 = new C.NumericValue();
            numericValue37.Text = "62.321089999999998";

            numericPoint35.Append(numericValue37);

            numberingCache5.Append(formatCode5);
            numberingCache5.Append(pointCount7);
            numberingCache5.Append(numericPoint29);
            numberingCache5.Append(numericPoint30);
            numberingCache5.Append(numericPoint31);
            numberingCache5.Append(numericPoint32);
            numberingCache5.Append(numericPoint33);
            numberingCache5.Append(numericPoint34);
            numberingCache5.Append(numericPoint35);

            numberReference5.Append(formula7);
            numberReference5.Append(numberingCache5);

            plus2.Append(numberReference5);

            C.Minus minus2 = new C.Minus();

            C.NumberReference numberReference6 = new C.NumberReference();
            C.Formula formula8 = new C.Formula();
            formula8.Text = $"\'{Sheet}\'!$E$26:$E$32";

            C.NumberingCache numberingCache6 = new C.NumberingCache();
            C.FormatCode formatCode6 = new C.FormatCode();
            formatCode6.Text = "General";
            C.PointCount pointCount8 = new C.PointCount() { Val = (UInt32Value)7U };

            C.NumericPoint numericPoint36 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue38 = new C.NumericValue();
            numericValue38.Text = "4.7214410000000004";

            numericPoint36.Append(numericValue38);

            C.NumericPoint numericPoint37 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue39 = new C.NumericValue();
            numericValue39.Text = "10.775539999999999";

            numericPoint37.Append(numericValue39);

            C.NumericPoint numericPoint38 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue40 = new C.NumericValue();
            numericValue40.Text = "214.67490000000001";

            numericPoint38.Append(numericValue40);

            C.NumericPoint numericPoint39 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue41 = new C.NumericValue();
            numericValue41.Text = "212.03030000000001";

            numericPoint39.Append(numericValue41);

            C.NumericPoint numericPoint40 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue42 = new C.NumericValue();
            numericValue42.Text = "77.603099999999998";

            numericPoint40.Append(numericValue42);

            C.NumericPoint numericPoint41 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue43 = new C.NumericValue();
            numericValue43.Text = "58.671860000000002";

            numericPoint41.Append(numericValue43);

            C.NumericPoint numericPoint42 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue44 = new C.NumericValue();
            numericValue44.Text = "62.321089999999998";

            numericPoint42.Append(numericValue44);

            numberingCache6.Append(formatCode6);
            numberingCache6.Append(pointCount8);
            numberingCache6.Append(numericPoint36);
            numberingCache6.Append(numericPoint37);
            numberingCache6.Append(numericPoint38);
            numberingCache6.Append(numericPoint39);
            numberingCache6.Append(numericPoint40);
            numberingCache6.Append(numericPoint41);
            numberingCache6.Append(numericPoint42);

            numberReference6.Append(formula8);
            numberReference6.Append(numberingCache6);

            minus2.Append(numberReference6);

            C.ChartShapeProperties chartShapeProperties5 = new C.ChartShapeProperties();
            A.NoFill noFill6 = new A.NoFill();

            A.Outline outline5 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill6 = new A.SolidFill();

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset4 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor6.Append(luminanceModulation4);
            schemeColor6.Append(luminanceOffset4);

            solidFill6.Append(schemeColor6);
            A.Round round2 = new A.Round();

            outline5.Append(solidFill6);
            outline5.Append(round2);
            A.EffectList effectList5 = new A.EffectList();

            chartShapeProperties5.Append(noFill6);
            chartShapeProperties5.Append(outline5);
            chartShapeProperties5.Append(effectList5);

            errorBars2.Append(errorBarType2);
            errorBars2.Append(errorBarValueType2);
            errorBars2.Append(noEndCap2);
            errorBars2.Append(plus2);
            errorBars2.Append(minus2);
            errorBars2.Append(chartShapeProperties5);

            C.CategoryAxisData categoryAxisData2 = new C.CategoryAxisData();

            C.NumberReference numberReference7 = new C.NumberReference();
            C.Formula formula9 = new C.Formula();
            formula9.Text = $"\'{Sheet}\'!$A$26:$A$32";

            C.NumberingCache numberingCache7 = new C.NumberingCache();
            C.FormatCode formatCode7 = new C.FormatCode();
            formatCode7.Text = "General";
            C.PointCount pointCount9 = new C.PointCount() { Val = (UInt32Value)7U };

            C.NumericPoint numericPoint43 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue45 = new C.NumericValue();
            numericValue45.Text = "1";

            numericPoint43.Append(numericValue45);

            C.NumericPoint numericPoint44 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue46 = new C.NumericValue();
            numericValue46.Text = "2";

            numericPoint44.Append(numericValue46);

            C.NumericPoint numericPoint45 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue47 = new C.NumericValue();
            numericValue47.Text = "4";

            numericPoint45.Append(numericValue47);

            C.NumericPoint numericPoint46 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue48 = new C.NumericValue();
            numericValue48.Text = "8";

            numericPoint46.Append(numericValue48);

            C.NumericPoint numericPoint47 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue49 = new C.NumericValue();
            numericValue49.Text = "16";

            numericPoint47.Append(numericValue49);

            C.NumericPoint numericPoint48 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue50 = new C.NumericValue();
            numericValue50.Text = "32";

            numericPoint48.Append(numericValue50);

            C.NumericPoint numericPoint49 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue51 = new C.NumericValue();
            numericValue51.Text = "64";

            numericPoint49.Append(numericValue51);

            numberingCache7.Append(formatCode7);
            numberingCache7.Append(pointCount9);
            numberingCache7.Append(numericPoint43);
            numberingCache7.Append(numericPoint44);
            numberingCache7.Append(numericPoint45);
            numberingCache7.Append(numericPoint46);
            numberingCache7.Append(numericPoint47);
            numberingCache7.Append(numericPoint48);
            numberingCache7.Append(numericPoint49);

            numberReference7.Append(formula9);
            numberReference7.Append(numberingCache7);

            categoryAxisData2.Append(numberReference7);

            C.Values values2 = new C.Values();

            C.NumberReference numberReference8 = new C.NumberReference();
            C.Formula formula10 = new C.Formula();
            formula10.Text = $"\'{Sheet}\'!$D$26:$D$32";

            C.NumberingCache numberingCache8 = new C.NumberingCache();
            C.FormatCode formatCode8 = new C.FormatCode();
            formatCode8.Text = "General";
            C.PointCount pointCount10 = new C.PointCount() { Val = (UInt32Value)7U };

            C.NumericPoint numericPoint50 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue52 = new C.NumericValue();
            numericValue52.Text = "1075.24";

            numericPoint50.Append(numericValue52);

            C.NumericPoint numericPoint51 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue53 = new C.NumericValue();
            numericValue53.Text = "1144.6420000000001";

            numericPoint51.Append(numericValue53);

            C.NumericPoint numericPoint52 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue54 = new C.NumericValue();
            numericValue54.Text = "1745.7850000000001";

            numericPoint52.Append(numericValue54);

            C.NumericPoint numericPoint53 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue55 = new C.NumericValue();
            numericValue55.Text = "1865.405";

            numericPoint53.Append(numericValue55);

            C.NumericPoint numericPoint54 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue56 = new C.NumericValue();
            numericValue56.Text = "1787.857";

            numericPoint54.Append(numericValue56);

            C.NumericPoint numericPoint55 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue57 = new C.NumericValue();
            numericValue57.Text = "1678.998";

            numericPoint55.Append(numericValue57);

            C.NumericPoint numericPoint56 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue58 = new C.NumericValue();
            numericValue58.Text = "1121.4349999999999";

            numericPoint56.Append(numericValue58);

            numberingCache8.Append(formatCode8);
            numberingCache8.Append(pointCount10);
            numberingCache8.Append(numericPoint50);
            numberingCache8.Append(numericPoint51);
            numberingCache8.Append(numericPoint52);
            numberingCache8.Append(numericPoint53);
            numberingCache8.Append(numericPoint54);
            numberingCache8.Append(numericPoint55);
            numberingCache8.Append(numericPoint56);

            numberReference8.Append(formula10);
            numberReference8.Append(numberingCache8);

            values2.Append(numberReference8);

            C.BarSerExtensionList barSerExtensionList2 = new C.BarSerExtensionList();

            C.BarSerExtension barSerExtension2 = new C.BarSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
            barSerExtension2.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

            OpenXmlUnknownElement openXmlUnknownElement3 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000001-496A-4610-8F22-94E58C7150D5}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

            barSerExtension2.Append(openXmlUnknownElement3);

            barSerExtensionList2.Append(barSerExtension2);

            barChartSeries2.Append(index2);
            barChartSeries2.Append(order2);
            barChartSeries2.Append(seriesText2);
            barChartSeries2.Append(chartShapeProperties4);
            barChartSeries2.Append(invertIfNegative2);
            barChartSeries2.Append(errorBars2);
            barChartSeries2.Append(categoryAxisData2);
            barChartSeries2.Append(values2);
            barChartSeries2.Append(barSerExtensionList2);

            C.DataLabels dataLabels1 = new C.DataLabels();
            C.ShowLegendKey showLegendKey1 = new C.ShowLegendKey() { Val = false };
            C.ShowValue showValue1 = new C.ShowValue() { Val = false };
            C.ShowCategoryName showCategoryName1 = new C.ShowCategoryName() { Val = false };
            C.ShowSeriesName showSeriesName1 = new C.ShowSeriesName() { Val = false };
            C.ShowPercent showPercent1 = new C.ShowPercent() { Val = false };
            C.ShowBubbleSize showBubbleSize1 = new C.ShowBubbleSize() { Val = false };

            dataLabels1.Append(showLegendKey1);
            dataLabels1.Append(showValue1);
            dataLabels1.Append(showCategoryName1);
            dataLabels1.Append(showSeriesName1);
            dataLabels1.Append(showPercent1);
            dataLabels1.Append(showBubbleSize1);
            C.GapWidth gapWidth1 = new C.GapWidth() { Val = (UInt16Value)219U };
            C.Overlap overlap1 = new C.Overlap() { Val = -27 };
            C.AxisId axisId1 = new C.AxisId() { Val = (UInt32Value)1191528143U };
            C.AxisId axisId2 = new C.AxisId() { Val = (UInt32Value)1189680815U };

            barChart1.Append(barDirection1);
            barChart1.Append(barGrouping1);
            barChart1.Append(varyColors1);
            barChart1.Append(barChartSeries1);
            barChart1.Append(barChartSeries2);
            barChart1.Append(dataLabels1);
            barChart1.Append(gapWidth1);
            barChart1.Append(overlap1);
            barChart1.Append(axisId1);
            barChart1.Append(axisId2);

            C.CategoryAxis categoryAxis1 = new C.CategoryAxis();
            C.AxisId axisId3 = new C.AxisId() { Val = (UInt32Value)1191528143U };

            C.Scaling scaling1 = new C.Scaling();
            C.Orientation orientation1 = new C.Orientation() { Val = C.OrientationValues.MinMax };

            scaling1.Append(orientation1);
            C.Delete delete1 = new C.Delete() { Val = false };
            C.AxisPosition axisPosition1 = new C.AxisPosition() { Val = C.AxisPositionValues.Bottom };

            C.Title title2 = new C.Title();

            C.ChartText chartText2 = new C.ChartText();

            C.RichText richText2 = new C.RichText();
            A.BodyProperties bodyProperties3 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle3 = new A.ListStyle();

            A.Paragraph paragraph3 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties3 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties3 = new A.DefaultRunProperties() { FontSize = 1000, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill7 = new A.SolidFill();

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset5 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor7.Append(luminanceModulation5);
            schemeColor7.Append(luminanceOffset5);

            solidFill7.Append(schemeColor7);
            A.LatinFont latinFont3 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont3 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont3 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties3.Append(solidFill7);
            defaultRunProperties3.Append(latinFont3);
            defaultRunProperties3.Append(eastAsianFont3);
            defaultRunProperties3.Append(complexScriptFont3);

            paragraphProperties3.Append(defaultRunProperties3);

            A.Run run3 = new A.Run();
            A.RunProperties runProperties3 = new A.RunProperties() { Language = "en-GB" };
            A.Text text3 = new A.Text();
            text3.Text = "Number";

            run3.Append(runProperties3);
            run3.Append(text3);

            A.Run run4 = new A.Run();
            A.RunProperties runProperties4 = new A.RunProperties() { Language = "en-GB", Baseline = 0 };
            A.Text text4 = new A.Text();
            text4.Text = " of tasks";

            run4.Append(runProperties4);
            run4.Append(text4);
            A.EndParagraphRunProperties endParagraphRunProperties3 = new A.EndParagraphRunProperties() { Language = "en-GB" };

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run3);
            paragraph3.Append(run4);
            paragraph3.Append(endParagraphRunProperties3);

            richText2.Append(bodyProperties3);
            richText2.Append(listStyle3);
            richText2.Append(paragraph3);

            chartText2.Append(richText2);
            C.Overlay overlay2 = new C.Overlay() { Val = false };

            C.ChartShapeProperties chartShapeProperties6 = new C.ChartShapeProperties();
            A.NoFill noFill7 = new A.NoFill();

            A.Outline outline6 = new A.Outline();
            A.NoFill noFill8 = new A.NoFill();

            outline6.Append(noFill8);
            A.EffectList effectList6 = new A.EffectList();

            chartShapeProperties6.Append(noFill7);
            chartShapeProperties6.Append(outline6);
            chartShapeProperties6.Append(effectList6);

            C.TextProperties textProperties2 = new C.TextProperties();
            A.BodyProperties bodyProperties4 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle4 = new A.ListStyle();

            A.Paragraph paragraph4 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties4 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties4 = new A.DefaultRunProperties() { FontSize = 1000, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill8 = new A.SolidFill();

            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset6 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor8.Append(luminanceModulation6);
            schemeColor8.Append(luminanceOffset6);

            solidFill8.Append(schemeColor8);
            A.LatinFont latinFont4 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont4 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont4 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties4.Append(solidFill8);
            defaultRunProperties4.Append(latinFont4);
            defaultRunProperties4.Append(eastAsianFont4);
            defaultRunProperties4.Append(complexScriptFont4);

            paragraphProperties4.Append(defaultRunProperties4);
            A.EndParagraphRunProperties endParagraphRunProperties4 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(endParagraphRunProperties4);

            textProperties2.Append(bodyProperties4);
            textProperties2.Append(listStyle4);
            textProperties2.Append(paragraph4);

            title2.Append(chartText2);
            title2.Append(overlay2);
            title2.Append(chartShapeProperties6);
            title2.Append(textProperties2);
            C.NumberingFormat numberingFormat1 = new C.NumberingFormat() { FormatCode = "General", SourceLinked = true };
            C.MajorTickMark majorTickMark1 = new C.MajorTickMark() { Val = C.TickMarkValues.None };
            C.MinorTickMark minorTickMark1 = new C.MinorTickMark() { Val = C.TickMarkValues.None };
            C.TickLabelPosition tickLabelPosition1 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };

            C.ChartShapeProperties chartShapeProperties7 = new C.ChartShapeProperties();
            A.NoFill noFill9 = new A.NoFill();

            A.Outline outline7 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill9 = new A.SolidFill();

            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset7 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor9.Append(luminanceModulation7);
            schemeColor9.Append(luminanceOffset7);

            solidFill9.Append(schemeColor9);
            A.Round round3 = new A.Round();

            outline7.Append(solidFill9);
            outline7.Append(round3);
            A.EffectList effectList7 = new A.EffectList();

            chartShapeProperties7.Append(noFill9);
            chartShapeProperties7.Append(outline7);
            chartShapeProperties7.Append(effectList7);

            C.TextProperties textProperties3 = new C.TextProperties();
            A.BodyProperties bodyProperties5 = new A.BodyProperties() { Rotation = -60000000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle5 = new A.ListStyle();

            A.Paragraph paragraph5 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties5 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties5 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill10 = new A.SolidFill();

            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset8 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor10.Append(luminanceModulation8);
            schemeColor10.Append(luminanceOffset8);

            solidFill10.Append(schemeColor10);
            A.LatinFont latinFont5 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont5 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont5 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties5.Append(solidFill10);
            defaultRunProperties5.Append(latinFont5);
            defaultRunProperties5.Append(eastAsianFont5);
            defaultRunProperties5.Append(complexScriptFont5);

            paragraphProperties5.Append(defaultRunProperties5);
            A.EndParagraphRunProperties endParagraphRunProperties5 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(endParagraphRunProperties5);

            textProperties3.Append(bodyProperties5);
            textProperties3.Append(listStyle5);
            textProperties3.Append(paragraph5);
            C.CrossingAxis crossingAxis1 = new C.CrossingAxis() { Val = (UInt32Value)1189680815U };
            C.Crosses crosses1 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
            C.AutoLabeled autoLabeled1 = new C.AutoLabeled() { Val = true };
            C.LabelAlignment labelAlignment1 = new C.LabelAlignment() { Val = C.LabelAlignmentValues.Center };
            C.LabelOffset labelOffset1 = new C.LabelOffset() { Val = (UInt16Value)100U };
            C.NoMultiLevelLabels noMultiLevelLabels1 = new C.NoMultiLevelLabels() { Val = false };

            categoryAxis1.Append(axisId3);
            categoryAxis1.Append(scaling1);
            categoryAxis1.Append(delete1);
            categoryAxis1.Append(axisPosition1);
            categoryAxis1.Append(title2);
            categoryAxis1.Append(numberingFormat1);
            categoryAxis1.Append(majorTickMark1);
            categoryAxis1.Append(minorTickMark1);
            categoryAxis1.Append(tickLabelPosition1);
            categoryAxis1.Append(chartShapeProperties7);
            categoryAxis1.Append(textProperties3);
            categoryAxis1.Append(crossingAxis1);
            categoryAxis1.Append(crosses1);
            categoryAxis1.Append(autoLabeled1);
            categoryAxis1.Append(labelAlignment1);
            categoryAxis1.Append(labelOffset1);
            categoryAxis1.Append(noMultiLevelLabels1);

            C.ValueAxis valueAxis1 = new C.ValueAxis();
            C.AxisId axisId4 = new C.AxisId() { Val = (UInt32Value)1189680815U };

            C.Scaling scaling2 = new C.Scaling();
            C.Orientation orientation2 = new C.Orientation() { Val = C.OrientationValues.MinMax };

            scaling2.Append(orientation2);
            C.Delete delete2 = new C.Delete() { Val = false };
            C.AxisPosition axisPosition2 = new C.AxisPosition() { Val = C.AxisPositionValues.Left };

            C.MajorGridlines majorGridlines1 = new C.MajorGridlines();

            C.ChartShapeProperties chartShapeProperties8 = new C.ChartShapeProperties();

            A.Outline outline8 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill11 = new A.SolidFill();

            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation9 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset9 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor11.Append(luminanceModulation9);
            schemeColor11.Append(luminanceOffset9);

            solidFill11.Append(schemeColor11);
            A.Round round4 = new A.Round();

            outline8.Append(solidFill11);
            outline8.Append(round4);
            A.EffectList effectList8 = new A.EffectList();

            chartShapeProperties8.Append(outline8);
            chartShapeProperties8.Append(effectList8);

            majorGridlines1.Append(chartShapeProperties8);

            C.Title title3 = new C.Title();

            C.ChartText chartText3 = new C.ChartText();

            C.RichText richText3 = new C.RichText();
            A.BodyProperties bodyProperties6 = new A.BodyProperties() { Rotation = -5400000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle6 = new A.ListStyle();

            A.Paragraph paragraph6 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties6 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties6 = new A.DefaultRunProperties() { FontSize = 1000, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill12 = new A.SolidFill();

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation10 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset10 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor12.Append(luminanceModulation10);
            schemeColor12.Append(luminanceOffset10);

            solidFill12.Append(schemeColor12);
            A.LatinFont latinFont6 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont6 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont6 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties6.Append(solidFill12);
            defaultRunProperties6.Append(latinFont6);
            defaultRunProperties6.Append(eastAsianFont6);
            defaultRunProperties6.Append(complexScriptFont6);

            paragraphProperties6.Append(defaultRunProperties6);

            A.Run run5 = new A.Run();
            A.RunProperties runProperties5 = new A.RunProperties() { Language = "en-GB" };
            A.Text text5 = new A.Text();
            text5.Text = "Bandwidth";

            run5.Append(runProperties5);
            run5.Append(text5);

            A.Run run6 = new A.Run();
            A.RunProperties runProperties6 = new A.RunProperties() { Language = "en-GB", Baseline = 0 };
            A.Text text6 = new A.Text();
            text6.Text = " MiB/s";

            run6.Append(runProperties6);
            run6.Append(text6);
            A.EndParagraphRunProperties endParagraphRunProperties6 = new A.EndParagraphRunProperties() { Language = "en-GB" };

            paragraph6.Append(paragraphProperties6);
            paragraph6.Append(run5);
            paragraph6.Append(run6);
            paragraph6.Append(endParagraphRunProperties6);

            richText3.Append(bodyProperties6);
            richText3.Append(listStyle6);
            richText3.Append(paragraph6);

            chartText3.Append(richText3);
            C.Overlay overlay3 = new C.Overlay() { Val = false };

            C.ChartShapeProperties chartShapeProperties9 = new C.ChartShapeProperties();
            A.NoFill noFill10 = new A.NoFill();

            A.Outline outline9 = new A.Outline();
            A.NoFill noFill11 = new A.NoFill();

            outline9.Append(noFill11);
            A.EffectList effectList9 = new A.EffectList();

            chartShapeProperties9.Append(noFill10);
            chartShapeProperties9.Append(outline9);
            chartShapeProperties9.Append(effectList9);

            C.TextProperties textProperties4 = new C.TextProperties();
            A.BodyProperties bodyProperties7 = new A.BodyProperties() { Rotation = -5400000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle7 = new A.ListStyle();

            A.Paragraph paragraph7 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties7 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties7 = new A.DefaultRunProperties() { FontSize = 1000, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill13 = new A.SolidFill();

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation11 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset11 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor13.Append(luminanceModulation11);
            schemeColor13.Append(luminanceOffset11);

            solidFill13.Append(schemeColor13);
            A.LatinFont latinFont7 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont7 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont7 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties7.Append(solidFill13);
            defaultRunProperties7.Append(latinFont7);
            defaultRunProperties7.Append(eastAsianFont7);
            defaultRunProperties7.Append(complexScriptFont7);

            paragraphProperties7.Append(defaultRunProperties7);
            A.EndParagraphRunProperties endParagraphRunProperties7 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph7.Append(paragraphProperties7);
            paragraph7.Append(endParagraphRunProperties7);

            textProperties4.Append(bodyProperties7);
            textProperties4.Append(listStyle7);
            textProperties4.Append(paragraph7);

            title3.Append(chartText3);
            title3.Append(overlay3);
            title3.Append(chartShapeProperties9);
            title3.Append(textProperties4);
            C.NumberingFormat numberingFormat2 = new C.NumberingFormat() { FormatCode = "General", SourceLinked = true };
            C.MajorTickMark majorTickMark2 = new C.MajorTickMark() { Val = C.TickMarkValues.None };
            C.MinorTickMark minorTickMark2 = new C.MinorTickMark() { Val = C.TickMarkValues.None };
            C.TickLabelPosition tickLabelPosition2 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };

            C.ChartShapeProperties chartShapeProperties10 = new C.ChartShapeProperties();
            A.NoFill noFill12 = new A.NoFill();

            A.Outline outline10 = new A.Outline();
            A.NoFill noFill13 = new A.NoFill();

            outline10.Append(noFill13);
            A.EffectList effectList10 = new A.EffectList();

            chartShapeProperties10.Append(noFill12);
            chartShapeProperties10.Append(outline10);
            chartShapeProperties10.Append(effectList10);

            C.TextProperties textProperties5 = new C.TextProperties();
            A.BodyProperties bodyProperties8 = new A.BodyProperties() { Rotation = -60000000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle8 = new A.ListStyle();

            A.Paragraph paragraph8 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties8 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties8 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill14 = new A.SolidFill();

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation12 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset12 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor14.Append(luminanceModulation12);
            schemeColor14.Append(luminanceOffset12);

            solidFill14.Append(schemeColor14);
            A.LatinFont latinFont8 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont8 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont8 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties8.Append(solidFill14);
            defaultRunProperties8.Append(latinFont8);
            defaultRunProperties8.Append(eastAsianFont8);
            defaultRunProperties8.Append(complexScriptFont8);

            paragraphProperties8.Append(defaultRunProperties8);
            A.EndParagraphRunProperties endParagraphRunProperties8 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph8.Append(paragraphProperties8);
            paragraph8.Append(endParagraphRunProperties8);

            textProperties5.Append(bodyProperties8);
            textProperties5.Append(listStyle8);
            textProperties5.Append(paragraph8);
            C.CrossingAxis crossingAxis2 = new C.CrossingAxis() { Val = (UInt32Value)1191528143U };
            C.Crosses crosses2 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
            C.CrossBetween crossBetween1 = new C.CrossBetween() { Val = C.CrossBetweenValues.Between };

            valueAxis1.Append(axisId4);
            valueAxis1.Append(scaling2);
            valueAxis1.Append(delete2);
            valueAxis1.Append(axisPosition2);
            valueAxis1.Append(majorGridlines1);
            valueAxis1.Append(title3);
            valueAxis1.Append(numberingFormat2);
            valueAxis1.Append(majorTickMark2);
            valueAxis1.Append(minorTickMark2);
            valueAxis1.Append(tickLabelPosition2);
            valueAxis1.Append(chartShapeProperties10);
            valueAxis1.Append(textProperties5);
            valueAxis1.Append(crossingAxis2);
            valueAxis1.Append(crosses2);
            valueAxis1.Append(crossBetween1);

            C.ShapeProperties shapeProperties1 = new C.ShapeProperties();
            A.NoFill noFill14 = new A.NoFill();

            A.Outline outline11 = new A.Outline();
            A.NoFill noFill15 = new A.NoFill();

            outline11.Append(noFill15);
            A.EffectList effectList11 = new A.EffectList();

            shapeProperties1.Append(noFill14);
            shapeProperties1.Append(outline11);
            shapeProperties1.Append(effectList11);

            plotArea1.Append(layout1);
            plotArea1.Append(barChart1);
            plotArea1.Append(categoryAxis1);
            plotArea1.Append(valueAxis1);
            plotArea1.Append(shapeProperties1);

            C.Legend legend1 = new C.Legend();
            C.LegendPosition legendPosition1 = new C.LegendPosition() { Val = C.LegendPositionValues.Bottom };
            C.Overlay overlay4 = new C.Overlay() { Val = false };

            C.ChartShapeProperties chartShapeProperties11 = new C.ChartShapeProperties();
            A.NoFill noFill16 = new A.NoFill();

            A.Outline outline12 = new A.Outline();
            A.NoFill noFill17 = new A.NoFill();

            outline12.Append(noFill17);
            A.EffectList effectList12 = new A.EffectList();

            chartShapeProperties11.Append(noFill16);
            chartShapeProperties11.Append(outline12);
            chartShapeProperties11.Append(effectList12);

            C.TextProperties textProperties6 = new C.TextProperties();
            A.BodyProperties bodyProperties9 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle9 = new A.ListStyle();

            A.Paragraph paragraph9 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties9 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties9 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill15 = new A.SolidFill();

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation13 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset13 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor15.Append(luminanceModulation13);
            schemeColor15.Append(luminanceOffset13);

            solidFill15.Append(schemeColor15);
            A.LatinFont latinFont9 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont9 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont9 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties9.Append(solidFill15);
            defaultRunProperties9.Append(latinFont9);
            defaultRunProperties9.Append(eastAsianFont9);
            defaultRunProperties9.Append(complexScriptFont9);

            paragraphProperties9.Append(defaultRunProperties9);
            A.EndParagraphRunProperties endParagraphRunProperties9 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph9.Append(paragraphProperties9);
            paragraph9.Append(endParagraphRunProperties9);

            textProperties6.Append(bodyProperties9);
            textProperties6.Append(listStyle9);
            textProperties6.Append(paragraph9);

            legend1.Append(legendPosition1);
            legend1.Append(overlay4);
            legend1.Append(chartShapeProperties11);
            legend1.Append(textProperties6);
            C.PlotVisibleOnly plotVisibleOnly1 = new C.PlotVisibleOnly() { Val = true };
            C.DisplayBlanksAs displayBlanksAs1 = new C.DisplayBlanksAs() { Val = C.DisplayBlanksAsValues.Gap };
            C.ShowDataLabelsOverMaximum showDataLabelsOverMaximum1 = new C.ShowDataLabelsOverMaximum() { Val = false };

            chart1.Append(title1);
            chart1.Append(autoTitleDeleted1);
            chart1.Append(plotArea1);
            chart1.Append(legend1);
            chart1.Append(plotVisibleOnly1);
            chart1.Append(displayBlanksAs1);
            chart1.Append(showDataLabelsOverMaximum1);

            C.ShapeProperties shapeProperties2 = new C.ShapeProperties();

            A.SolidFill solidFill16 = new A.SolidFill();
            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill16.Append(schemeColor16);

            A.Outline outline13 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill17 = new A.SolidFill();

            A.SchemeColor schemeColor17 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation14 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset14 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor17.Append(luminanceModulation14);
            schemeColor17.Append(luminanceOffset14);

            solidFill17.Append(schemeColor17);
            A.Round round5 = new A.Round();

            outline13.Append(solidFill17);
            outline13.Append(round5);
            A.EffectList effectList13 = new A.EffectList();

            shapeProperties2.Append(solidFill16);
            shapeProperties2.Append(outline13);
            shapeProperties2.Append(effectList13);

            C.TextProperties textProperties7 = new C.TextProperties();
            A.BodyProperties bodyProperties10 = new A.BodyProperties();
            A.ListStyle listStyle10 = new A.ListStyle();

            A.Paragraph paragraph10 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties10 = new A.ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties10 = new A.DefaultRunProperties();

            paragraphProperties10.Append(defaultRunProperties10);
            A.EndParagraphRunProperties endParagraphRunProperties10 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph10.Append(paragraphProperties10);
            paragraph10.Append(endParagraphRunProperties10);

            textProperties7.Append(bodyProperties10);
            textProperties7.Append(listStyle10);
            textProperties7.Append(paragraph10);

            C.PrintSettings printSettings1 = new C.PrintSettings();
            C.HeaderFooter headerFooter1 = new C.HeaderFooter();
            C.PageMargins pageMargins1 = new C.PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            C.PageSetup pageSetup1 = new C.PageSetup();

            printSettings1.Append(headerFooter1);
            printSettings1.Append(pageMargins1);
            printSettings1.Append(pageSetup1);

            chartSpace1.Append(date19041);
            chartSpace1.Append(editingLanguage1);
            chartSpace1.Append(roundedCorners1);
            chartSpace1.Append(alternateContent1);
            chartSpace1.Append(chart1);
            chartSpace1.Append(shapeProperties2);
            chartSpace1.Append(textProperties7);
            chartSpace1.Append(printSettings1);

            chartPart1.ChartSpace = chartSpace1;
        }

        // Generates content of chartColorStylePart1.
        private void GenerateChartColorStylePart1Content(ChartColorStylePart chartColorStylePart1)
        {
            Cs.ColorStyle colorStyle1 = new Cs.ColorStyle() { Method = "cycle", Id = (UInt32Value)10U };
            colorStyle1.AddNamespaceDeclaration("cs", "http://schemas.microsoft.com/office/drawing/2012/chartStyle");
            colorStyle1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            A.SchemeColor schemeColor18 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.SchemeColor schemeColor19 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent2 };
            A.SchemeColor schemeColor20 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent3 };
            A.SchemeColor schemeColor21 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent4 };
            A.SchemeColor schemeColor22 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent5 };
            A.SchemeColor schemeColor23 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent6 };
            Cs.ColorStyleVariation colorStyleVariation1 = new Cs.ColorStyleVariation();

            Cs.ColorStyleVariation colorStyleVariation2 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation15 = new A.LuminanceModulation() { Val = 60000 };

            colorStyleVariation2.Append(luminanceModulation15);

            Cs.ColorStyleVariation colorStyleVariation3 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation16 = new A.LuminanceModulation() { Val = 80000 };
            A.LuminanceOffset luminanceOffset15 = new A.LuminanceOffset() { Val = 20000 };

            colorStyleVariation3.Append(luminanceModulation16);
            colorStyleVariation3.Append(luminanceOffset15);

            Cs.ColorStyleVariation colorStyleVariation4 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation17 = new A.LuminanceModulation() { Val = 80000 };

            colorStyleVariation4.Append(luminanceModulation17);

            Cs.ColorStyleVariation colorStyleVariation5 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation18 = new A.LuminanceModulation() { Val = 60000 };
            A.LuminanceOffset luminanceOffset16 = new A.LuminanceOffset() { Val = 40000 };

            colorStyleVariation5.Append(luminanceModulation18);
            colorStyleVariation5.Append(luminanceOffset16);

            Cs.ColorStyleVariation colorStyleVariation6 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation19 = new A.LuminanceModulation() { Val = 50000 };

            colorStyleVariation6.Append(luminanceModulation19);

            Cs.ColorStyleVariation colorStyleVariation7 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation20 = new A.LuminanceModulation() { Val = 70000 };
            A.LuminanceOffset luminanceOffset17 = new A.LuminanceOffset() { Val = 30000 };

            colorStyleVariation7.Append(luminanceModulation20);
            colorStyleVariation7.Append(luminanceOffset17);

            Cs.ColorStyleVariation colorStyleVariation8 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation21 = new A.LuminanceModulation() { Val = 70000 };

            colorStyleVariation8.Append(luminanceModulation21);

            Cs.ColorStyleVariation colorStyleVariation9 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation22 = new A.LuminanceModulation() { Val = 50000 };
            A.LuminanceOffset luminanceOffset18 = new A.LuminanceOffset() { Val = 50000 };

            colorStyleVariation9.Append(luminanceModulation22);
            colorStyleVariation9.Append(luminanceOffset18);

            colorStyle1.Append(schemeColor18);
            colorStyle1.Append(schemeColor19);
            colorStyle1.Append(schemeColor20);
            colorStyle1.Append(schemeColor21);
            colorStyle1.Append(schemeColor22);
            colorStyle1.Append(schemeColor23);
            colorStyle1.Append(colorStyleVariation1);
            colorStyle1.Append(colorStyleVariation2);
            colorStyle1.Append(colorStyleVariation3);
            colorStyle1.Append(colorStyleVariation4);
            colorStyle1.Append(colorStyleVariation5);
            colorStyle1.Append(colorStyleVariation6);
            colorStyle1.Append(colorStyleVariation7);
            colorStyle1.Append(colorStyleVariation8);
            colorStyle1.Append(colorStyleVariation9);

            chartColorStylePart1.ColorStyle = colorStyle1;
        }

        // Generates content of chartStylePart1.
        private void GenerateChartStylePart1Content(ChartStylePart chartStylePart1)
        {
            Cs.ChartStyle chartStyle1 = new Cs.ChartStyle() { Id = (UInt32Value)201U };
            chartStyle1.AddNamespaceDeclaration("cs", "http://schemas.microsoft.com/office/drawing/2012/chartStyle");
            chartStyle1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            Cs.AxisTitle axisTitle1 = new Cs.AxisTitle();
            Cs.LineReference lineReference1 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference1 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference1 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference1 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor24 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation23 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset19 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor24.Append(luminanceModulation23);
            schemeColor24.Append(luminanceOffset19);

            fontReference1.Append(schemeColor24);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType1 = new Cs.TextCharacterPropertiesType() { FontSize = 1000, Kerning = 1200 };

            axisTitle1.Append(lineReference1);
            axisTitle1.Append(fillReference1);
            axisTitle1.Append(effectReference1);
            axisTitle1.Append(fontReference1);
            axisTitle1.Append(textCharacterPropertiesType1);

            Cs.CategoryAxis categoryAxis2 = new Cs.CategoryAxis();
            Cs.LineReference lineReference2 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference2 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference2 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference2 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor25 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation24 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset20 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor25.Append(luminanceModulation24);
            schemeColor25.Append(luminanceOffset20);

            fontReference2.Append(schemeColor25);

            Cs.ShapeProperties shapeProperties3 = new Cs.ShapeProperties();

            A.Outline outline14 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill18 = new A.SolidFill();

            A.SchemeColor schemeColor26 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation25 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset21 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor26.Append(luminanceModulation25);
            schemeColor26.Append(luminanceOffset21);

            solidFill18.Append(schemeColor26);
            A.Round round6 = new A.Round();

            outline14.Append(solidFill18);
            outline14.Append(round6);

            shapeProperties3.Append(outline14);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType2 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            categoryAxis2.Append(lineReference2);
            categoryAxis2.Append(fillReference2);
            categoryAxis2.Append(effectReference2);
            categoryAxis2.Append(fontReference2);
            categoryAxis2.Append(shapeProperties3);
            categoryAxis2.Append(textCharacterPropertiesType2);

            Cs.ChartArea chartArea1 = new Cs.ChartArea() { Modifiers = new ListValue<StringValue>() { InnerText = "allowNoFillOverride allowNoLineOverride" } };
            Cs.LineReference lineReference3 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference3 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference3 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference3 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor27 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference3.Append(schemeColor27);

            Cs.ShapeProperties shapeProperties4 = new Cs.ShapeProperties();

            A.SolidFill solidFill19 = new A.SolidFill();
            A.SchemeColor schemeColor28 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill19.Append(schemeColor28);

            A.Outline outline15 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill20 = new A.SolidFill();

            A.SchemeColor schemeColor29 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation26 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset22 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor29.Append(luminanceModulation26);
            schemeColor29.Append(luminanceOffset22);

            solidFill20.Append(schemeColor29);
            A.Round round7 = new A.Round();

            outline15.Append(solidFill20);
            outline15.Append(round7);

            shapeProperties4.Append(solidFill19);
            shapeProperties4.Append(outline15);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType3 = new Cs.TextCharacterPropertiesType() { FontSize = 1000, Kerning = 1200 };

            chartArea1.Append(lineReference3);
            chartArea1.Append(fillReference3);
            chartArea1.Append(effectReference3);
            chartArea1.Append(fontReference3);
            chartArea1.Append(shapeProperties4);
            chartArea1.Append(textCharacterPropertiesType3);

            Cs.DataLabel dataLabel1 = new Cs.DataLabel();
            Cs.LineReference lineReference4 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference4 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference4 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference4 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor30 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation27 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset23 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor30.Append(luminanceModulation27);
            schemeColor30.Append(luminanceOffset23);

            fontReference4.Append(schemeColor30);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType4 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            dataLabel1.Append(lineReference4);
            dataLabel1.Append(fillReference4);
            dataLabel1.Append(effectReference4);
            dataLabel1.Append(fontReference4);
            dataLabel1.Append(textCharacterPropertiesType4);

            Cs.DataLabelCallout dataLabelCallout1 = new Cs.DataLabelCallout();
            Cs.LineReference lineReference5 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference5 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference5 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference5 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor31 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation28 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset24 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor31.Append(luminanceModulation28);
            schemeColor31.Append(luminanceOffset24);

            fontReference5.Append(schemeColor31);

            Cs.ShapeProperties shapeProperties5 = new Cs.ShapeProperties();

            A.SolidFill solidFill21 = new A.SolidFill();
            A.SchemeColor schemeColor32 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            solidFill21.Append(schemeColor32);

            A.Outline outline16 = new A.Outline();

            A.SolidFill solidFill22 = new A.SolidFill();

            A.SchemeColor schemeColor33 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation29 = new A.LuminanceModulation() { Val = 25000 };
            A.LuminanceOffset luminanceOffset25 = new A.LuminanceOffset() { Val = 75000 };

            schemeColor33.Append(luminanceModulation29);
            schemeColor33.Append(luminanceOffset25);

            solidFill22.Append(schemeColor33);

            outline16.Append(solidFill22);

            shapeProperties5.Append(solidFill21);
            shapeProperties5.Append(outline16);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType5 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            Cs.TextBodyProperties textBodyProperties1 = new Cs.TextBodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Clip, HorizontalOverflow = A.TextHorizontalOverflowValues.Clip, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 36576, TopInset = 18288, RightInset = 36576, BottomInset = 18288, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ShapeAutoFit shapeAutoFit1 = new A.ShapeAutoFit();

            textBodyProperties1.Append(shapeAutoFit1);

            dataLabelCallout1.Append(lineReference5);
            dataLabelCallout1.Append(fillReference5);
            dataLabelCallout1.Append(effectReference5);
            dataLabelCallout1.Append(fontReference5);
            dataLabelCallout1.Append(shapeProperties5);
            dataLabelCallout1.Append(textCharacterPropertiesType5);
            dataLabelCallout1.Append(textBodyProperties1);

            Cs.DataPoint dataPoint1 = new Cs.DataPoint();
            Cs.LineReference lineReference6 = new Cs.LineReference() { Index = (UInt32Value)0U };

            Cs.FillReference fillReference6 = new Cs.FillReference() { Index = (UInt32Value)1U };
            Cs.StyleColor styleColor1 = new Cs.StyleColor() { Val = "auto" };

            fillReference6.Append(styleColor1);
            Cs.EffectReference effectReference6 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference6 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor34 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference6.Append(schemeColor34);

            dataPoint1.Append(lineReference6);
            dataPoint1.Append(fillReference6);
            dataPoint1.Append(effectReference6);
            dataPoint1.Append(fontReference6);

            Cs.DataPoint3D dataPoint3D1 = new Cs.DataPoint3D();
            Cs.LineReference lineReference7 = new Cs.LineReference() { Index = (UInt32Value)0U };

            Cs.FillReference fillReference7 = new Cs.FillReference() { Index = (UInt32Value)1U };
            Cs.StyleColor styleColor2 = new Cs.StyleColor() { Val = "auto" };

            fillReference7.Append(styleColor2);
            Cs.EffectReference effectReference7 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference7 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor35 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference7.Append(schemeColor35);

            dataPoint3D1.Append(lineReference7);
            dataPoint3D1.Append(fillReference7);
            dataPoint3D1.Append(effectReference7);
            dataPoint3D1.Append(fontReference7);

            Cs.DataPointLine dataPointLine1 = new Cs.DataPointLine();

            Cs.LineReference lineReference8 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.StyleColor styleColor3 = new Cs.StyleColor() { Val = "auto" };

            lineReference8.Append(styleColor3);
            Cs.FillReference fillReference8 = new Cs.FillReference() { Index = (UInt32Value)1U };
            Cs.EffectReference effectReference8 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference8 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor36 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference8.Append(schemeColor36);

            Cs.ShapeProperties shapeProperties6 = new Cs.ShapeProperties();

            A.Outline outline17 = new A.Outline() { Width = 28575, CapType = A.LineCapValues.Round };

            A.SolidFill solidFill23 = new A.SolidFill();
            A.SchemeColor schemeColor37 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill23.Append(schemeColor37);
            A.Round round8 = new A.Round();

            outline17.Append(solidFill23);
            outline17.Append(round8);

            shapeProperties6.Append(outline17);

            dataPointLine1.Append(lineReference8);
            dataPointLine1.Append(fillReference8);
            dataPointLine1.Append(effectReference8);
            dataPointLine1.Append(fontReference8);
            dataPointLine1.Append(shapeProperties6);

            Cs.DataPointMarker dataPointMarker1 = new Cs.DataPointMarker();

            Cs.LineReference lineReference9 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.StyleColor styleColor4 = new Cs.StyleColor() { Val = "auto" };

            lineReference9.Append(styleColor4);

            Cs.FillReference fillReference9 = new Cs.FillReference() { Index = (UInt32Value)1U };
            Cs.StyleColor styleColor5 = new Cs.StyleColor() { Val = "auto" };

            fillReference9.Append(styleColor5);
            Cs.EffectReference effectReference9 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference9 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor38 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference9.Append(schemeColor38);

            Cs.ShapeProperties shapeProperties7 = new Cs.ShapeProperties();

            A.Outline outline18 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill24 = new A.SolidFill();
            A.SchemeColor schemeColor39 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill24.Append(schemeColor39);

            outline18.Append(solidFill24);

            shapeProperties7.Append(outline18);

            dataPointMarker1.Append(lineReference9);
            dataPointMarker1.Append(fillReference9);
            dataPointMarker1.Append(effectReference9);
            dataPointMarker1.Append(fontReference9);
            dataPointMarker1.Append(shapeProperties7);
            Cs.MarkerLayoutProperties markerLayoutProperties1 = new Cs.MarkerLayoutProperties() { Symbol = Cs.MarkerStyle.Circle, Size = 5 };

            Cs.DataPointWireframe dataPointWireframe1 = new Cs.DataPointWireframe();

            Cs.LineReference lineReference10 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.StyleColor styleColor6 = new Cs.StyleColor() { Val = "auto" };

            lineReference10.Append(styleColor6);
            Cs.FillReference fillReference10 = new Cs.FillReference() { Index = (UInt32Value)1U };
            Cs.EffectReference effectReference10 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference10 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor40 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference10.Append(schemeColor40);

            Cs.ShapeProperties shapeProperties8 = new Cs.ShapeProperties();

            A.Outline outline19 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Round };

            A.SolidFill solidFill25 = new A.SolidFill();
            A.SchemeColor schemeColor41 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill25.Append(schemeColor41);
            A.Round round9 = new A.Round();

            outline19.Append(solidFill25);
            outline19.Append(round9);

            shapeProperties8.Append(outline19);

            dataPointWireframe1.Append(lineReference10);
            dataPointWireframe1.Append(fillReference10);
            dataPointWireframe1.Append(effectReference10);
            dataPointWireframe1.Append(fontReference10);
            dataPointWireframe1.Append(shapeProperties8);

            Cs.DataTableStyle dataTableStyle1 = new Cs.DataTableStyle();
            Cs.LineReference lineReference11 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference11 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference11 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference11 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor42 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation30 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset26 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor42.Append(luminanceModulation30);
            schemeColor42.Append(luminanceOffset26);

            fontReference11.Append(schemeColor42);

            Cs.ShapeProperties shapeProperties9 = new Cs.ShapeProperties();
            A.NoFill noFill18 = new A.NoFill();

            A.Outline outline20 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill26 = new A.SolidFill();

            A.SchemeColor schemeColor43 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation31 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset27 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor43.Append(luminanceModulation31);
            schemeColor43.Append(luminanceOffset27);

            solidFill26.Append(schemeColor43);
            A.Round round10 = new A.Round();

            outline20.Append(solidFill26);
            outline20.Append(round10);

            shapeProperties9.Append(noFill18);
            shapeProperties9.Append(outline20);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType6 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            dataTableStyle1.Append(lineReference11);
            dataTableStyle1.Append(fillReference11);
            dataTableStyle1.Append(effectReference11);
            dataTableStyle1.Append(fontReference11);
            dataTableStyle1.Append(shapeProperties9);
            dataTableStyle1.Append(textCharacterPropertiesType6);

            Cs.DownBar downBar1 = new Cs.DownBar();
            Cs.LineReference lineReference12 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference12 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference12 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference12 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor44 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference12.Append(schemeColor44);

            Cs.ShapeProperties shapeProperties10 = new Cs.ShapeProperties();

            A.SolidFill solidFill27 = new A.SolidFill();

            A.SchemeColor schemeColor45 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation32 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset28 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor45.Append(luminanceModulation32);
            schemeColor45.Append(luminanceOffset28);

            solidFill27.Append(schemeColor45);

            A.Outline outline21 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill28 = new A.SolidFill();

            A.SchemeColor schemeColor46 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation33 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset29 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor46.Append(luminanceModulation33);
            schemeColor46.Append(luminanceOffset29);

            solidFill28.Append(schemeColor46);

            outline21.Append(solidFill28);

            shapeProperties10.Append(solidFill27);
            shapeProperties10.Append(outline21);

            downBar1.Append(lineReference12);
            downBar1.Append(fillReference12);
            downBar1.Append(effectReference12);
            downBar1.Append(fontReference12);
            downBar1.Append(shapeProperties10);

            Cs.DropLine dropLine1 = new Cs.DropLine();
            Cs.LineReference lineReference13 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference13 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference13 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference13 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor47 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference13.Append(schemeColor47);

            Cs.ShapeProperties shapeProperties11 = new Cs.ShapeProperties();

            A.Outline outline22 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill29 = new A.SolidFill();

            A.SchemeColor schemeColor48 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation34 = new A.LuminanceModulation() { Val = 35000 };
            A.LuminanceOffset luminanceOffset30 = new A.LuminanceOffset() { Val = 65000 };

            schemeColor48.Append(luminanceModulation34);
            schemeColor48.Append(luminanceOffset30);

            solidFill29.Append(schemeColor48);
            A.Round round11 = new A.Round();

            outline22.Append(solidFill29);
            outline22.Append(round11);

            shapeProperties11.Append(outline22);

            dropLine1.Append(lineReference13);
            dropLine1.Append(fillReference13);
            dropLine1.Append(effectReference13);
            dropLine1.Append(fontReference13);
            dropLine1.Append(shapeProperties11);

            Cs.ErrorBar errorBar1 = new Cs.ErrorBar();
            Cs.LineReference lineReference14 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference14 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference14 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference14 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor49 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference14.Append(schemeColor49);

            Cs.ShapeProperties shapeProperties12 = new Cs.ShapeProperties();

            A.Outline outline23 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill30 = new A.SolidFill();

            A.SchemeColor schemeColor50 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation35 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset31 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor50.Append(luminanceModulation35);
            schemeColor50.Append(luminanceOffset31);

            solidFill30.Append(schemeColor50);
            A.Round round12 = new A.Round();

            outline23.Append(solidFill30);
            outline23.Append(round12);

            shapeProperties12.Append(outline23);

            errorBar1.Append(lineReference14);
            errorBar1.Append(fillReference14);
            errorBar1.Append(effectReference14);
            errorBar1.Append(fontReference14);
            errorBar1.Append(shapeProperties12);

            Cs.Floor floor1 = new Cs.Floor();
            Cs.LineReference lineReference15 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference15 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference15 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference15 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor51 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference15.Append(schemeColor51);

            Cs.ShapeProperties shapeProperties13 = new Cs.ShapeProperties();
            A.NoFill noFill19 = new A.NoFill();

            A.Outline outline24 = new A.Outline();
            A.NoFill noFill20 = new A.NoFill();

            outline24.Append(noFill20);

            shapeProperties13.Append(noFill19);
            shapeProperties13.Append(outline24);

            floor1.Append(lineReference15);
            floor1.Append(fillReference15);
            floor1.Append(effectReference15);
            floor1.Append(fontReference15);
            floor1.Append(shapeProperties13);

            Cs.GridlineMajor gridlineMajor1 = new Cs.GridlineMajor();
            Cs.LineReference lineReference16 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference16 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference16 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference16 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor52 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference16.Append(schemeColor52);

            Cs.ShapeProperties shapeProperties14 = new Cs.ShapeProperties();

            A.Outline outline25 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill31 = new A.SolidFill();

            A.SchemeColor schemeColor53 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation36 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset32 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor53.Append(luminanceModulation36);
            schemeColor53.Append(luminanceOffset32);

            solidFill31.Append(schemeColor53);
            A.Round round13 = new A.Round();

            outline25.Append(solidFill31);
            outline25.Append(round13);

            shapeProperties14.Append(outline25);

            gridlineMajor1.Append(lineReference16);
            gridlineMajor1.Append(fillReference16);
            gridlineMajor1.Append(effectReference16);
            gridlineMajor1.Append(fontReference16);
            gridlineMajor1.Append(shapeProperties14);

            Cs.GridlineMinor gridlineMinor1 = new Cs.GridlineMinor();
            Cs.LineReference lineReference17 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference17 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference17 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference17 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor54 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference17.Append(schemeColor54);

            Cs.ShapeProperties shapeProperties15 = new Cs.ShapeProperties();

            A.Outline outline26 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill32 = new A.SolidFill();

            A.SchemeColor schemeColor55 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation37 = new A.LuminanceModulation() { Val = 5000 };
            A.LuminanceOffset luminanceOffset33 = new A.LuminanceOffset() { Val = 95000 };

            schemeColor55.Append(luminanceModulation37);
            schemeColor55.Append(luminanceOffset33);

            solidFill32.Append(schemeColor55);
            A.Round round14 = new A.Round();

            outline26.Append(solidFill32);
            outline26.Append(round14);

            shapeProperties15.Append(outline26);

            gridlineMinor1.Append(lineReference17);
            gridlineMinor1.Append(fillReference17);
            gridlineMinor1.Append(effectReference17);
            gridlineMinor1.Append(fontReference17);
            gridlineMinor1.Append(shapeProperties15);

            Cs.HiLoLine hiLoLine1 = new Cs.HiLoLine();
            Cs.LineReference lineReference18 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference18 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference18 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference18 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor56 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference18.Append(schemeColor56);

            Cs.ShapeProperties shapeProperties16 = new Cs.ShapeProperties();

            A.Outline outline27 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill33 = new A.SolidFill();

            A.SchemeColor schemeColor57 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation38 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset34 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor57.Append(luminanceModulation38);
            schemeColor57.Append(luminanceOffset34);

            solidFill33.Append(schemeColor57);
            A.Round round15 = new A.Round();

            outline27.Append(solidFill33);
            outline27.Append(round15);

            shapeProperties16.Append(outline27);

            hiLoLine1.Append(lineReference18);
            hiLoLine1.Append(fillReference18);
            hiLoLine1.Append(effectReference18);
            hiLoLine1.Append(fontReference18);
            hiLoLine1.Append(shapeProperties16);

            Cs.LeaderLine leaderLine1 = new Cs.LeaderLine();
            Cs.LineReference lineReference19 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference19 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference19 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference19 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor58 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference19.Append(schemeColor58);

            Cs.ShapeProperties shapeProperties17 = new Cs.ShapeProperties();

            A.Outline outline28 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill34 = new A.SolidFill();

            A.SchemeColor schemeColor59 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation39 = new A.LuminanceModulation() { Val = 35000 };
            A.LuminanceOffset luminanceOffset35 = new A.LuminanceOffset() { Val = 65000 };

            schemeColor59.Append(luminanceModulation39);
            schemeColor59.Append(luminanceOffset35);

            solidFill34.Append(schemeColor59);
            A.Round round16 = new A.Round();

            outline28.Append(solidFill34);
            outline28.Append(round16);

            shapeProperties17.Append(outline28);

            leaderLine1.Append(lineReference19);
            leaderLine1.Append(fillReference19);
            leaderLine1.Append(effectReference19);
            leaderLine1.Append(fontReference19);
            leaderLine1.Append(shapeProperties17);

            Cs.LegendStyle legendStyle1 = new Cs.LegendStyle();
            Cs.LineReference lineReference20 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference20 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference20 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference20 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor60 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation40 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset36 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor60.Append(luminanceModulation40);
            schemeColor60.Append(luminanceOffset36);

            fontReference20.Append(schemeColor60);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType7 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            legendStyle1.Append(lineReference20);
            legendStyle1.Append(fillReference20);
            legendStyle1.Append(effectReference20);
            legendStyle1.Append(fontReference20);
            legendStyle1.Append(textCharacterPropertiesType7);

            Cs.PlotArea plotArea2 = new Cs.PlotArea() { Modifiers = new ListValue<StringValue>() { InnerText = "allowNoFillOverride allowNoLineOverride" } };
            Cs.LineReference lineReference21 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference21 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference21 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference21 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor61 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference21.Append(schemeColor61);

            plotArea2.Append(lineReference21);
            plotArea2.Append(fillReference21);
            plotArea2.Append(effectReference21);
            plotArea2.Append(fontReference21);

            Cs.PlotArea3D plotArea3D1 = new Cs.PlotArea3D() { Modifiers = new ListValue<StringValue>() { InnerText = "allowNoFillOverride allowNoLineOverride" } };
            Cs.LineReference lineReference22 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference22 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference22 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference22 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor62 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference22.Append(schemeColor62);

            plotArea3D1.Append(lineReference22);
            plotArea3D1.Append(fillReference22);
            plotArea3D1.Append(effectReference22);
            plotArea3D1.Append(fontReference22);

            Cs.SeriesAxis seriesAxis1 = new Cs.SeriesAxis();
            Cs.LineReference lineReference23 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference23 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference23 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference23 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor63 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation41 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset37 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor63.Append(luminanceModulation41);
            schemeColor63.Append(luminanceOffset37);

            fontReference23.Append(schemeColor63);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType8 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            seriesAxis1.Append(lineReference23);
            seriesAxis1.Append(fillReference23);
            seriesAxis1.Append(effectReference23);
            seriesAxis1.Append(fontReference23);
            seriesAxis1.Append(textCharacterPropertiesType8);

            Cs.SeriesLine seriesLine1 = new Cs.SeriesLine();
            Cs.LineReference lineReference24 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference24 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference24 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference24 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor64 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference24.Append(schemeColor64);

            Cs.ShapeProperties shapeProperties18 = new Cs.ShapeProperties();

            A.Outline outline29 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill35 = new A.SolidFill();

            A.SchemeColor schemeColor65 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation42 = new A.LuminanceModulation() { Val = 35000 };
            A.LuminanceOffset luminanceOffset38 = new A.LuminanceOffset() { Val = 65000 };

            schemeColor65.Append(luminanceModulation42);
            schemeColor65.Append(luminanceOffset38);

            solidFill35.Append(schemeColor65);
            A.Round round17 = new A.Round();

            outline29.Append(solidFill35);
            outline29.Append(round17);

            shapeProperties18.Append(outline29);

            seriesLine1.Append(lineReference24);
            seriesLine1.Append(fillReference24);
            seriesLine1.Append(effectReference24);
            seriesLine1.Append(fontReference24);
            seriesLine1.Append(shapeProperties18);

            Cs.TitleStyle titleStyle1 = new Cs.TitleStyle();
            Cs.LineReference lineReference25 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference25 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference25 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference25 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor66 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation43 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset39 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor66.Append(luminanceModulation43);
            schemeColor66.Append(luminanceOffset39);

            fontReference25.Append(schemeColor66);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType9 = new Cs.TextCharacterPropertiesType() { FontSize = 1400, Bold = false, Kerning = 1200, Spacing = 0, Baseline = 0 };

            titleStyle1.Append(lineReference25);
            titleStyle1.Append(fillReference25);
            titleStyle1.Append(effectReference25);
            titleStyle1.Append(fontReference25);
            titleStyle1.Append(textCharacterPropertiesType9);

            Cs.TrendlineStyle trendlineStyle1 = new Cs.TrendlineStyle();

            Cs.LineReference lineReference26 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.StyleColor styleColor7 = new Cs.StyleColor() { Val = "auto" };

            lineReference26.Append(styleColor7);
            Cs.FillReference fillReference26 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference26 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference26 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor67 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference26.Append(schemeColor67);

            Cs.ShapeProperties shapeProperties19 = new Cs.ShapeProperties();

            A.Outline outline30 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Round };

            A.SolidFill solidFill36 = new A.SolidFill();
            A.SchemeColor schemeColor68 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill36.Append(schemeColor68);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.SystemDot };

            outline30.Append(solidFill36);
            outline30.Append(presetDash1);

            shapeProperties19.Append(outline30);

            trendlineStyle1.Append(lineReference26);
            trendlineStyle1.Append(fillReference26);
            trendlineStyle1.Append(effectReference26);
            trendlineStyle1.Append(fontReference26);
            trendlineStyle1.Append(shapeProperties19);

            Cs.TrendlineLabel trendlineLabel1 = new Cs.TrendlineLabel();
            Cs.LineReference lineReference27 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference27 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference27 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference27 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor69 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation44 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset40 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor69.Append(luminanceModulation44);
            schemeColor69.Append(luminanceOffset40);

            fontReference27.Append(schemeColor69);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType10 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            trendlineLabel1.Append(lineReference27);
            trendlineLabel1.Append(fillReference27);
            trendlineLabel1.Append(effectReference27);
            trendlineLabel1.Append(fontReference27);
            trendlineLabel1.Append(textCharacterPropertiesType10);

            Cs.UpBar upBar1 = new Cs.UpBar();
            Cs.LineReference lineReference28 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference28 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference28 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference28 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor70 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference28.Append(schemeColor70);

            Cs.ShapeProperties shapeProperties20 = new Cs.ShapeProperties();

            A.SolidFill solidFill37 = new A.SolidFill();
            A.SchemeColor schemeColor71 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            solidFill37.Append(schemeColor71);

            A.Outline outline31 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill38 = new A.SolidFill();

            A.SchemeColor schemeColor72 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation45 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset41 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor72.Append(luminanceModulation45);
            schemeColor72.Append(luminanceOffset41);

            solidFill38.Append(schemeColor72);

            outline31.Append(solidFill38);

            shapeProperties20.Append(solidFill37);
            shapeProperties20.Append(outline31);

            upBar1.Append(lineReference28);
            upBar1.Append(fillReference28);
            upBar1.Append(effectReference28);
            upBar1.Append(fontReference28);
            upBar1.Append(shapeProperties20);

            Cs.ValueAxis valueAxis2 = new Cs.ValueAxis();
            Cs.LineReference lineReference29 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference29 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference29 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference29 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor73 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation46 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset42 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor73.Append(luminanceModulation46);
            schemeColor73.Append(luminanceOffset42);

            fontReference29.Append(schemeColor73);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType11 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            valueAxis2.Append(lineReference29);
            valueAxis2.Append(fillReference29);
            valueAxis2.Append(effectReference29);
            valueAxis2.Append(fontReference29);
            valueAxis2.Append(textCharacterPropertiesType11);

            Cs.Wall wall1 = new Cs.Wall();
            Cs.LineReference lineReference30 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference30 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference30 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference30 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor74 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference30.Append(schemeColor74);

            Cs.ShapeProperties shapeProperties21 = new Cs.ShapeProperties();
            A.NoFill noFill21 = new A.NoFill();

            A.Outline outline32 = new A.Outline();
            A.NoFill noFill22 = new A.NoFill();

            outline32.Append(noFill22);

            shapeProperties21.Append(noFill21);
            shapeProperties21.Append(outline32);

            wall1.Append(lineReference30);
            wall1.Append(fillReference30);
            wall1.Append(effectReference30);
            wall1.Append(fontReference30);
            wall1.Append(shapeProperties21);

            chartStyle1.Append(axisTitle1);
            chartStyle1.Append(categoryAxis2);
            chartStyle1.Append(chartArea1);
            chartStyle1.Append(dataLabel1);
            chartStyle1.Append(dataLabelCallout1);
            chartStyle1.Append(dataPoint1);
            chartStyle1.Append(dataPoint3D1);
            chartStyle1.Append(dataPointLine1);
            chartStyle1.Append(dataPointMarker1);
            chartStyle1.Append(markerLayoutProperties1);
            chartStyle1.Append(dataPointWireframe1);
            chartStyle1.Append(dataTableStyle1);
            chartStyle1.Append(downBar1);
            chartStyle1.Append(dropLine1);
            chartStyle1.Append(errorBar1);
            chartStyle1.Append(floor1);
            chartStyle1.Append(gridlineMajor1);
            chartStyle1.Append(gridlineMinor1);
            chartStyle1.Append(hiLoLine1);
            chartStyle1.Append(leaderLine1);
            chartStyle1.Append(legendStyle1);
            chartStyle1.Append(plotArea2);
            chartStyle1.Append(plotArea3D1);
            chartStyle1.Append(seriesAxis1);
            chartStyle1.Append(seriesLine1);
            chartStyle1.Append(titleStyle1);
            chartStyle1.Append(trendlineStyle1);
            chartStyle1.Append(trendlineLabel1);
            chartStyle1.Append(upBar1);
            chartStyle1.Append(valueAxis2);
            chartStyle1.Append(wall1);

            chartStylePart1.ChartStyle = chartStyle1;
        }

        // Generates content of part.
        private void GeneratePartContent(WorksheetPart part)
        {
            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac xr xr2 xr3" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            worksheet1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            worksheet1.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
            worksheet1.AddNamespaceDeclaration("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
            worksheet1.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{00000000-0001-0000-0000-000000000000}"));
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A25:E32" };

            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            Selection selection1 = new Selection() { ActiveCell = "L22", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "L22" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 14.6D, DyDescent = 0.4D };

            SheetData sheetData1 = new SheetData();

            Row row1 = new Row() { RowIndex = (UInt32Value)25U, Spans = new ListValue<StringValue>() { InnerText = "1:5" }, DyDescent = 0.4D };

            Cell cell1 = new Cell() { CellReference = "A25", DataType = CellValues.SharedString };
            CellValue cellValue1 = new CellValue();
            cellValue1.Text = "0";

            cell1.Append(cellValue1);

            Cell cell2 = new Cell() { CellReference = "B25", DataType = CellValues.SharedString };
            CellValue cellValue2 = new CellValue();
            cellValue2.Text = "1";

            cell2.Append(cellValue2);

            Cell cell3 = new Cell() { CellReference = "C25", DataType = CellValues.SharedString };
            CellValue cellValue3 = new CellValue();
            cellValue3.Text = "2";

            cell3.Append(cellValue3);

            Cell cell4 = new Cell() { CellReference = "D25", DataType = CellValues.SharedString };
            CellValue cellValue4 = new CellValue();
            cellValue4.Text = "3";

            cell4.Append(cellValue4);

            Cell cell5 = new Cell() { CellReference = "E25", DataType = CellValues.SharedString };
            CellValue cellValue5 = new CellValue();
            cellValue5.Text = "4";

            cell5.Append(cellValue5);

            row1.Append(cell1);
            row1.Append(cell2);
            row1.Append(cell3);
            row1.Append(cell4);
            row1.Append(cell5);

            Row row2 = new Row() { RowIndex = (UInt32Value)26U, Spans = new ListValue<StringValue>() { InnerText = "1:5" }, DyDescent = 0.4D };

            Cell cell6 = new Cell() { CellReference = "A26" };
            CellValue cellValue6 = new CellValue();
            cellValue6.Text = "1";

            cell6.Append(cellValue6);

            Cell cell7 = new Cell() { CellReference = "B26" };
            CellValue cellValue7 = new CellValue();
            cellValue7.Text = "750.04499999999996";

            cell7.Append(cellValue7);

            Cell cell8 = new Cell() { CellReference = "C26" };
            CellValue cellValue8 = new CellValue();
            cellValue8.Text = "18.26923";

            cell8.Append(cellValue8);

            Cell cell9 = new Cell() { CellReference = "D26" };
            CellValue cellValue9 = new CellValue();
            cellValue9.Text = "1075.24";

            cell9.Append(cellValue9);

            Cell cell10 = new Cell() { CellReference = "E26" };
            CellValue cellValue10 = new CellValue();
            cellValue10.Text = "4.7214410000000004";

            cell10.Append(cellValue10);

            row2.Append(cell6);
            row2.Append(cell7);
            row2.Append(cell8);
            row2.Append(cell9);
            row2.Append(cell10);

            Row row3 = new Row() { RowIndex = (UInt32Value)27U, Spans = new ListValue<StringValue>() { InnerText = "1:5" }, DyDescent = 0.4D };

            Cell cell11 = new Cell() { CellReference = "A27" };
            CellValue cellValue11 = new CellValue();
            cellValue11.Text = "2";

            cell11.Append(cellValue11);

            Cell cell12 = new Cell() { CellReference = "B27" };
            CellValue cellValue12 = new CellValue();
            cellValue12.Text = "686.71";

            cell12.Append(cellValue12);

            Cell cell13 = new Cell() { CellReference = "C27" };
            CellValue cellValue13 = new CellValue();
            cellValue13.Text = "161.9091";

            cell13.Append(cellValue13);

            Cell cell14 = new Cell() { CellReference = "D27" };
            CellValue cellValue14 = new CellValue();
            cellValue14.Text = "1144.6420000000001";

            cell14.Append(cellValue14);

            Cell cell15 = new Cell() { CellReference = "E27" };
            CellValue cellValue15 = new CellValue();
            cellValue15.Text = "10.775539999999999";

            cell15.Append(cellValue15);

            row3.Append(cell11);
            row3.Append(cell12);
            row3.Append(cell13);
            row3.Append(cell14);
            row3.Append(cell15);

            Row row4 = new Row() { RowIndex = (UInt32Value)28U, Spans = new ListValue<StringValue>() { InnerText = "1:5" }, DyDescent = 0.4D };

            Cell cell16 = new Cell() { CellReference = "A28" };
            CellValue cellValue16 = new CellValue();
            cellValue16.Text = "4";

            cell16.Append(cellValue16);

            Cell cell17 = new Cell() { CellReference = "B28" };
            CellValue cellValue17 = new CellValue();
            cellValue17.Text = "690.04499999999996";

            cell17.Append(cellValue17);

            Cell cell18 = new Cell() { CellReference = "C28" };
            CellValue cellValue18 = new CellValue();
            cellValue18.Text = "173.05860000000001";

            cell18.Append(cellValue18);

            Cell cell19 = new Cell() { CellReference = "D28" };
            CellValue cellValue19 = new CellValue();
            cellValue19.Text = "1745.7850000000001";

            cell19.Append(cellValue19);

            Cell cell20 = new Cell() { CellReference = "E28" };
            CellValue cellValue20 = new CellValue();
            cellValue20.Text = "214.67490000000001";

            cell20.Append(cellValue20);

            row4.Append(cell16);
            row4.Append(cell17);
            row4.Append(cell18);
            row4.Append(cell19);
            row4.Append(cell20);

            Row row5 = new Row() { RowIndex = (UInt32Value)29U, Spans = new ListValue<StringValue>() { InnerText = "1:5" }, DyDescent = 0.4D };

            Cell cell21 = new Cell() { CellReference = "A29" };
            CellValue cellValue21 = new CellValue();
            cellValue21.Text = "8";

            cell21.Append(cellValue21);

            Cell cell22 = new Cell() { CellReference = "B29" };
            CellValue cellValue22 = new CellValue();
            cellValue22.Text = "609.10749999999996";

            cell22.Append(cellValue22);

            Cell cell23 = new Cell() { CellReference = "C29" };
            CellValue cellValue23 = new CellValue();
            cellValue23.Text = "122.4585";

            cell23.Append(cellValue23);

            Cell cell24 = new Cell() { CellReference = "D29" };
            CellValue cellValue24 = new CellValue();
            cellValue24.Text = "1865.405";

            cell24.Append(cellValue24);

            Cell cell25 = new Cell() { CellReference = "E29" };
            CellValue cellValue25 = new CellValue();
            cellValue25.Text = "212.03030000000001";

            cell25.Append(cellValue25);

            row5.Append(cell21);
            row5.Append(cell22);
            row5.Append(cell23);
            row5.Append(cell24);
            row5.Append(cell25);

            Row row6 = new Row() { RowIndex = (UInt32Value)30U, Spans = new ListValue<StringValue>() { InnerText = "1:5" }, DyDescent = 0.4D };

            Cell cell26 = new Cell() { CellReference = "A30" };
            CellValue cellValue26 = new CellValue();
            cellValue26.Text = "16";

            cell26.Append(cellValue26);

            Cell cell27 = new Cell() { CellReference = "B30" };
            CellValue cellValue27 = new CellValue();
            cellValue27.Text = "585.09749999999997";

            cell27.Append(cellValue27);

            Cell cell28 = new Cell() { CellReference = "C30" };
            CellValue cellValue28 = new CellValue();
            cellValue28.Text = "118.0836";

            cell28.Append(cellValue28);

            Cell cell29 = new Cell() { CellReference = "D30" };
            CellValue cellValue29 = new CellValue();
            cellValue29.Text = "1787.857";

            cell29.Append(cellValue29);

            Cell cell30 = new Cell() { CellReference = "E30" };
            CellValue cellValue30 = new CellValue();
            cellValue30.Text = "77.603099999999998";

            cell30.Append(cellValue30);

            row6.Append(cell26);
            row6.Append(cell27);
            row6.Append(cell28);
            row6.Append(cell29);
            row6.Append(cell30);

            Row row7 = new Row() { RowIndex = (UInt32Value)31U, Spans = new ListValue<StringValue>() { InnerText = "1:5" }, DyDescent = 0.4D };

            Cell cell31 = new Cell() { CellReference = "A31" };
            CellValue cellValue31 = new CellValue();
            cellValue31.Text = "32";

            cell31.Append(cellValue31);

            Cell cell32 = new Cell() { CellReference = "B31" };
            CellValue cellValue32 = new CellValue();
            cellValue32.Text = "426.45499999999998";

            cell32.Append(cellValue32);

            Cell cell33 = new Cell() { CellReference = "C31" };
            CellValue cellValue33 = new CellValue();
            cellValue33.Text = "9.2654160000000001";

            cell33.Append(cellValue33);

            Cell cell34 = new Cell() { CellReference = "D31" };
            CellValue cellValue34 = new CellValue();
            cellValue34.Text = "1678.998";

            cell34.Append(cellValue34);

            Cell cell35 = new Cell() { CellReference = "E31" };
            CellValue cellValue35 = new CellValue();
            cellValue35.Text = "58.671860000000002";

            cell35.Append(cellValue35);

            row7.Append(cell31);
            row7.Append(cell32);
            row7.Append(cell33);
            row7.Append(cell34);
            row7.Append(cell35);

            Row row8 = new Row() { RowIndex = (UInt32Value)32U, Spans = new ListValue<StringValue>() { InnerText = "1:5" }, DyDescent = 0.4D };

            Cell cell36 = new Cell() { CellReference = "A32" };
            CellValue cellValue36 = new CellValue();
            cellValue36.Text = "64";

            cell36.Append(cellValue36);

            Cell cell37 = new Cell() { CellReference = "B32" };
            CellValue cellValue37 = new CellValue();
            cellValue37.Text = "290.10250000000002";

            cell37.Append(cellValue37);

            Cell cell38 = new Cell() { CellReference = "C32" };
            CellValue cellValue38 = new CellValue();
            cellValue38.Text = "7.8593120000000001";

            cell38.Append(cellValue38);

            Cell cell39 = new Cell() { CellReference = "D32" };
            CellValue cellValue39 = new CellValue();
            cellValue39.Text = "1121.4349999999999";

            cell39.Append(cellValue39);

            Cell cell40 = new Cell() { CellReference = "E32" };
            CellValue cellValue40 = new CellValue();
            cellValue40.Text = "62.321089999999998";

            cell40.Append(cellValue40);

            row8.Append(cell36);
            row8.Append(cell37);
            row8.Append(cell38);
            row8.Append(cell39);
            row8.Append(cell40);

            sheetData1.Append(row1);
            sheetData1.Append(row2);
            sheetData1.Append(row3);
            sheetData1.Append(row4);
            sheetData1.Append(row5);
            sheetData1.Append(row6);
            sheetData1.Append(row7);
            sheetData1.Append(row8);
            PageMargins pageMargins2 = new PageMargins() { Left = 0.75D, Right = 0.75D, Top = 0.75D, Bottom = 0.5D, Header = 0.5D, Footer = 0.75D };
            Drawing drawing1 = new Drawing() { Id = "rId1" };

            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(pageMargins2);
            worksheet1.Append(drawing1);

            part.Worksheet = worksheet1;
        }

    }
}

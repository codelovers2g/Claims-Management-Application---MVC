using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Buisness;
using Common.Case.Template;
using System;
using Common.QuickBooks;
using Buisness.Services;
using Common.Case.Models;
using System.IO;
using Common.Master;
using System.Globalization;

using AA = DocumentFormat.OpenXml.Drawing;
using DWW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PICS = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using System.Text;
using System.Xml.Linq;

namespace SchedulingApp.CommonClass
{

    /// <summary>
    /// Provides extension methods for working with OpenXml document, particularly Word.
    /// </summary>
    public static class OpenXmlWordHelpers
    {
        /// <summary>
        /// Gets merge fields contained in a document, including the header and footer sections. 
        /// </summary>
        /// <param name="mergeFieldName">Optional name for the merge fields to look for.</param>
        /// <returns>If a merge field name is specified, only merge fields with that name are returned. Otherwise, it returns all merge fields contained in the document.</returns>
        public static IEnumerable<FieldCode> FillMergeFields(this WordprocessingDocument doc, Dictionary<string, string> pDictionaryMerge, NYCTemplate nYCTemplateObj)
        {
            if (doc == null)
                return null;

            /// Find all fields in HealthCarePractioner
            if (nYCTemplateObj.fieldNamesInGroup != null)
            {
                var textToFind = "«Provider.FullName»";
                var insertSuccessful = false;
                var totalCount = nYCTemplateObj.HealthCarePractioners?.Count;
                var position = 1;
                textToFind = "«" + nYCTemplateObj.fieldNamesInGroup[position] + "»";

            loopStart:
                IEnumerable<OpenXmlElement> elems = doc.MainDocumentPart.Document.Body.Descendants().ToList();
                foreach (OpenXmlElement elem in elems)
                {
                    if (elem is Text && (elem.InnerText == textToFind))
                    {
                        try
                        {
                            if (position < nYCTemplateObj.fieldNamesInGroup.Length - 1)
                            {
                                position++;
                                textToFind = "«" + nYCTemplateObj.fieldNamesInGroup[position] + "»";
                            }
                            Run run = (Run)elem.Parent;
                            Paragraph p = (Paragraph)run.GetParagraph();
                            var previous = p.PreviousSibling();

                            var PROP = "";
                            var ParaPROP = "";
                            if (!insertSuccessful)
                            {
                                foreach (RunProperties placeholderrpr in run.Descendants<RunProperties>().ToList())
                                {
                                    PROP = placeholderrpr.OuterXml;
                                    if (!string.IsNullOrEmpty(PROP))
                                        break;
                                }
                                foreach (RunProperties placeholderrpr in p.Descendants<RunProperties>().ToList())
                                {
                                    ParaPROP = placeholderrpr.OuterXml;
                                    if (!string.IsNullOrEmpty(PROP))
                                        break;
                                }
                                //var allfCodes = p.Descendants<FieldCode>();
                                //var allsFields = p.Descendants<SimpleField>();
                                //var par = run.Parent;
                                //p.Remove();
                                //p.Append(run);
                                string instructionText = String.Format(" MERGEFIELD  {0}  \\* MERGEFORMAT", "Provider.FullName");
                                SimpleField newSimpleField = new SimpleField() { Instruction = instructionText };

                                Run run1 = new Run();
                                //RunProperties runProperties1 = new RunProperties();
                                //NoProof noProof1 = new NoProof();
                                //runProperties1.Append(noProof1);
                                Text text1 = new Text();
                                text1.Text = String.Format("«{0}»", "Provider.FullName");

                                run1.Append(text1);
                                //run1.Append(runProperties1);
                                run1.PrependChild(new RunProperties(PROP));
                                newSimpleField.Append(run1);

                                /// New para graph
                                Paragraph para = new Paragraph();

                                p.RemoveAllChildren<Run>();
                                p.RemoveAllChildren<SimpleField>();

                                //var property = p.FirstChild;
                                //para.PrependChild(property);
                                p.Append(new OpenXmlElement[] { newSimpleField });

                                //var body = p.Parent;
                                /////p.InsertBefore(new, old);
                                //body.ReplaceChild(para, p);
                                insertSuccessful = true;

                                var hasTableStart = false;
                                foreach (var placeholderrpr in previous.Descendants<Text>().ToList())
                                {
                                    if (placeholderrpr.InnerText == "«TableStart:HealthCarePractitioners»")
                                    {
                                        hasTableStart = true;
                                    }
                                }

                                if (hasTableStart)
                                {
                                    previous.RemoveAllChildren();
                                    previous.Remove();
                                }

                                goto loopStart;
                            }
                            else
                            {
                                var hasProvider = false;
                                foreach (var placeholderrpr in p.Descendants<Text>().ToList())
                                {
                                    if (placeholderrpr.InnerText == "«Provider.FullName»")
                                    {
                                        hasProvider = true;
                                    }
                                }
                                if (!hasProvider)
                                {
                                    p.RemoveAllChildren();
                                    p.Remove();
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {

                            try
                            {
                                int _index = 0;
                                if (ex != null)
                                {
                                    CustomLogger.LogError.GetLogger.Log(ex, "Error in Replacing Fields in groupFunction.");
                                    QuickBookLogCQB _QuickBookLogCQB = new QuickBookLogCQB();
                                    _QuickBookLogCQB.AccessToken = " "; //acces_token variable is defined in AppController as this controller is inherited from AppController
                                    _QuickBookLogCQB.SyncRequest = "88888";
                                    _QuickBookLogCQB.RealmId = " ";
                                    _QuickBookLogCQB.CreatedDate = DateTime.UtcNow;
                                    _QuickBookLogCQB.SyncUptoDate = null;
                                    _QuickBookLogCQB.FailureReason = "Inner, Reason: " + Convert.ToString(ex?.InnerException?.Message == null ? ex?.Message : ex.InnerException?.Message);

                                    QuickBooksManager qb = new QuickBooksManager();
                                    bool SyncDetails = qb.QuickBooksSyncLog(_QuickBookLogCQB);
                                }
                            }
                            catch (Exception ex1)
                            {
                                CustomLogger.LogError.GetLogger.Log(ex1, "Error in AddDocument writing error Function");
                            }
                        }
                    }
                    else if (elem is Text && (elem.InnerText == "«TableEnd:HealthCarePractitioners»"))
                    {
                        try
                        {
                            Run run = (Run)elem.Parent;
                            Paragraph p = run.GetParagraph();
                            p.RemoveAllChildren();
                            p.Remove();
                        }
                        catch (Exception ex) { }
                    }
                }
            }

            List<Table> tables = doc.MainDocumentPart.Document.Body.Descendants<Table>().ToList();
            List<FieldCode> mergeFields = doc.MainDocumentPart.RootElement.Descendants<FieldCode>().ToList();
            foreach (var header in doc.MainDocumentPart.HeaderParts)
            {
                mergeFields.AddRange(header.RootElement.Descendants<FieldCode>());
            }
            foreach (var footer in doc.MainDocumentPart.FooterParts)
            {
                mergeFields.AddRange(footer.RootElement.Descendants<FieldCode>());
            }

            if (mergeFields != null && mergeFields.Count() > 0)
            {
                foreach (var mergeField in mergeFields)
                {
                    //var t = mergeField.Instruction?.InnerText.Replace("MERGEFIELD", "").Replace("\\* MERGEFORMAT", "").ToLower().Trim();
                    //var text = mergeField.InnerText.Replace("MERGEFIELD", "").Replace("\\* MERGEFORMAT", "").ToLower().Trim();
                    string result1 = string.Join(string.Empty, mergeField.InnerText.Replace("MERGEFIELD", "").Replace("\\* MERGEFORMAT", "").ToLower().Trim().Skip(5));
                    var result = pDictionaryMerge.Where(w => w.Key.ToLower().Substring(5) == result1).LastOrDefault();
                    if (result.Key == null)
                    {
                        result = pDictionaryMerge.Where(w => w.Key.ToLower().Contains(mergeField.InnerText.Replace("MERGEFIELD", "").Replace("\\* MERGEFORMAT", "").ToLower().Trim())).LastOrDefault();
                    }
                    if (result.Key != null)
                    {
                        try
                        {
                            nYCTemplateObj.Currentfield = result.Key;
                            mergeField.ReplaceWithText(result.Value.DefaultIfNull(), nYCTemplateObj);
                        }
                        catch (System.Exception ex)
                        {
                        }
                    }
                }
            }
            List<SimpleField> SimpleFields = doc.MainDocumentPart.RootElement.Descendants<SimpleField>().ToList();

            var count = 0;
            if (SimpleFields != null && SimpleFields.Count() > 0)
            {
                foreach (var SimpleField in SimpleFields)
                {
                    count++;
                    //var t = SimpleField.Instruction?.InnerText.Replace("MERGEFIELD", "").Replace("\\* MERGEFORMAT", "").ToLower().Trim();
                    //var text = SimpleField.InnerText.Replace("«", "").Replace("»", "").ToLower();
                    string result1 = string.Join(string.Empty, SimpleField.InnerText.Replace("«", "").Replace("»", "").ToLower().Skip(5));
                    var result = pDictionaryMerge.Where(w => w.Key.ToLower() == SimpleField.Instruction?.InnerText.Replace("MERGEFIELD", "").Replace("\\* MERGEFORMAT", "").ToLower().Trim()).LastOrDefault();
                    if (result.Key == null)
                    {
                        result = pDictionaryMerge.Where(w => w.Key.ToLower().Contains(SimpleField.InnerText.Replace("«", "").Replace("»", "").ToLower())).LastOrDefault();
                    }
                    if (result.Key != null && SimpleField.InnerText != "")
                    {
                        try
                        {
                            //nYCTemplateObj.Currentfield = result.Key;
                            //if (count > 18)
                            //{
                            //    if (SimpleField.InnerText.Contains("(39)"))
                            //    {
                            //        SimpleField.ReplaceWithText(result.Value.DefaultIfNull(), nYCTemplateObj);
                            //        pDictionaryMerge.Remove(result.Key);
                            //    }
                            //    else
                            //    {
                            //        SimpleField.ReplaceWithText(result.Value.DefaultIfNull(), nYCTemplateObj);
                            //    }
                            //}
                            //else
                            //{
                            SimpleField.ReplaceWithText(result.Value.DefaultIfNull(), nYCTemplateObj);
                            //}
                        }
                        catch (Exception e)
                        {
                        }

                    }
                }
            }
            IEnumerable<OpenXmlElement> elements = doc.MainDocumentPart.Document.Body.Descendants().ToList();
            foreach (OpenXmlElement elem in elements)
            {
                if (elem is Text && (elem.InnerText.Contains("~")) || (elem.InnerText.Contains("\n")))
                {
                    try
                    {
                        var ru = elem.Parent;
                        var isParagr = ru is Paragraph;
                        if ((ru is Run) == false)
                        {
                            continue;
                        }
                        var t = ru.GetType();
                        Run run = (Run)elem.Parent;
                        var PROP1 = "";

                        var parent = run.Parent;
                        var Para = parent.GetParagraph();

                        var PROP = "";
                        var ParaPROP = "";

                        foreach (RunProperties placeholderrpr in run.Descendants<RunProperties>().ToList())
                        {
                            PROP = placeholderrpr.OuterXml;
                            break;
                        }

                        foreach (RunProperties placeholderrpr in Para.Descendants<RunProperties>())
                        {
                            ParaPROP = placeholderrpr.OuterXml;
                            break;
                        }
                        /// Remove Current text and append new with line breaks
                        foreach (var element in run.Descendants<Text>())
                        {
                            element.Remove();
                            break;
                        }

                        var containsNewLines = false;

                        string[] AddressPart = new string[] { };
                        if (elem.InnerText.Contains("\n"))
                        {
                            AddressPart = elem.InnerText.Split('\n');
                            containsNewLines = true;
                        }
                        else if (elem.InnerText.Contains("~"))
                        {
                            AddressPart = elem.InnerText.Split('~');
                        }
                        else
                        {
                            AddressPart = elem.InnerText.Split('~');
                        }
                        var partCount = 0;
                        //foreach (string line in AddressPart)
                        //{
                        //    partCount++;
                        //    if (!string.IsNullOrEmpty(line))
                        //    {
                        //        run.Append(new Text(line));
                        //        if (partCount < AddressPart.Length)
                        //        {
                        //            run.Append(new Break());
                        //        }
                        //    }
                        //}
                        foreach (string line in AddressPart)
                        {
                            partCount++;
                            if (!string.IsNullOrEmpty(line))
                            {
                                run.Append(new Text(line));
                                if (partCount < AddressPart.Length)
                                {
                                    run.Append(new Break());
                                }
                            }
                            else
                            {
                                if (containsNewLines)
                                    run.Append(new Break());
                            }
                        }
                        //elem.InnerText = "";
                        //RunCreated.PrependChild(new RunProperties(PROP1));

                        //var allfCodes = Para.Descendants<FieldCode>().ToList();
                        //var allsFields = Para.Descendants<SimpleField>().ToList();
                        //var T = parent.GetType();
                        //if (T.FullName.Contains("SimpleField"))
                        //{
                        //    //    parent.Remove();
                        //    //    //Para.ReplaceChild(parent, RunCreated);
                        //    //    Para.AppendChild<Run>(RunCreated);

                        //}
                        //else
                        //{
                        //}
                        //parent.Remove();
                        //Para.ReplaceChild(parent, RunCreated);
                        //Para.AppendChild<Run>(RunCreated);


                    }
                    catch (System.Exception ex)
                    {

                    }
                    //p.RemoveAllChildren();
                    ////p.Remove();
                    //p.Append(run);
                }
                else if (elem is Text && (elem.InnerText.Contains("»")))
                {

                    Run run = (Run)elem.Parent;

                    foreach (var element in run.Descendants<Text>())
                    {
                        element.Remove();
                        break;
                    }
                }
            }
            return mergeFields;
        }

        public static IEnumerable<FieldCode> FillMergeFieldsForReports(this WordprocessingDocument doc, List<IMEDataReport2Updated> _IMEDataReport2, IMEReportFiltter IMEReportFiltter)
        {
            if (doc == null)
                return null;


            List<Table> tables = doc.MainDocumentPart.Document.Body.Descendants<Table>().ToList();
            Table table = new Table();
            Table template = new Table();
            Table HeaderTable = new Table();

            if (IMEReportFiltter.isDailyStatusReport)
            {
                HeaderTable = doc.MainDocumentPart.Document.Body.Descendants<Table>().ElementAt(0);
                template = doc.MainDocumentPart.Document.Body.Descendants<Table>().ElementAt(1);
                table = doc.MainDocumentPart.Document.Body.Descendants<Table>().ElementAt(1);
            }
            else
            { 
                table = doc.MainDocumentPart.Document.Body.Descendants<Table>().FirstOrDefault();
            }
            
            List<FieldCode> mergeFields = doc.MainDocumentPart.RootElement.Descendants<FieldCode>().ToList();
            foreach (var header in doc.MainDocumentPart.HeaderParts)
            {
                mergeFields.AddRange(header.RootElement.Descendants<FieldCode>());
            }
            foreach (var footer in doc.MainDocumentPart.FooterParts)
            {
                var foot = footer.RootElement.Descendants<FieldCode>();

                var dt = DateTime.UtcNow;
                var df = dt.ToLongDateString();

                if (footer.Footer.InnerText.Contains("Tuesday, June 2, 2020"))
                {
                    //footer.Footer.InnerText.Insert(0 "");
                }

                mergeFields.AddRange(footer.RootElement.Descendants<FieldCode>());
            }


            Dictionary<string, string> pDictionaryMerge = new Dictionary<string, string>();
            try
            {

                pDictionaryMerge.Add("DailyStatusSheetDate", IMEReportFiltter.For != null ? Convert.ToDateTime(IMEReportFiltter.For).ToString("MM/dd/yyyy") : "N/A");
                pDictionaryMerge.Add("DoctorName", _IMEDataReport2.FirstOrDefault()?.Doctor ?? "");
                pDictionaryMerge.Add("ClinicName", _IMEDataReport2.FirstOrDefault()?.Location ?? "");
                pDictionaryMerge.Add("TotalExaminations", "Total Examinations:      " + _IMEDataReport2.Count.ToString());


            }
            catch (Exception ex)
            {

            }
            foreach (var table11 in tables)
            {

                foreach (var TB in table11.Descendants<Text>().ToList())
                {

                    if (TB.InnerText == "Totalexamination ")
                    {

                    }
                    if (TB.InnerText == "Totalexamination")
                    {

                    }
                }
            }

            if (IMEReportFiltter.isDailyStatusReport)
            {
                var nYCTemplateObj = new NYCTemplate();

                if (mergeFields != null && mergeFields.Count() > 0)
                {
                    foreach (var mergeField in mergeFields)
                    {
                        string result1 = string.Join(string.Empty, mergeField.InnerText.Replace("MERGEFIELD", "").Replace("\\* MERGEFORMAT", "").ToLower().Trim());
                        var result = pDictionaryMerge.Where(w => w.Key.ToLower() == result1).LastOrDefault();
                        if (result.Key == null)
                        {
                            result = pDictionaryMerge.Where(w => w.Key.ToLower().Contains(mergeField.InnerText.Replace("MERGEFIELD", "").Replace("\\* MERGEFORMAT", "").ToLower().Trim())).LastOrDefault();
                        }
                        if (result.Key != null)
                        {
                            try
                            {
                                mergeField.ReplaceWithText(result.Value.DefaultIfNull(), nYCTemplateObj);
                            }
                            catch (System.Exception ex)
                            {
                            }
                        }
                    }
                }
                List<SimpleField> SimpleFields = doc.MainDocumentPart.RootElement.Descendants<SimpleField>().ToList();

                if (SimpleFields != null && SimpleFields.Count() > 0)
                {
                    foreach (var SimpleField in SimpleFields)
                    {
                        string result1 = string.Join(string.Empty, SimpleField.InnerText.Replace("«", "").Replace("»", "").ToLower());
                        var result = pDictionaryMerge.Where(w => w.Key.ToLower() == SimpleField.Instruction?.InnerText.Replace("MERGEFIELD", "").Replace("\\* MERGEFORMAT", "").ToLower().Trim()).LastOrDefault();
                        if (result.Key == null)
                        {
                            result = pDictionaryMerge.Where(w => w.Key.ToLower().Contains(SimpleField.InnerText.Replace("«", "").Replace("»", "").ToLower())).LastOrDefault();
                        }
                        if (result.Key != null && SimpleField.InnerText != "")
                        {
                            try
                            {
                                SimpleField.ReplaceWithText(result.Value.DefaultIfNull(), nYCTemplateObj);
                            }
                            catch (Exception e)
                            {
                            }

                        }
                    }
                }
            }
            string CellProperties = string.Empty, LastCellProperties = string.Empty, SecondCellProperties = string.Empty, ThirdCellProperties = string.Empty,
                FourthCellProperties = string.Empty, FifthCellProperties = string.Empty, EighthCellProperties = string.Empty;
            if (tables.Count > 0)
            {
                /// For appointment Log
                if (IMEReportFiltter.isAppointmentLogReport)
                {
                    string AppointmentsFromTo = "";

                    if (IMEReportFiltter.FromIsValid == false && IMEReportFiltter.ToIsValid == false)
                    {
                        AppointmentsFromTo = "All Appointments";
                    }
                    else
                    {
                        if (IMEReportFiltter.FromIsValid)
                        {
                            AppointmentsFromTo += "Appointments From " + (IMEReportFiltter.From != null ? Convert.ToDateTime(IMEReportFiltter.From).ToString("MM/dd/yyyy") : "N/A");

                        }
                        if (IMEReportFiltter.ToIsValid)
                        {
                            if (IMEReportFiltter.FromIsValid)
                            {
                                AppointmentsFromTo += " To " + (IMEReportFiltter.To != null ? Convert.ToDateTime(IMEReportFiltter.To).ToString("MM/dd/yyyy") : "N/A");
                            }
                            else
                            {
                                AppointmentsFromTo += " Appointments till " + (IMEReportFiltter.To != null ? Convert.ToDateTime(IMEReportFiltter.To).ToString("MM/dd/yyyy") : "N/A");
                            }
                        }
                    }

                    pDictionaryMerge.Add("AppointmentsFromTo", AppointmentsFromTo);


                    var TR1 = table.Descendants<TableRow>().Skip(2).FirstOrDefault();//skip(1)
                    var TRwithdefaultData = table.Descendants<TableRow>().Skip(1).FirstOrDefault();

                    if (TR1.Descendants<TableCell>().ToList().Count > 0)
                    {
                        var FirstCell = TR1.Descendants<TableCell>().FirstOrDefault();
                        CellProperties = FirstCell.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;

                        var SecondCell = TR1.Descendants<TableCell>().Skip(1).FirstOrDefault();
                        SecondCellProperties = SecondCell.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;


                        var ThirdCell = TR1.Descendants<TableCell>().Skip(2).FirstOrDefault();
                        ThirdCellProperties = ThirdCell.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;


                        var FourthCell = TR1.Descendants<TableCell>().Skip(3).FirstOrDefault();
                        FourthCellProperties = FourthCell.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;

                        var FifthCell = TR1.Descendants<TableCell>().Skip(4).FirstOrDefault();
                        FifthCellProperties = FifthCell.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;

                        //var EighthCell = TR1.Descendants<TableCell>().Skip(7).FirstOrDefault();
                        //EighthCellProperties = EighthCell.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;
                        ////var LastCell = TR1.Descendants<TableCell>().Skip(1).LastOrDefault();
                        //LastCellProperties = SecondCell.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;
                    }
                    TR1.Remove();
                    TRwithdefaultData.Remove();

                    foreach (var record in _IMEDataReport2)
                    {
                        TableRow tr = new TableRow();
                        // Create a cell.
                        TableCell tc1 = new TableCell();
                        TableCell tc2 = new TableCell();
                        TableCell tc3 = new TableCell();
                        TableCell tc4 = new TableCell();
                        TableCell tc5 = new TableCell();
                        TableCell tc6 = new TableCell();
                        TableCell tc7 = new TableCell();
                        TableCell tc8 = new TableCell();
                        TableCell tc9 = new TableCell();

                        // Specify the width property of the table cell.
                        tc1.Append(new TableCellProperties(CellProperties));
                        tc2.Append(new TableCellProperties(CellProperties));
                        tc3.Append(new TableCellProperties(CellProperties));
                        tc4.Append(new TableCellProperties(CellProperties));
                        tc5.Append(new TableCellProperties(CellProperties));
                        tc6.Append(new TableCellProperties(CellProperties));
                        tc7.Append(new TableCellProperties(CellProperties));
                        tc8.Append(new TableCellProperties(CellProperties));
                        tc9.Append(new TableCellProperties(CellProperties));
                        //tc1.Append(new TableCellProperties(
                        //    new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));

                        // Specify the table cell content.
                        tc1.Append(new Paragraph(new Run(new Text(record.casenumber ?? " "))));
                        tc2.Append(new Paragraph(new Run(new Text(record.ClaimantFullName ?? " "))));
                        tc3.Append(new Paragraph(new Run(new Text(record.Service ?? " "))));
                        tc4.Append(new Paragraph(new Run(new Text(record.Scheduler ?? " "))));
                        tc5.Append(new Paragraph(new Run(new Text(record.client ?? " "))));
                        tc6.Append(new Paragraph(new Run(new Text(record.company ?? " "))));
                        tc7.Append(new Paragraph(new Run(new Text(record.AppointmentDate.ToString() ?? " "))));
                        tc8.Append(new Paragraph(new Run(new Text(record.Doctor.ToString() ?? " "))));
                        tc9.Append(new Paragraph(new Run(new Text(record.Location.ToString() ?? " "))));

                        tc9.Descendants<Run>().FirstOrDefault().AppendChild(new Break());
                        // Append the table cell to the table row.
                        tr.Append(tc1);
                        tr.Append(tc2);
                        tr.Append(tc3);
                        tr.Append(tc4);
                        tr.Append(tc5);
                        tr.Append(tc6);
                        tr.Append(tc7);
                        tr.Append(tc8);
                        tr.Append(tc9);
                        table.Append(tr);
                    }
                }
                /// For Daily status
                else if (IMEReportFiltter.isDailyStatusReport)
                {
                    var TR1 = table.Descendants<TableRow>().Skip(1).FirstOrDefault();

                    var TRTotalExaminationsRow = table.Descendants<TableRow>().Skip(7).FirstOrDefault();
                    //var TRTotalExaminationsRow1 = table.Descendants<TableRow>().Skip(8).FirstOrDefault();

                    foreach (var TB in doc.MainDocumentPart.Document.Body.Descendants<Text>().ToList())
                    {

                        if (TB.InnerText == "Totalexamination ")
                        {

                        }
                        if (TB.InnerText == "Totalexamination")
                        {

                        }
                    }

                    var rowProperties = TR1.Descendants<TableRowProperties>().FirstOrDefault().OuterXml;


                    if (TR1.Descendants<TableCell>().ToList().Count > 0)
                    {
                        var FirstCell = TR1.Descendants<TableCell>().FirstOrDefault();
                        CellProperties = FirstCell.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;

                        var SecondCell = TR1.Descendants<TableCell>().Skip(1).FirstOrDefault();
                        SecondCellProperties = SecondCell.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;

                        if (TR1.Descendants<TableCell>().ToList().Count > 0)
                        {
                            var FirstCellD = TR1.Descendants<TableCell>().FirstOrDefault();
                            CellProperties = FirstCellD.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;

                            var SecondCellD = TR1.Descendants<TableCell>().Skip(1).FirstOrDefault();
                            SecondCellProperties = SecondCellD.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;


                            var ThirdCellD = TR1.Descendants<TableCell>().Skip(2).FirstOrDefault();
                            ThirdCellProperties = ThirdCellD.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;


                            var FourthCellD = TR1.Descendants<TableCell>().Skip(3).FirstOrDefault();
                            FourthCellProperties = FourthCellD.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;

                            var FifthCellD = TR1.Descendants<TableCell>().Skip(4).FirstOrDefault();
                            FifthCellProperties = FifthCellD.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;

                            var EighthCellD = TR1.Descendants<TableCell>().Skip(9).FirstOrDefault();
                            EighthCellProperties = EighthCellD.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;
                            //var LastCell = TR1.Descendants<TableCell>().Skip(1).LastOrDefault();
                            //LastCellProperties = SecondCell.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;
                        }
                    }

                    //var tABList = TR1.Descendants<Table>().ToList();
                    var whiteTable = TR1.Descendants<Table>().FirstOrDefault();
                    var blackTable = TR1.Descendants<Table>().Skip(1).FirstOrDefault();
                    OpenXmlElement clonedWTable = whiteTable.CloneNode(true);
                    OpenXmlElement clonedBTable = blackTable.CloneNode(true);

                    var whiteTableCell = TR1.Descendants<TableCell>().Skip(4).FirstOrDefault();

                    var whiteCellProperties = whiteTableCell.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;
                    var whiteCellPara = whiteTableCell.Descendants<Paragraph>().FirstOrDefault();
                    OpenXmlElement clonedSmallPara = whiteCellPara.CloneNode(true);


                    var XML = whiteTable.InnerXml;
                    TR1.RemoveAllChildren();
                    TR1.Remove();


                    Table currentTable = table;
                    var tempDoc = "";
                    var tempDocLoc = "";
                    int count = 0;
                    int docTotalAppCount = 0;
                    foreach (var record in _IMEDataReport2)
                    {


                        if (tempDoc != record.Doctor || tempDocLoc != record.Location)
                        {
                            count++;
                            Table tableHeader = new Table();

                            // Create a TableProperties object and specify its border information.
                            TableProperties tblProp = new TableProperties(
                                new TableBorders(
                                    new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 8 },
                                    new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 8 },
                                    new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 8 },
                                    new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 8 }
                                ),
                                new TableCellBorders(
                                    new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 8 },
                                    new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 8 },
                                    new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 8 },
                                    new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 8 }
                                )
                            //new TableRow(
                            //    new TableRowHeight() { Val = Convert.ToUInt32("20") }
                            //)
                            );
                            Paragraph paragraph1 = new Paragraph();
                            Paragraph paragraph2 = new Paragraph();
                            ParagraphProperties paraProperties = new ParagraphProperties();
                            Justification justification = new Justification() { Val = JustificationValues.Center };
                            RunProperties runProperties = new RunProperties();
                            FontSize fontSize = new FontSize();
                            fontSize.Val = "28";
                            Bold bold1 = new Bold();
                            bold1.Val = true;
                            runProperties.Append(new RunFonts() { Ascii = "Times New Roman" });
                            runProperties.Append(fontSize);
                            runProperties.Append(bold1);
                            paraProperties.Append(justification);

                            ParagraphProperties paraProperties2 = new ParagraphProperties();
                            Justification justification2 = new Justification() { Val = JustificationValues.Center };
                            RunProperties runProperties2 = new RunProperties();
                            FontSize fontSize2 = new FontSize();
                            fontSize2.Val = "28";
                            Bold bold2 = new Bold();
                            bold2.Val = true;
                            runProperties2.Append(new RunFonts() { Ascii = "Times New Roman" });
                            runProperties2.Append(fontSize2);
                            runProperties2.Append(bold2);
                            paraProperties2.Append(justification2);

                            // Append the TableProperties object to the empty table.
                            tableHeader.AppendChild<TableProperties>(tblProp);

                            // Now we create a new layout and make it "fixed".
                            TableLayout tl = new TableLayout() { Type = TableLayoutValues.Autofit };
                            tblProp.TableLayout = tl;

                            TableRow trHeader = new TableRow();

                            // Create a cell.
                            TableCell tcHeader1 = new TableCell();
                            TableCell tcHeader2 = new TableCell();
                            //TableCell tc9 = new TableCell();

                            //TableCellVerticalAlignment tcVA = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                            // Specify the width property of the table cell.
                            tcHeader1.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "4800" }), new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }, new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 8 });
                            tcHeader2.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "4800" }), new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center });
                            //create Run and Text 
                            Run run1 = new Run();
                            Text text1 = new Text();

                            //add content in Text
                            text1.Text = record.Doctor.Trim() ?? " ";

                            //add Text to Run
                            run1.Append(runProperties);
                            run1.Append(text1);

                            //add Run to paragraph
                            paragraph1.Append(paraProperties);
                            paragraph1.Append(run1);
                            Run run2 = new Run();
                            Text text2 = new Text();

                            //add content in Text
                            text2.Text = record.Location.Trim() ?? " ";

                            //add Text to Run
                            run2.Append(runProperties2);
                            run2.Append(text2);

                            //add Run to paragraph
                            paragraph2.Append(paraProperties2);
                            paragraph2.Append(run2);

                            // Specify the table cell content.
                            tcHeader1.Append(paragraph1);
                            tcHeader2.Append(paragraph2);

                            TableRowProperties tableRowProperties1 = new TableRowProperties();
                            TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = (UInt32Value)558U };

                            tableRowProperties1.Append(tableRowHeight1);
                            trHeader.Append(tableRowProperties1);

                            trHeader.Append(tcHeader1);
                            trHeader.Append(tcHeader2);
                            //tr.Append(tc9);
                            tableHeader.Append(trHeader);

                            if (count >= 1)
                            {
                                table = (Table)doc.MainDocumentPart.Document.Body.InsertAfter(template.CloneNode(true), currentTable);
                                int countRow = 1;
                                int rowss = table.Elements<TableRow>().Count();
                                foreach (var row in table.Elements<TableRow>())
                                {
                                    if (countRow >= 2)
                                    {
                                        table.RemoveChild<TableRow>(row);
                                    }
                                    countRow++;
                                }
                                currentTable = table;
                            }
                            doc.MainDocumentPart.Document.Body.InsertBefore(tableHeader, currentTable);
                            if (count >= 2)
                            {

                                Paragraph paraCount = new Paragraph();
                                Run runCount = new Run();
                                //add Text to Run
                                runCount.Append(new Break());
                                runCount.Append(new Break());
                                runCount.Append(new Break());
                                ParagraphProperties paraProperty = new ParagraphProperties();
                                RunProperties runProperty = new RunProperties();
                                Bold boldCount = new Bold();
                                boldCount.Val = true;
                                Italic italicCount = new Italic();
                                italicCount.Val = true;
                                runProperty.Append(new RunFonts() { Ascii = "Times New Roman" });
                                runProperty.Append(italicCount);
                                runProperty.Append(boldCount);
                                runCount.Append(runProperty);
                                runCount.Append(new Text("TotalCount:   " + docTotalAppCount));
                                //add Run to paragraph
                                paraCount.Append(runCount);

                                doc.MainDocumentPart.Document.Body.InsertBefore(paraCount, tableHeader);
                                docTotalAppCount = 0;

                                var tempHeaderTable = (Table)doc.MainDocumentPart.Document.Body.InsertBefore(HeaderTable.CloneNode(true), tableHeader);

                                Paragraph para = new Paragraph();
                                Run run = new Run();
                                //add Text to Run
                                run.Append(new Break() { Type = BreakValues.Page });
                                //add Run to paragraph
                                para.Append(run);
                                doc.MainDocumentPart.Document.Body.InsertBefore(para, tempHeaderTable);


                                Paragraph paraHeader = new Paragraph();
                                Run runHeader = new Run();
                                //add Text to Run
                                runHeader.Append(new Text(""));
                                //add Run to paragraph
                                paraHeader.Append(runHeader);

                                doc.MainDocumentPart.Document.Body.InsertAfter(paraHeader, tempHeaderTable);


                            }
                        }

                        docTotalAppCount++;
                        tempDoc = record.Doctor;
                        tempDocLoc = record.Location;

                        TableRow tr = new TableRow();

                        // Create a cell.
                        TableCell tc1 = new TableCell();
                        TableCell tc2 = new TableCell();
                        TableCell tc3 = new TableCell();
                        TableCell tc4 = new TableCell();
                        TableCell tc5 = new TableCell();
                        TableCell tc6 = new TableCell();
                        TableCell tc7 = new TableCell();
                        TableCell tc8 = new TableCell();
                        //TableCell tc9 = new TableCell();

                        // Specify the width property of the table cell.
                        tc1.Append(new TableCellProperties(CellProperties));
                        tc2.Append(new TableCellProperties(SecondCellProperties));
                        tc3.Append(new TableCellProperties(ThirdCellProperties));
                        tc4.Append(new TableCellProperties(FourthCellProperties));
                        tc5.Append(new TableCellProperties(whiteCellProperties));
                        tc6.Append(new TableCellProperties(whiteCellProperties));
                        tc7.Append(new TableCellProperties(whiteCellProperties));
                        tc8.Append(new TableCellProperties(EighthCellProperties));

                        // Specify the table cell content.
                        tc1.Append(new Paragraph(new Run(new Text(record.AppointmentTimeString ?? " "))));
                        tc2.Append(new Paragraph(new Run(new Text(record.ClaimantFullName ?? " "))));
                        tc3.Append(new Paragraph(new Run(new Text(record.casenumber ?? " ")), new Run(new Break())));
                        tc4.Append(new Paragraph(new Run(new Text(record.Service ?? " "))));
                        tc5.Append(clonedSmallPara.CloneNode(true));
                        tc6.Append(clonedSmallPara.CloneNode(true));
                        tc7.Append(clonedSmallPara.CloneNode(true));
                        tc5.Append((record.ShowsUp == true ? clonedBTable.CloneNode(true) : clonedWTable.CloneNode(true)));
                        tc6.Append((record.NoShow == true ? clonedBTable.CloneNode(true) : clonedWTable.CloneNode(true)));
                        tc7.Append((record.UnableToExamine == true ? clonedBTable.CloneNode(true) : clonedWTable.CloneNode(true))); tc8.Append(new Paragraph(new Run(new Text(record.Comments ?? " "))));

                        tr.Append(tc1);
                        tr.Append(tc2);
                        tr.Append(tc3);
                        tr.Append(tc4);
                        tr.Append(tc5);
                        tr.Append(tc6);
                        tr.Append(tc7);
                        tr.Append(tc8);
                        //tr.Append(tc9);
                        currentTable.Append(tr);
                        if (_IMEDataReport2.IndexOf(record) == _IMEDataReport2.Count - 1)
                        {
                            Paragraph paraCount = new Paragraph();
                            Run runCount = new Run();
                            //add Text to Run
                            ParagraphProperties paraProperty = new ParagraphProperties();
                            RunProperties runProperty = new RunProperties();
                            Bold boldCount = new Bold();
                            boldCount.Val = true;
                            Italic italicCount = new Italic();
                            italicCount.Val = true;
                            runProperty.Append(new RunFonts() { Ascii = "Times New Roman" });
                            runProperty.Append(italicCount);
                            runProperty.Append(boldCount);
                            runCount.Append(runProperty);
                            runCount.Append(new Text("TotalCount:   " + docTotalAppCount));
                            //add Run to paragraph
                            paraCount.Append(runCount);

                            doc.MainDocumentPart.Document.Body.Append(paraCount);
                            docTotalAppCount = 0;
                        }
                    }
                }

                foreach (var row in template.Elements<TableRow>())
                {
                    template.RemoveChild<TableRow>(row);
                }

                //var TRTotalExaminations = table.Descendants<TableRow>().Skip(7).FirstOrDefault();

                //        try
                //        {
                //            foreach (var record in _IMEDataReport2)
                //            {
                //                TableRow tr = new TableRow();
                //                tr.Append(new TableRowProperties(rowProperties));

                //                // Create a cell.
                //                TableCell tc1 = new TableCell();
                //                TableCell tc2 = new TableCell();
                //                TableCell tc3 = new TableCell();
                //                TableCell tc4 = new TableCell();
                //                TableCell tc5 = new TableCell();
                //                TableCell tc6 = new TableCell();
                //                TableCell tc7 = new TableCell();
                //                TableCell tc8 = new TableCell();
                //                //TableCell tc9 = new TableCell();

                //                // Specify the width property of the table cell.
                //                tc1.Append(new TableCellProperties(CellProperties));
                //                tc2.Append(new TableCellProperties(SecondCellProperties));
                //                tc3.Append(new TableCellProperties(ThirdCellProperties));
                //                tc4.Append(new TableCellProperties(FourthCellProperties));
                //                tc5.Append(new TableCellProperties(whiteCellProperties));
                //                tc6.Append(new TableCellProperties(whiteCellProperties));
                //                tc7.Append(new TableCellProperties(whiteCellProperties));
                //                tc8.Append(new TableCellProperties(EighthCellProperties));
                //                //tc9.Append(new TableCellProperties(CellProperties));
                //                //tc1.Append(new TableCellProperties(
                //                //    new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));

                //                //// Create an empty table.
                //                //Table squareTable = new Table();

                //                //// Create a TableProperties object and specify its border information.
                //                //TableProperties tblProp = new TableProperties(
                //                //    new TableBorders(
                //                //        new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 8 },
                //                //        new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 8 },
                //                //        new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 8 },
                //                //        new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 8 },
                //                //        new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 8 },
                //                //        new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 8 }
                //                //    )
                //                //);
                //                //// Append the TableProperties object to the empty table.
                //                //squareTable.AppendChild<TableProperties>(tblProp);



                //                // Specify the table cell content.
                //                tc1.Append(new Paragraph(new Run(new Text(record.AppointmentTimeString ?? " "))));
                //                tc2.Append(new Paragraph(new Run(new Text(record.ClaimantFullName ?? " "))));
                //                tc3.Append(new Paragraph(new Run(new Text(record.casenumber ?? " "))));
                //                tc4.Append(new Paragraph(new Run(new Text(record.Service ?? " "))));


                //                tc5.Append(clonedSmallPara.CloneNode(true));
                //                tc6.Append(clonedSmallPara.CloneNode(true));
                //                tc7.Append(clonedSmallPara.CloneNode(true));
                //                tc5.Append((record.ShowsUp == true ? clonedBTable.CloneNode(true) : clonedWTable.CloneNode(true)));
                //                tc6.Append((record.NoShow == true ? clonedBTable.CloneNode(true) : clonedWTable.CloneNode(true)));
                //                tc7.Append((record.UnableToExamine == true ? clonedBTable.CloneNode(true) : clonedWTable.CloneNode(true)));
                //                //tc5.Append(new Paragraph(new Run(new Text(record.Show.ToString() ?? " "))));
                //                //tc6.Append(new Paragraph(new Run(new Text(record.Show.ToString() ?? " "))));
                //                //tc7.Append(new Paragraph(new Run(new Text(record.ShowsUp.ToString() ?? " "))));
                //                //tc6.Append(new Paragraph((clonedWTable.CloneNode(true))));

                //                //var NEWTAB = new Table();
                //                //NEWTAB.InnerXml = XML;

                //                //tc6.Append(NEWTAB);
                //                //tc6.Append(new Paragraph(new Run()));
                //                //tc7.Append((clonedWTable.CloneNode(true)));

                //                tc8.Append(new Paragraph(new Run(new Text(record.Comments ?? " "))));
                //                //tc9.Append(new Paragraph(new Run(new Text(record.Location.ToString()?? " "))));

                //                //tc9.Descendants<Run>().FirstOrDefault().AppendChild(new Break());
                //                // Append the table cell to the table row.
                //                tr.Append(tc1);
                //                tr.Append(tc2);
                //                tr.Append(tc3);
                //                tr.Append(tc4);
                //                tr.Append(tc5);
                //                tr.Append(tc6);
                //                tr.Append(tc7);
                //                tr.Append(tc8);
                //                //tr.Append(tc9);
                //                table.Append(tr);
                //            }
                //        }
                //        catch (Exception e)
                //        {

                //        }
                //    } 
            }

            if (IMEReportFiltter.isAppointmentLogReport)
            {
                var nYCTemplateObj = new NYCTemplate();

                if (mergeFields != null && mergeFields.Count() > 0)
                {
                    foreach (var mergeField in mergeFields)
                    {
                        string result1 = string.Join(string.Empty, mergeField.InnerText.Replace("MERGEFIELD", "").Replace("\\* MERGEFORMAT", "").ToLower().Trim());
                        var result = pDictionaryMerge.Where(w => w.Key.ToLower() == result1).LastOrDefault();
                        if (result.Key == null)
                        {
                            result = pDictionaryMerge.Where(w => w.Key.ToLower().Contains(mergeField.InnerText.Replace("MERGEFIELD", "").Replace("\\* MERGEFORMAT", "").ToLower().Trim())).LastOrDefault();
                        }
                        if (result.Key != null)
                        {
                            try
                            {
                                mergeField.ReplaceWithText(result.Value.DefaultIfNull(), nYCTemplateObj);
                            }
                            catch (System.Exception ex)
                            {

                            }
                        }
                    }
                }
                List<SimpleField> SimpleFields = doc.MainDocumentPart.RootElement.Descendants<SimpleField>().ToList();

                if (SimpleFields != null && SimpleFields.Count() > 0)
                {
                    foreach (var SimpleField in SimpleFields)
                    {
                        string result1 = string.Join(string.Empty, SimpleField.InnerText.Replace("«", "").Replace("»", "").ToLower());
                        var result = pDictionaryMerge.Where(w => w.Key.ToLower() == SimpleField.Instruction?.InnerText.Replace("MERGEFIELD", "").Replace("\\* MERGEFORMAT", "").ToLower().Trim()).LastOrDefault();
                        if (result.Key == null)
                        {
                            result = pDictionaryMerge.Where(w => w.Key.ToLower().Contains(SimpleField.InnerText.Replace("«", "").Replace("»", "").ToLower())).LastOrDefault();
                        }
                        if (result.Key != null && SimpleField.InnerText != "")
                        {
                            try
                            {
                                SimpleField.ReplaceWithText(result.Value.DefaultIfNull(), nYCTemplateObj);
                            }
                            catch (Exception e)
                            {

                            }

                        }
                    }
                }
            }

            return mergeFields;
        }
        public static IEnumerable<FieldCode> FillMergeFieldsForExportReport(this WordprocessingDocument doc, List<CaseReport> _IMEDataReport2, IMEReportFiltter IMEReportFiltter)
        {
            if (doc == null)
                return null;


            List<Table> tables = doc.MainDocumentPart.Document.Body.Descendants<Table>().ToList();
            Table table = new Table();
            Table template = new Table();
            Table HeaderTable = new Table();

            HeaderTable = doc.MainDocumentPart.Document.Body.Descendants<Table>().ElementAt(0);
            template = doc.MainDocumentPart.Document.Body.Descendants<Table>().ElementAt(1);
            table = doc.MainDocumentPart.Document.Body.Descendants<Table>().ElementAt(1);

            List<FieldCode> mergeFields = doc.MainDocumentPart.RootElement.Descendants<FieldCode>().ToList();
            foreach (var header in doc.MainDocumentPart.HeaderParts)
            {
                mergeFields.AddRange(header.RootElement.Descendants<FieldCode>());
            }
            foreach (var footer in doc.MainDocumentPart.FooterParts)
            {
                mergeFields.AddRange(footer.RootElement.Descendants<FieldCode>());
            }

            Dictionary<string, string> pDictionaryMerge = new Dictionary<string, string>();

            string CellProperties = string.Empty, LastCellProperties = string.Empty, SecondCellProperties = string.Empty, ThirdCellProperties = string.Empty,
                FourthCellProperties = string.Empty, FifthCellProperties = string.Empty, EighthCellProperties = string.Empty;
            if (tables.Count > 0)
            {
                /// For Export Log
                var TR1 = table.Descendants<TableRow>().Skip(1).FirstOrDefault();

                var rowProperties = TR1.Descendants<TableRowProperties>().FirstOrDefault().OuterXml;

                if (TR1.Descendants<TableCell>().ToList().Count > 0)
                {
                    var FirstCell = TR1.Descendants<TableCell>().FirstOrDefault();
                    CellProperties = FirstCell.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;

                    var SecondCell = TR1.Descendants<TableCell>().Skip(1).FirstOrDefault();
                    SecondCellProperties = SecondCell.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;

                    if (TR1.Descendants<TableCell>().ToList().Count > 0)
                    {
                        var FirstCellD = TR1.Descendants<TableCell>().FirstOrDefault();
                        CellProperties = FirstCellD.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;

                        var SecondCellD = TR1.Descendants<TableCell>().Skip(1).FirstOrDefault();
                        SecondCellProperties = SecondCellD.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;

                        var ThirdCellD = TR1.Descendants<TableCell>().Skip(2).FirstOrDefault();
                        ThirdCellProperties = ThirdCellD.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;

                        var FourthCellD = TR1.Descendants<TableCell>().Skip(3).FirstOrDefault();
                        FourthCellProperties = FourthCellD.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;

                        var FifthCellD = TR1.Descendants<TableCell>().Skip(4).FirstOrDefault();
                        FifthCellProperties = FifthCellD.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;
                    }
                }

                //var tABList = TR1.Descendants<Table>().ToList();
                var whiteTable = TR1.Descendants<Table>().FirstOrDefault();
                var blackTable = TR1.Descendants<Table>().Skip(1).FirstOrDefault();

                var whiteTableCell = TR1.Descendants<TableCell>().Skip(4).FirstOrDefault();

                var whiteCellProperties = whiteTableCell.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;
                var whiteCellPara = whiteTableCell.Descendants<Paragraph>().FirstOrDefault();

                TR1.RemoveAllChildren();
                TR1.Remove();


                Table currentTable = table;
                var tempDoc = "";
                var tempDocLoc = "";
                int count = 0;
                int docTotalAppCount = 0;
                decimal totalcharges = 0;
                foreach (var record in _IMEDataReport2)
                {
                    if (tempDoc.Trim() != record.Doctor.Trim())
                    {
                        count++;
                        Table tableHeader = new Table();

                        // Create a TableProperties object and specify its border information.
                        TableProperties tblProp = new TableProperties(
                            new TableBorders(
                                new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                                new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                                new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                                new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 }
                            ),
                            new TableCellBorders(
                                new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                                new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                                new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                                new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 }
                            )
                        //new TableRow(
                        //    new TableRowHeight() { Val = Convert.ToUInt32("20") }
                        //)
                        );
                        Paragraph paragraph1 = new Paragraph();
                        Paragraph paragraph2 = new Paragraph();
                        ParagraphProperties paraProperties = new ParagraphProperties();
                        Justification justification = new Justification() { Val = JustificationValues.Left };
                        RunProperties runProperties = new RunProperties();
                        FontSize fontSize = new FontSize();
                        fontSize.Val = "28";
                        Bold bold1 = new Bold();
                        bold1.Val = true;
                        runProperties.Append(new RunFonts() { Ascii = "Times New Roman" });
                        runProperties.Append(fontSize);
                        runProperties.Append(bold1);
                        paraProperties.Append(justification);

                        // Append the TableProperties object to the empty table.
                        tableHeader.AppendChild<TableProperties>(tblProp);

                        // Now we create a new layout and make it "fixed".
                        TableLayout tl = new TableLayout() { Type = TableLayoutValues.Autofit };
                        tblProp.TableLayout = tl;

                        TableRow trHeader = new TableRow();

                        // Create a cell.
                        TableCell tcHeader1 = new TableCell();

                        //TableCellVerticalAlignment tcVA = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                        // Specify the width property of the table cell.
                        tcHeader1.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "9600" }), new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }, new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 8 });
                        //create Run and Text 
                        Run run1 = new Run();
                        Text text1 = new Text();

                        //add content in Text
                        text1.Text = record.Doctor.Trim() ?? " ";

                        //add Text to Run
                        run1.Append(runProperties);
                        run1.Append(text1);

                        //add Run to paragraph
                        paragraph1.Append(paraProperties);
                        paragraph1.Append(run1);

                        // Specify the table cell content.
                        tcHeader1.Append(paragraph1);

                        TableRowProperties tableRowProperties1 = new TableRowProperties();
                        TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = (UInt32Value)558U };

                        tableRowProperties1.Append(tableRowHeight1);
                        trHeader.Append(tableRowProperties1);

                        trHeader.Append(tcHeader1);
                        tableHeader.Append(trHeader);

                        if (count >= 1)
                        {
                            table = (Table)doc.MainDocumentPart.Document.Body.InsertAfter(template.CloneNode(true), currentTable);
                            int countRow = 1;
                            int rowss = table.Elements<TableRow>().Count();
                            var temp = table.Elements<TableRow>();
                            foreach (var row in table.Elements<TableRow>())
                            {
                                if (countRow >= 2)
                                {
                                    table.RemoveChild<TableRow>(row);
                                }
                                countRow++;
                            }
                        }
                        doc.MainDocumentPart.Document.Body.InsertBefore(tableHeader, table);
                        if (count >= 2)
                        {
                            var charge = "$ " + totalcharges.ToString("F", CultureInfo.InvariantCulture);

                            TableRow trroot = new TableRow();
                            TableCell tcroot1 = new TableCell();
                            tcroot1.Append(new TableCellProperties(CellProperties));
                            TableCell tcroot2 = new TableCell();
                            tcroot2.Append(new TableCellProperties(CellProperties));
                            TableCell tcroot3 = new TableCell();
                            tcroot3.Append(new TableCellProperties(CellProperties));
                            TableCell tcroot4 = new TableCell();
                            tcroot4.Append(new TableCellProperties(CellProperties));
                            TableCell tcroot5 = new TableCell();
                            tcroot5.Append(new TableCellProperties(CellProperties));
                            tcroot5.Append(new Paragraph(new Run(new Text("TotalCharges:   " + charge))));
                            trroot.Append(tcroot1, tcroot2, tcroot3, tcroot4, tcroot5);
                            currentTable.Append(trroot);
                            //doc.MainDocumentPart.Document.Body.InsertBefore(paraCount, tableHeader);
                            docTotalAppCount = 0;

                            var tempHeaderTable = (Table)doc.MainDocumentPart.Document.Body.InsertBefore(HeaderTable.CloneNode(true), tableHeader);

                            Paragraph para = new Paragraph();
                            Run run = new Run();
                            //add Text to Run
                            run.Append(new Break() { Type = BreakValues.Page });
                            //add Run to paragraph
                            para.Append(run);
                            doc.MainDocumentPart.Document.Body.InsertBefore(para, tempHeaderTable);


                            Paragraph paraHeader = new Paragraph();
                            Run runHeader = new Run();
                            //add Text to Run
                            runHeader.Append(new Text(""));
                            //add Run to paragraph
                            paraHeader.Append(runHeader);

                            doc.MainDocumentPart.Document.Body.InsertAfter(paraHeader, tempHeaderTable);

                            totalcharges = 0;
                        }
                    }

                    currentTable = table;
                    docTotalAppCount++;
                    tempDoc = record.Doctor;
                    tempDocLoc = record.DoctorLocation;

                    TableRow tr = new TableRow();

                    // Create a cell.
                    TableCell tc1 = new TableCell();
                    TableCell tc2 = new TableCell();
                    TableCell tc3 = new TableCell();
                    TableCell tc4 = new TableCell();
                    TableCell tc5 = new TableCell();

                    // Specify the width property of the table cell.
                    tc1.Append(new TableCellProperties(CellProperties));
                    tc2.Append(new TableCellProperties(SecondCellProperties));
                    tc3.Append(new TableCellProperties(ThirdCellProperties));
                    tc4.Append(new TableCellProperties(FourthCellProperties));
                    tc5.Append(new TableCellProperties(whiteCellProperties));

                    // Specify the table cell content.
                    tc1.Append(new Paragraph(new Run(new Text(record.ClaimantName ?? " "))));
                    tc2.Append(new Paragraph(new Run(new Text(record.dateofexam ?? " "))));
                    tc3.Append(new Paragraph(new Run(new Text(record.Doctor ?? " "))));
                    tc4.Append(new Paragraph(new Run(new Text(record.DoctorLocation ?? " "))));
                    tc5.Append(new Paragraph(new Run(new Text(record.CaseChargeCount.ToString("F", CultureInfo.InvariantCulture) ?? " "))));

                    tr.Append(tc1);
                    tr.Append(tc2);
                    tr.Append(tc3);
                    tr.Append(tc4);
                    tr.Append(tc5);

                    currentTable.Append(tr);
                    totalcharges += record.CaseChargeCount;
                    if (_IMEDataReport2.IndexOf(record) == _IMEDataReport2.Count - 1)
                    {
                        var charge = "$ " + totalcharges.ToString("F", CultureInfo.InvariantCulture);

                        //doc.MainDocumentPart.Document.Body.Append(paraCount);
                        TableRow trroot = new TableRow();
                        TableCell tcroot1 = new TableCell();
                        tcroot1.Append(new TableCellProperties(CellProperties));
                        TableCell tcroot2 = new TableCell();
                        tcroot2.Append(new TableCellProperties(CellProperties));
                        TableCell tcroot3 = new TableCell();
                        tcroot3.Append(new TableCellProperties(CellProperties));
                        TableCell tcroot4 = new TableCell();
                        tcroot4.Append(new TableCellProperties(CellProperties));
                        TableCell tcroot5 = new TableCell();
                        tcroot5.Append(new TableCellProperties(CellProperties));
                        tcroot5.Append(new Paragraph(new Run(new Text("TotalCharges:   " + charge))));
                        trroot.Append(tcroot1, tcroot2, tcroot3, tcroot4, tcroot5);
                        currentTable.Append(trroot);
                        docTotalAppCount = 0;
                    }

                }

                foreach (var row in template.Elements<TableRow>())
                {
                    template.RemoveChild<TableRow>(row);
                }

            }

            return mergeFields;
        }
        public static IEnumerable<FieldCode> FillMergeFieldsForUserReport(this WordprocessingDocument doc, List<IMEDataReport2Updated> _IMEDataReport2, IMEReportFiltter IMEReportFiltter)
        {
            if (doc == null)
                return null;


            List<Table> tables = doc.MainDocumentPart.Document.Body.Descendants<Table>().ToList();
            Table table = new Table();
            Table template = new Table();
            Table HeaderTable = new Table();

            HeaderTable = doc.MainDocumentPart.Document.Body.Descendants<Table>().ElementAt(0);
            table = doc.MainDocumentPart.Document.Body.Descendants<Table>().ElementAt(1);

            List<FieldCode> mergeFields = doc.MainDocumentPart.RootElement.Descendants<FieldCode>().ToList();
            foreach (var header in doc.MainDocumentPart.HeaderParts)
            {
                mergeFields.AddRange(header.RootElement.Descendants<FieldCode>());
            }
            foreach (var footer in doc.MainDocumentPart.FooterParts)
            {
                mergeFields.AddRange(footer.RootElement.Descendants<FieldCode>());
            }

            Dictionary<string, string> pDictionaryMerge = new Dictionary<string, string>();

            string CellProperties = string.Empty, LastCellProperties = string.Empty, SecondCellProperties = string.Empty, ThirdCellProperties = string.Empty;
            if (tables.Count > 0)
            {
                // For Export Log
                var TR1 = table.Descendants<TableRow>().Skip(1).FirstOrDefault();

                var rowProperties = TR1.Descendants<TableRowProperties>().FirstOrDefault().OuterXml;

                if (TR1.Descendants<TableCell>().ToList().Count > 0)
                {
                    var FirstCell = TR1.Descendants<TableCell>().FirstOrDefault();
                    CellProperties = FirstCell.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;

                    var SecondCell = TR1.Descendants<TableCell>().Skip(1).FirstOrDefault();
                    SecondCellProperties = SecondCell.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;

                    if (TR1.Descendants<TableCell>().ToList().Count > 0)
                    {
                        var FirstCellD = TR1.Descendants<TableCell>().FirstOrDefault();
                        CellProperties = FirstCellD.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;

                        var SecondCellD = TR1.Descendants<TableCell>().Skip(1).FirstOrDefault();
                        SecondCellProperties = SecondCellD.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;

                        var ThirdCellD = TR1.Descendants<TableCell>().Skip(2).FirstOrDefault();
                        ThirdCellProperties = ThirdCellD.Descendants<TableCellProperties>().FirstOrDefault().OuterXml;

                    }
                }

                //var tABList = TR1.Descendants<Table>().ToList();

                var whiteTableCell = TR1.Descendants<TableCell>().Skip(4).FirstOrDefault();

                TR1.RemoveAllChildren();
                TR1.Remove();

                //user as header
                Table tableHeader = new Table();

                // Create a TableProperties object and specify its border information.
                TableProperties tblProp = new TableProperties(
                    new TableBorders(
                        new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                        new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                        new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                        new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 }
                    ),
                    new TableCellBorders(
                        new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                        new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                        new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 },
                        new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.None), Size = 0 }
                    )
                );
                Paragraph paragraph1 = new Paragraph();
                Paragraph paragraph2 = new Paragraph();
                ParagraphProperties paraProperties = new ParagraphProperties();
                Justification justification = new Justification() { Val = JustificationValues.Left };
                RunProperties runProperties = new RunProperties();
                FontSize fontSize = new FontSize();
                fontSize.Val = "28";
                Bold bold1 = new Bold();
                bold1.Val = true;
                runProperties.Append(new RunFonts() { Ascii = "Times New Roman" });
                runProperties.Append(fontSize);
                runProperties.Append(bold1);
                paraProperties.Append(justification);

                // Append the TableProperties object to the empty table.
                tableHeader.AppendChild<TableProperties>(tblProp);

                // Now we create a new layout and make it "fixed".
                TableLayout tl = new TableLayout() { Type = TableLayoutValues.Autofit };
                tblProp.TableLayout = tl;

                TableRow trHeader = new TableRow();

                // Create a cell.
                TableCell tcHeader1 = new TableCell();

                // Specify the width property of the table cell.
                tcHeader1.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "4800" }), new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }, new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 8 });
                //create Run and Text 
                Run run1 = new Run();
                Text text1 = new Text();
                Text text2 = new Text();

                //add content in Text
                text1.Text = _IMEDataReport2[0].fullName.Trim() ?? " ";


                text2.Text = " (TotalCases: " + _IMEDataReport2[0].totalCaseCount.ToString()+")";

                //add Text to Run
                run1.Append(runProperties);
                run1.Append(text1);
                run1.Append(text2);

                //add Run to paragraph
                paragraph1.Append(paraProperties);
                paragraph1.Append(run1);

                // Specify the table cell content.
                tcHeader1.Append(paragraph1);

                TableRowProperties tableRowProperties1 = new TableRowProperties();
                TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = (UInt32Value)558U };

                tableRowProperties1.Append(tableRowHeight1);
                trHeader.Append(tableRowProperties1);

                trHeader.Append(tcHeader1);
                tableHeader.Append(trHeader);

                doc.MainDocumentPart.Document.Body.InsertBefore(tableHeader, table);

                foreach (var record in _IMEDataReport2)
                {
                    TableRow tr = new TableRow();

                    // Create a cell.
                    TableCell tc1 = new TableCell();
                    TableCell tc2 = new TableCell();
                    TableCell tc3 = new TableCell();

                    // Specify the width property of the table cell.
                    tc1.Append(new TableCellProperties(CellProperties));
                    tc2.Append(new TableCellProperties(SecondCellProperties));
                    tc3.Append(new TableCellProperties(ThirdCellProperties));

                    // Specify the table cell content.
                    tc1.Append(new Paragraph(new Run(new Text(record.casenumber ?? " "))));
                    tc2.Append(new Paragraph(new Run(new Text(record.History.Replace("<br>", "") ?? " "))));
                    tc3.Append(new Paragraph(new Run(new Text(record.HistoryDate ?? " "))));

                    tr.Append(tc1, tc2, tc3);
                    table.Append(tr);
                }
                TableRow trroot = new TableRow();
                TableCell tcroot1 = new TableCell();
                tcroot1.Append(new TableCellProperties(CellProperties));
                TableCell tcroot2 = new TableCell();
                tcroot2.Append(new TableCellProperties(CellProperties));
                TableCell tcroot3 = new TableCell();
                tcroot3.Append(new TableCellProperties(CellProperties));
                tcroot3.Append(new Paragraph(new Run(new Text("TotalWorkDescription:   " + (_IMEDataReport2.Count - _IMEDataReport2[0].withoutstatus)))));
                trroot.Append(tcroot1, tcroot2, tcroot3);
                table.Append(trroot);

                foreach (var row in template.Elements<TableRow>())
                {
                    template.RemoveChild<TableRow>(row);
                }
            }
            return mergeFields;
        }

        /// <summary>
        /// Gets merge fields contained in the given element.
        /// </summary>
        /// <param name="mergeFieldName">Optional name for the merge fields to look for..</param>
        /// <returns>If a merge field name is specified, only merge fields with that name are returned. Otherwise, it returns all merge fields contained in the given element.</returns>
        public static IEnumerable<FieldCode> GetMergeFields(this OpenXmlElement xmlElement, string mergeFieldName = null)
        {
            if (xmlElement == null)
                return null;

            if (string.IsNullOrWhiteSpace(mergeFieldName))
                return xmlElement
                    .Descendants<FieldCode>();

            return xmlElement
                .Descendants<FieldCode>()
                .Where(f => f.InnerText.StartsWith(GetMergeFieldStartString(mergeFieldName)));
        }

        /// <summary>
        /// Filters merge fields by the given name.
        /// </summary>
        /// <param name="mergeFieldName">The merge field name.</param>
        /// <returns>Returns all merge fields with the given name. If the merge field name is null or blank, it returns nothing.</returns>
        public static IEnumerable<FieldCode> WhereNameIs(this IEnumerable<FieldCode> mergeFields, string mergeFieldName)
        {
            if (mergeFields == null || mergeFields.Count() == 0)
                return null;

            return mergeFields
                .Where(f => f.InnerText.StartsWith(GetMergeFieldStartString(mergeFieldName)));
        }

        /// <summary>
        /// Gets the immediate containing paragraph of a given element.
        /// </summary>
        /// <returns>If the given element is a paragraph, that element is returned. Otherwise, it returns the immediate ancestor that is a paragraph, or null if none is found.</returns>
        public static Paragraph GetParagraph(this OpenXmlElement xmlElement)
        {
            if (xmlElement == null)
                return null;

            Paragraph paragraph = null;
            if (xmlElement is Paragraph)
                paragraph = (Paragraph)xmlElement;
            else if (xmlElement.Parent is Paragraph)
                paragraph = (Paragraph)xmlElement.Parent;
            else
                paragraph = xmlElement.Ancestors<Paragraph>().FirstOrDefault();
            return paragraph;
        }

        /// <summary>
        /// Removes a merge field from the containing document and replaces it with the given text content. 
        /// </summary>
        /// <param name="replacementText">The content to replace the merge field with.</param>
        public static void ReplaceWithText(this FieldCode field, string replacementText, NYCTemplate nYCTemplateObj)
        {
            if (field == null)
                return;

            Run rFldCode = field.Parent as Run;
            Run rBegin = rFldCode.PreviousSibling<Run>();
            Run rSep = rFldCode.NextSibling<Run>();
            Run rText = rSep.NextSibling<Run>();
            Run rEnd = rText.NextSibling<Run>();

            //try
            //{
            //    Run rEnd1 = rEnd.NextSibling<Run>();
            //    rEnd.Remove();
            //}
            //catch (Exception e)
            //{

            //}
            //var propCopied = rText.PreviousSibling<RunProperties>().FirstOrDefault();

            bool IsSplitDone = false;
            bool NeedToDelete = false;


            var property = rText.GetParagraph().Descendants<RunProperties>();

            var PROP = "";
            foreach (RunProperties placeholderrpr in rText.GetParagraph().Descendants<RunProperties>())
            {
                PROP = placeholderrpr.OuterXml;
                if (!string.IsNullOrEmpty(PROP))
                    break;
            }

            var prop = property.FirstOrDefault();
            string fontSize = null;
            string fontAscii = null;
            if (prop != null)
            {
                fontSize = prop.FontSize?.Val ?? nYCTemplateObj.fontSize;
                fontAscii = prop.RunFonts?.Ascii ?? "Arial";
                if (prop.FontSize != null && prop.FontSize?.Val != null)
                    nYCTemplateObj.fontSize = prop.FontSize?.Val;
            }
            /// For break line containg ~
            //if (replacementText.Contains("~"))
            //{
            //    IsSplitDone = true;
            //    string[] AddressPart = replacementText.Split('~');
            //    foreach (string line in AddressPart.Skip(1))
            //    {
            //        if (!string.IsNullOrEmpty(line))
            //        {
            //            var r = rFldCode.GetParagraph().AppendChild(new Run());
            //            //r.Descendants<fldChar>
            //            //var fldChar = rFldCode.GetParagraph().Descendants<FieldChar>();
            //            var fldChar1 = rFldCode.GetParagraph().Descendants<FieldChar>();

            //            //rFldCode.Append(new Text(line.Trim()));
            //            var count1 = fldChar1.Count();

            //            r.AppendChild(new Text(line.Trim()));
            //            r.AppendChild(new Break());

            //            /// Font and Size
            //            RunProperties runProp = new RunProperties();

            //            var runFont = new RunFonts { Ascii = fontAscii ?? "Arial" };
            //            // 48 half-point font size
            //            var size = prop.FontSize?.Val;
            //            var sizeprop = new FontSize { Val = new StringValue(fontSize ?? "21") };
            //            runProp.Append(runFont);
            //            runProp.Append(sizeprop);
            //            if (String.IsNullOrEmpty(PROP))
            //            {
            //                r.PrependChild(new RunProperties(PROP));
            //            }
            //            else
            //            {
            //                r.PrependChild(runProp);
            //            }
            //        }
            //    }
            //}
            /// For HealthCarePractioners
            foreach (var a in rText.Descendants<Text>())
            {
                char[] charsToTrim = { '«', '»' };
                string cleanString = a.Text.Trim(charsToTrim);

                if (cleanString == "Provider.FullName")
                {
                    NeedToDelete = true;
                    var count = 0;
                    foreach (var practioner in nYCTemplateObj.HealthCarePractioners)
                    {
                        var r = rFldCode.GetParagraph().AppendChild(new Run());
                        if (count != 0)
                        {
                            r.AppendChild(new Break());
                            if (practioner.ServiceProviderAddress != null)
                            {
                                r.AppendChild(new Break());
                            }
                        }
                        count++;
                        ///Appending name
                        r.AppendChild(new Text(practioner.ServiceProviderName));
                        if (practioner.ServiceProviderAddress != null)
                        {
                            if (practioner.ServiceProviderAddress.Contains('~'))
                            {
                                string[] AddressPartInPractioner = practioner.ServiceProviderAddress.Split('~');
                                foreach (string line in AddressPartInPractioner)
                                {
                                    ///Appending address with multiple lines
                                    if (!string.IsNullOrEmpty(line))
                                    {
                                        r.AppendChild(new Break());
                                        r.AppendChild(new Text(line));
                                    }
                                }
                            }
                            else
                            {
                                ///Appending address 
                                r.AppendChild(new Break());
                                r.AppendChild(new Text(practioner.ServiceProviderAddress));
                            }
                            //r.AppendChild(new Text(practioner.ServiceProviderAddress));
                        }
                        RunProperties runProp = new RunProperties();
                        var runFont = new RunFonts { Ascii = "Arial" };
                        // 48 half-point font size
                        var size = new FontSize { Val = new StringValue("24") };
                        runProp.Append(runFont);
                        runProp.Append(size);

                        if (String.IsNullOrEmpty(PROP))
                        {
                            r.PrependChild(new RunProperties(PROP));
                        }
                        else
                        {
                            r.PrependChild(runProp);
                        }
                    }
                }
            }

            Text t = rText.GetFirstChild<Text>();
            if (t != null)
            {
                t.Text = (replacementText != null) ? replacementText : string.Empty;
            }
            rFldCode.Remove();
            rBegin.Remove();
            rSep.Remove();
            rEnd.Remove();

            ///Removing line Space of provider name in TableStart
            if (NeedToDelete)
            {
                var FieldCodeNearToProvider = rFldCode.Descendants<FieldCode>();
                var f = FieldCodeNearToProvider.FirstOrDefault();
                if (f != null)
                    f.Remove();
            }
        }

        /// <summary>
        /// Removes the merge fields from the containing document and replaces them with the given text content. 
        /// </summary>
        /// <param name="replacementText">The content to replace the merge field with.</param>
        public static void ReplaceWithText(this IEnumerable<FieldCode> fields, string replacementText, NYCTemplate nYCTemplateObj)
        {
            if (fields == null || fields.Count() == 0)
            {
                return;
            }
            foreach (var field in fields)
            {
                field.ReplaceWithText(replacementText, nYCTemplateObj);
            }
        }

        /// <summary>
        /// Removes the merge fields from the containing document and replaces them with the given texts. 
        /// </summary>
        /// <param name="replacementTexts">The text values to replace the merge fields with</param>
        /// <param name="removeExcess">Optional value to indicate that excess merge fields are removes instead of replacing with blank values</param>
        public static void ReplaceWithText(this IEnumerable<FieldCode> fields, IEnumerable<string> replacementTexts, NYCTemplate nYCTemplateObj, bool removeExcess = false)
        {
            if (fields == null || fields.Count() == 0)
                return;

            int replacementCount = replacementTexts.Count();
            int index = 0;
            foreach (var field in fields)
            {
                if (index < replacementCount)
                    field.ReplaceWithText(replacementTexts.ElementAt(index), nYCTemplateObj);
                else if (removeExcess)
                    field.GetParagraph().Remove();
                else
                    field.ReplaceWithText(string.Empty, nYCTemplateObj);

                index++;
            }
        }

        #region Private Methods

        private static string GetMergeFieldStartString(string mergeFieldName)
        {
            return " MERGEFIELD  " + (!string.IsNullOrWhiteSpace(mergeFieldName) ? mergeFieldName : "<NoNameMergeField>");
        }

        #endregion Private Methods
        public static IEnumerable<SimpleField> GetSimpleFieldMergeFields(this WordprocessingDocument doc, string mergeFieldName = null)
        {
            if (doc == null)
                return null;
            List<SimpleField> SimpleFields = doc.MainDocumentPart.RootElement.Descendants<SimpleField>().ToList();

            foreach (var header in doc.MainDocumentPart.HeaderParts)
            {
                SimpleFields.AddRange(header.RootElement.Descendants<SimpleField>());
            }

            foreach (var footer in doc.MainDocumentPart.FooterParts)
            {
                SimpleFields.AddRange(footer.RootElement.Descendants<SimpleField>());
            }

            if (!string.IsNullOrWhiteSpace(mergeFieldName) && SimpleFields != null && SimpleFields.Count() > 0)
                return SimpleFields.WhereNameIs(mergeFieldName);

            return SimpleFields;
        }

        public static IEnumerable<SimpleField> WhereNameIs(this IEnumerable<SimpleField> SimpleFields, string mergeFieldName)
        {
            if (SimpleFields == null || SimpleFields.Count() == 0)
                return null;

            var a = SimpleFields
                .Where(f => f.InnerText.Contains(mergeFieldName));
            return a;
        }
        public static void ReplaceWithText(this IEnumerable<SimpleField> fields, string replacementText, NYCTemplate nYCTemplateObj)
        {
            foreach (var field in fields)
            {

                if (field != null)
                {
                    var textChildren = field.Descendants<Text>();

                    textChildren.First().Text = replacementText;
                    foreach (var others in textChildren.Skip(1))
                    {
                        others.Remove();
                    }
                }
            }
        }
        public static void ReplaceWithText(this SimpleField field, string replacementText, NYCTemplate nYCTemplateObj)
        {
            var sf = field as SimpleField;
            string fontSize = null;
            string fontAscii = null;
            if (sf != null)
            {
                var property = sf.GetParagraph().Descendants<RunProperties>();
                var nm = sf.Descendants<Run>();
                var prop = property.FirstOrDefault();
                if (prop != null)
                {
                    fontSize = prop.FontSize?.Val ?? "15";
                    fontAscii = prop.RunFonts?.Ascii ?? "Arial";
                    if (prop.FontSize != null && prop.FontSize?.Val != null)
                        nYCTemplateObj.fontSize = prop.FontSize?.Val;
                }
                //var p = sf.Parent;
                //p.ChildElements
                //var r= sf.GetParagraph().AppendChild(new Run());
                bool IsSplitDone = false;
                //if (replacementText.Contains("~"))
                //{
                //    IsSplitDone = true;
                //    string[] AddressPart = replacementText.Split('~');
                //    var fldChar1 = sf.GetParagraph().Descendants<FieldChar>();
                //    var count1 = fldChar1.Count();
                //    if (prop.FontSize == null)
                //    {
                //        fontSize = nYCTemplateObj.fontSize;
                //        foreach (string line in AddressPart.Skip(1))
                //        {
                //            if (!string.IsNullOrEmpty(line))
                //            {
                //                //var nRun = new Run();
                //                //sf.GetParagraph().ReplaceChild<Run>(nRun, nRun);
                //                var r = sf.GetParagraph().AppendChild(new Run());
                //                if (nYCTemplateObj.Currentfield.Contains("Claimant.FullAddress"))
                //                {
                //                    r.AppendChild(new Text(line));
                //                    r.AppendChild(new Break());
                //                }
                //                else
                //                {
                //                    r.AppendChild(new Break());
                //                    r.AppendChild(new Text(line));
                //                }
                //                RunProperties runProp = new RunProperties();
                //                var runFont = new RunFonts { Ascii = fontAscii ?? "Arial" };
                //                var size = new FontSize { Val = new StringValue(fontSize ?? "21") };
                //                runProp.Append(runFont);
                //                runProp.Append(size);
                //                r.PrependChild(runProp);
                //                //replacementText = ;
                //            }
                //        }
                //        //r.AppendChild(new Break());
                //    }
                //    else
                //    {
                //        foreach (string line in AddressPart.Skip(1))
                //        {
                //            if (!string.IsNullOrEmpty(line))
                //            {
                //                //var nRun = new Run();
                //                //sf.GetParagraph().ReplaceChild<Run>(nRun, nRun);
                //                var r = sf.GetParagraph().AppendChild(new Run());
                //                r.AppendChild(new Break());
                //                r.AppendChild(new Text(line));
                //                RunProperties runProp = new RunProperties();
                //                var runFont = new RunFonts { Ascii = fontAscii ?? "Arial" };
                //                // 48 half-point font size
                //                var size = new FontSize { Val = new StringValue(fontSize ?? "54") };
                //                runProp.Append(runFont);
                //                runProp.Append(size);
                //                r.PrependChild(runProp);
                //                //replacementText = "";
                //            }
                //        }
                //    }
                //}
                if (IsSplitDone)
                {
                    string[] AddressPart = replacementText.Split('~');
                    replacementText = AddressPart.FirstOrDefault() ?? " ";
                }
                var Isremoved = false;
                var loopBreak = false;
                /// For HealthCarePractioners
                foreach (var a in sf.Descendants<Text>())
                {
                    if (!loopBreak)
                    {
                        char[] charsToTrim = { '«', '»' };
                        string cleanString = a.Text.Trim(charsToTrim);
                        ///Removing line Space of TableEnd
                        if (a.Text.Contains("TableEnd"))
                        {
                            var Para = sf.GetParagraph();
                            Para.RemoveAllChildren();
                            Para.Remove();
                            Isremoved = true;
                            loopBreak = true;
                            break;
                        }
                        else if (cleanString == "Provider.FullName")
                        {
                            replacementText = "";
                            var PROP = "";
                            try
                            {
                                foreach (RunProperties placeholderrpr in sf.Descendants<RunProperties>().ToList())
                                {
                                    PROP = placeholderrpr.OuterXml;
                                    break;
                                }
                            }
                            catch (Exception e) { }
                            RunProperties customProp = new RunProperties();
                            var runFont = new RunFonts { Ascii = "Arial" };
                            var size = new FontSize { Val = new StringValue("24") };
                            customProp.Append(runFont);
                            customProp.Append(size);

                            var count = 0;
                            loopBreak = true;
                            if (nYCTemplateObj.HealthCarePractioners != null)
                            {
                                foreach (var practioner in nYCTemplateObj.HealthCarePractioners)
                                {

                                    /// Adding New Run
                                    var r = sf.GetParagraph().AppendChild(new Run());


                                    /// Adding font properties
                                    try
                                    {
                                        /// Inherited props
                                        r.PrependChild(new RunProperties(PROP));
                                    }
                                    catch (Exception e)
                                    {
                                        /// Custom props
                                        r.PrependChild(customProp);
                                    }
                                    if (count != 0)
                                    {
                                        /// Adding breaks not in first line
                                        r.AppendChild(new Break());
                                        if (practioner.ServiceProviderAddress != null)
                                        {
                                            r.AppendChild(new Break());
                                        }
                                    }
                                    count++;

                                    /// Adding Provider name
                                    r.AppendChild(new Text(practioner.ServiceProviderName));

                                    /// Adding provider Address if not null
                                    if (practioner.ServiceProviderAddress != null)
                                    {
                                        /// if Address contains '~' 
                                        if (practioner.ServiceProviderAddress.Contains('~'))
                                        {
                                            string[] AddressPartInPractioner = practioner.ServiceProviderAddress.Split('~');
                                            foreach (string line in AddressPartInPractioner)
                                            {
                                                if (!string.IsNullOrEmpty(line))
                                                {
                                                    r.AppendChild(new Break());
                                                    r.AppendChild(new Text(line));
                                                }
                                            }
                                        }
                                        else
                                        {
                                            /// Address in one line
                                            r.AppendChild(new Break());
                                            r.AppendChild(new Text(practioner.ServiceProviderAddress));
                                        }
                                    }

                                }
                            }
                            if (loopBreak)
                                break;
                        }
                    }
                }

                if (!Isremoved)
                {
                    var textChildren = sf.Descendants<Text>();
                    textChildren.First().Text = replacementText;

                    foreach (var others in textChildren.Skip(1))
                    {
                        others.Remove();
                    }
                }
            }
        }

        private static bool ContainsCharType(this Run run, FieldCharValues fieldCharType)
        {
            var fc = run.GetFirstChild<FieldChar>();
            return fc == null
                ? false
                : fc.FieldCharType.Value == fieldCharType;
        }

        public static void PrependHeader(string headerTemplatePath, string documentPath)
        {
            // Open docx where we need to add header
            using (var wdDoc = WordprocessingDocument.Open(documentPath, true))
            {
                var mainPart = wdDoc.MainDocumentPart;

                // Delete exist header
                mainPart.DeleteParts(mainPart.HeaderParts);

                // Create new header
                var headerPart = mainPart.AddNewPart<HeaderPart>();

                // Get id of new header
                var rId = mainPart.GetIdOfPart(headerPart);

                // Open the header document to be copied
                using (var wdDocSource = WordprocessingDocument.Open(headerTemplatePath, true))
                {
                    // Get header part
                    var firstHeader = wdDocSource.MainDocumentPart.HeaderParts.FirstOrDefault();
                    if (firstHeader != null)
                    {
                        // Copy content of header to new header
                        headerPart.FeedData(firstHeader.GetStream());
                        // Keep formatting
                        foreach (var childElement in headerPart.Header.Descendants<Paragraph>())
                        {
                            var paragraph = (Paragraph)childElement;
                            if (paragraph.ParagraphProperties.SpacingBetweenLines == null)
                            {
                                paragraph.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines
                                {
                                    After = "0"
                                };
                                paragraph.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId
                                {
                                    Val = "No Spacing"
                                };
                            }
                        }
                        // Get all ids of every 'Parts'
                        var listToAdd = new List<KeyValuePair<Type, Stream>>();
                        foreach (var idPart in firstHeader.Parts)
                        {
                            var part = firstHeader.GetPartById(idPart.RelationshipId);
                            if (part is ImagePart)
                            {
                                headerPart.AddNewPart<ImagePart>("image/png", idPart.RelationshipId);
                                listToAdd.Add(new KeyValuePair<Type, Stream>(typeof(ImagePart), part.GetStream()));
                            }
                            else if (part is DiagramStylePart)
                            {
                                headerPart.AddNewPart<DiagramStylePart>(idPart.RelationshipId);
                                listToAdd.Add(new KeyValuePair<Type, Stream>(typeof(DiagramStylePart), part.GetStream()));
                            }
                            else if (part is DiagramColorsPart)
                            {
                                headerPart.AddNewPart<DiagramColorsPart>(idPart.RelationshipId);
                                listToAdd.Add(new KeyValuePair<Type, Stream>(typeof(DiagramColorsPart),
                                    part.GetStream()));
                            }
                            else if (part is DiagramDataPart)
                            {
                                headerPart.AddNewPart<DiagramDataPart>(idPart.RelationshipId);
                                listToAdd.Add(new KeyValuePair<Type, Stream>(typeof(DiagramDataPart), part.GetStream()));
                            }
                            else if (part is DiagramLayoutDefinitionPart)
                            {
                                headerPart.AddNewPart<DiagramStylePart>(idPart.RelationshipId);
                                listToAdd.Add(new KeyValuePair<Type, Stream>(typeof(DiagramStylePart), part.GetStream()));
                            }
                            else if (part is DiagramPersistLayoutPart)
                            {
                                headerPart.AddNewPart<DiagramPersistLayoutPart>(idPart.RelationshipId);
                                listToAdd.Add(new KeyValuePair<Type, Stream>(typeof(DiagramPersistLayoutPart),
                                    part.GetStream()));
                            }
                        }
                        // Foreach Part, copy stream to new header
                        var i = 0;
                        foreach (var idPart in headerPart.Parts)
                        {
                            var part = headerPart.GetPartById(idPart.RelationshipId);
                            if (part.GetType() == listToAdd[i].Key)
                            {
                                part.FeedData(listToAdd[i].Value);
                            }
                            i++;
                        }
                    }
                    else
                    {
                        mainPart.DeleteParts(mainPart.HeaderParts);
                        var sectToRemovePrs = mainPart.Document.Body.Descendants<SectionProperties>();
                        foreach (var sectPr in sectToRemovePrs)
                        {
                            // Delete reference of old header
                            sectPr.RemoveAllChildren<HeaderReference>();
                        }
                        return;
                    }
                }

                // Get all sections, and past header to each section (Page)
                var sectPrs = mainPart.Document.Body.Descendants<SectionProperties>();
                foreach (var sectPr in sectPrs)
                {
                    // Remove old header reference 
                    sectPr.RemoveAllChildren<HeaderReference>();
                    // Add new header reference
                    sectPr.PrependChild(new HeaderReference { Id = rId });
                }
            }
        }

        public static void ChangeHeader(string documentPath, IME4HeaderData iME4HeaderData)
        {
            // Replace header in target document with header of source document.
            using (WordprocessingDocument document = WordprocessingDocument.Open(documentPath, true))
            {
                //Get the main document part
                MainDocumentPart mainDocumentPart = document.MainDocumentPart;
                //foreach (var headerPar in document.MainDocumentPart.HeaderParts)
                //{
                //    //Gets the text in headers
                //    foreach (var currentText in headerPar.RootElement.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                //    {
                //        currentText.Text = currentText.Text.Replace(currentText.ToString(), "Thanks");
                //    }
                //}
                // Delete the existing header
                mainDocumentPart.DeleteParts(mainDocumentPart.HeaderParts);

                // Create a new header
                HeaderPart headerPart = mainDocumentPart.AddNewPart<HeaderPart>();

                // Get Id of the headerPart
                string headerPartId = mainDocumentPart.GetIdOfPart(headerPart);

                GenerateHeaderPartContent(headerPart, iME4HeaderData);

                //Get SectionProperties and Replace HeaderReference
                IEnumerable<SectionProperties> sections = mainDocumentPart.Document.Body.Elements<SectionProperties>();

                foreach (var section in sections)
                {
                    // Delete existing references to headers
                    section.RemoveAllChildren<HeaderReference>();

                    // Create the new header reference node
                    section.PrependChild<HeaderReference>(new HeaderReference() { Id = headerPartId });
                }
                mainDocumentPart.Document.Save();
                document.Close();
            }
        }
        public static void ChangeImage(string documentPath, string fileName)
        {
            // Replace header in target document with header of source document.
            using (WordprocessingDocument document = WordprocessingDocument.Open(documentPath, true))
            {
                //Get the main document part
                MainDocumentPart mainDocumentPart = document.MainDocumentPart;

                ImagePart imagePart = mainDocumentPart.AddImagePart(ImagePartType.Png);
                using (FileStream stream = new FileStream(fileName, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }
                if (mainDocumentPart.ImageParts.ToList().Count() < 2)
                {
                    AddImage(document, mainDocumentPart.GetIdOfPart(imagePart), fileName);
                }

                mainDocumentPart.Document.Save();
                document.Close();
            }
        }
        public static void ChangeImageHeader(string documentPath, string fileName)
        {
            // Replace header in target document with header of source document.
            using (WordprocessingDocument document = WordprocessingDocument.Open(documentPath, true))
            {
                //Get the main document part
                MainDocumentPart mainDocumentPart = document.MainDocumentPart;

                List<string> images = ExtractImages(document.MainDocumentPart.Document.Body, mainDocumentPart);

                mainDocumentPart.DeleteParts(mainDocumentPart.HeaderParts);

                // Create a new header
                HeaderPart headerPart = mainDocumentPart.AddNewPart<HeaderPart>();

                // Get Id of the headerPart
                string headerPartId = mainDocumentPart.GetIdOfPart(headerPart);
                ImagePart imagePart = headerPart.AddImagePart(ImagePartType.Png);
                using (FileStream stream = new FileStream(fileName, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }

                AddImageToHeader(headerPart, headerPart.GetIdOfPart(imagePart), fileName);

                //Get SectionProperties and Replace HeaderReference
                IEnumerable<SectionProperties> sections = mainDocumentPart.Document.Body.Elements<SectionProperties>();

                foreach (var section in sections)
                {
                    // Delete existing references to headers
                    section.RemoveAllChildren<HeaderReference>();

                    // Create the new header reference node
                    section.PrependChild<HeaderReference>(new HeaderReference() { Id = headerPartId, Type = HeaderFooterValues.First });
                }
                mainDocumentPart.Document.Save();
                document.Close();
            }
        }

        private static List<string> ExtractImages(Body content, MainDocumentPart wDoc)

        {

            List<string> imageList = new List<string>();

            foreach (Paragraph par in content.Descendants<Paragraph>())

            {

                ParagraphProperties paragraphProperties = par.ParagraphProperties;

                foreach (Run run in par.Descendants<Run>())

                {

                    //detect if the run contains an image and upload it to wordpress

                    DocumentFormat.OpenXml.Wordprocessing.Drawing image =
                    run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().FirstOrDefault();

                    if (image != null)

                    {

                        var imageFirst = image.Inline.Graphic.GraphicData.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().FirstOrDefault();

                        var blip = imageFirst.BlipFill.Blip.Embed.Value;

                        string folder = System.IO.Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);

                        ImagePart img = (ImagePart)wDoc.Document.MainDocumentPart.GetPartById(blip);

                        string imageFileName = string.Empty;

                        //the image is stored in a zip file code behind, so it must be extracted

                        using (System.Drawing.Image toSaveImage = Bitmap.FromStream(img.GetStream()))

                        {

                            imageFileName = folder + @"\TestExtractor_" + DateTime.UtcNow.Month.ToString().Trim() +
                            DateTime.UtcNow.Day.ToString() + DateTime.UtcNow.Year.ToString() + DateTime.UtcNow.Hour.ToString() +
                            DateTime.UtcNow.Minute.ToString() +

                            DateTime.UtcNow.Second.ToString() + DateTime.UtcNow.Millisecond.ToString() + ".png";

                            try

                            {

                                toSaveImage.Save(imageFileName, ImageFormat.Png);

                            }

                            catch (Exception ex)

                            {

                                //TODO: handle image issues

                            }

                        }

                        imageList.Add(imageFileName);

                    }

                }

            }

            return imageList;

        }

        private static void AddImageToHeader(HeaderPart part, string relationshipId, string fileName)
        {
            // Define the reference of the image.
            int iWidth = 0;
            int iHeight = 0;
            using (System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(fileName))
            {
                iWidth = bmp.Width;
                iHeight = bmp.Height;
            }
            iWidth = (int)Math.Round((decimal)iWidth * 9525);
            iHeight = (int)Math.Round((decimal)iHeight * 9525);
            var element =
                    new DocumentFormat.OpenXml.Wordprocessing.Drawing(
                        new DWW.Inline(
                            new DWW.Extent() { Cx = iWidth, Cy = iHeight },
                            new DWW.EffectExtent()
                            {
                                LeftEdge = 0L,
                                TopEdge = 0L,
                                RightEdge = 0L,
                                BottomEdge = 0L
                            },
                            new DWW.DocProperties()
                            {
                                Id = (UInt32Value)1U,
                                Name = fileName
                            },
                            new DWW.NonVisualGraphicFrameDrawingProperties(
                                new AA.GraphicFrameLocks() { NoChangeAspect = true }),
                            new AA.Graphic(
                                new AA.GraphicData(
                                    new PICS.Picture(
                                        new PICS.NonVisualPictureProperties(
                                            new PICS.NonVisualDrawingProperties()
                                            {
                                                Id = (UInt32Value)0U,
                                                Name = fileName
                                            },
                                            new PICS.NonVisualPictureDrawingProperties()),
                                        new PICS.BlipFill(
                                            new AA.Blip(
                                                new AA.BlipExtensionList(
                                                    new AA.BlipExtension()
                                                    {
                                                        Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                    })
                                            )
                                            {
                                                Embed = relationshipId,
                                                CompressionState =
                                                AA.BlipCompressionValues.Print
                                            },
                                            new AA.Stretch(
                                                new AA.FillRectangle())),
                                        new PICS.ShapeProperties(
                                            new AA.Transform2D(
                                                new AA.Offset() { X = 0L, Y = 0L },
                                                new AA.Extents() { Cx = iWidth, Cy = iHeight }),
                                            new AA.PresetGeometry(
                                                new AA.AdjustValueList()
                                            )
                                            { Preset = AA.ShapeTypeValues.Rectangle }))
                                )
                                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                        )
                        {
                            DistanceFromTop = (UInt32Value)0U,
                            DistanceFromBottom = (UInt32Value)0U,
                            DistanceFromLeft = (UInt32Value)0U,
                            DistanceFromRight = (UInt32Value)0U,
                            EditId = "50D07946"
                        });

            Header header1 = new Header() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
            header1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00164C17", RsidRunAdditionDefault = "00164C17" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

            paragraphProperties1.Append(paragraphStyleId1);
            //paragraphProperties1.Indentation = center;
            Run lineBreak = new Run(new Break());
            Run run = new Run();
            run.Append(element);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run);

            header1.Append(paragraph1);

            part.Header = header1;
        }
        private static void AddImage(WordprocessingDocument wordDoc, string relationshipId, string fileName)
        {
            // Define the reference of the image.
            int iWidth = 0;
            int iHeight = 0;
            using (System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(fileName))
            {
                iWidth = bmp.Width;
                iHeight = bmp.Height;
            }
            iWidth = (int)Math.Round((decimal)iWidth * 9525);
            iHeight = (int)Math.Round((decimal)iHeight * 9525);
            var element =
                    new DocumentFormat.OpenXml.Wordprocessing.Drawing(
                        new DWW.Inline(
                            new DWW.Extent() { Cx = iWidth, Cy = iHeight - 10 },
                            new DWW.EffectExtent()
                            {
                                LeftEdge = 0L,
                                TopEdge = 0L,
                                RightEdge = 0L,
                                BottomEdge = 0L
                            },
                            new DWW.DocProperties()
                            {
                                Id = (UInt32Value)1U,
                                Name = fileName
                            },
                            new DWW.NonVisualGraphicFrameDrawingProperties(
                                new AA.GraphicFrameLocks() { NoChangeAspect = true }),
                            new AA.Graphic(
                                new AA.GraphicData(
                                    new PICS.Picture(
                                        new PICS.NonVisualPictureProperties(
                                            new PICS.NonVisualDrawingProperties()
                                            {
                                                Id = (UInt32Value)0U,
                                                Name = fileName
                                            },
                                            new PICS.NonVisualPictureDrawingProperties()),
                                        new PICS.BlipFill(
                                            new AA.Blip(
                                                new AA.BlipExtensionList(
                                                    new AA.BlipExtension()
                                                    {
                                                        Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                    })
                                            )
                                            {
                                                Embed = relationshipId,
                                                CompressionState =
                                                AA.BlipCompressionValues.Print
                                            },
                                            new AA.Stretch(
                                                new AA.FillRectangle())),
                                        new PICS.ShapeProperties(
                                            new AA.Transform2D(
                                                new AA.Offset() { X = 0L, Y = 0L },
                                                new AA.Extents() { Cx = iWidth, Cy = iHeight - 10 }),
                                            new AA.PresetGeometry(
                                                new AA.AdjustValueList()
                                            )
                                            { Preset = AA.ShapeTypeValues.Rectangle }))
                                )
                                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                        )
                        {
                            DistanceFromTop = (UInt32Value)0U,
                            DistanceFromBottom = (UInt32Value)0U,
                            DistanceFromLeft = (UInt32Value)0U,
                            DistanceFromRight = (UInt32Value)0U,
                            EditId = "50D07946"
                        });

            wordDoc.MainDocumentPart.Document.Body.PrependChild(new Paragraph(new Run(element)));
        }

        static void GenerateHeaderPartContent(HeaderPart part, IME4HeaderData iMEHearderData)
        {
            Header header1 = new Header() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
            header1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            //// Create an empty table.
            //Table table = new Table();

            //// Create a TableProperties object and specify its border information.
            //TableProperties tblProp = new TableProperties(
            //    new TableBorders(
            //        new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
            //        new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
            //        new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
            //        new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
            //        new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
            //        new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 }
            //    )
            //);
            //// Append the TableProperties object to the empty table.
            //table.AppendChild<TableProperties>(tblProp);

            //// Create a row.
            //TableRow tr = new TableRow();

            //// Create a cell.
            //TableCell tc1 = new TableCell();

            //// Specify the width property of the table cell.
            //tc1.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));

            //// Specify the table cell content.
            //tc1.Append(new Paragraph(new Run(new Text("Hello, World!"))));

            //// Append the table cell to the table row.
            //tr.Append(tc1);

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00164C17", RsidRunAdditionDefault = "00164C17" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

            paragraphProperties1.Append(paragraphStyleId1);
            //paragraphProperties1.Indentation = center;
            Run lineBreak = new Run(new Break());
            Run run1 = new Run();
            Text text1 = new Text();
            Text text2 = new Text();
            Text text3 = new Text();
            Text text4 = new Text();
            //var lineBreak = "<w:br/>";

            var runProp = new RunProperties();

            var runFont = new RunFonts { Ascii = "Arial" };

            // 48 half-point font size
            var size = new FontSize { Val = new StringValue("24") };

            runProp.Append(runFont);
            runProp.Append(size);

            run1.PrependChild(runProp);

            text1.Text = "RE:                  " + iMEHearderData.ClaimantName;
            text2.Text = "CLAIM:            " + iMEHearderData.Claim;
            text3.Text = "WCB:               " + iMEHearderData.WCB;
            string date = iMEHearderData.DateOfExam != null ? Convert.ToDateTime(iMEHearderData.DateOfExam).ToString("MM/dd/yyyy") : "";
            text4.Text = "DateOfExam:   " + date;

            run1.Append(text1);
            run1.Append(new Break());
            run1.Append(text2);
            run1.Append(new Break());
            run1.Append(text3);
            run1.Append(new Break());
            run1.Append(text4);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);
            paragraph1.Append(new Break());

            header1.Append(paragraph1);
            //header1.Append(table);

            part.Header = header1;
        }
    }

}
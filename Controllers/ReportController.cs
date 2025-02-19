using ClassLibrary6;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Xml;
using Telerik.Reporting;
using Telerik.Reporting.Drawing;

namespace programmatic_asp_net_core.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ReportController : ControllerBase
    {

        public class RequiredFields
        {
            public List<string> DataSets { get; set; } = new List<string>();
            public List<string> Parameters { get; set; } = new List<string>();
            public List<string> ReportParametersLayout { get; set; } = new List<string>();
            public List<string> ReportSections { get; set; } = new List<string>();
            public List<Fields> HeaderFields { get; set; } = new List<Fields>();
        }

        public class Fields
        {
            public string ColumnName { get; set; }
            public string ColumnType { get; set; }
        }

        [HttpPost("RenderReport")]
        public bool RenderReport()
        {
            var reportProcessor = new Telerik.Reporting.Processing.ReportProcessor();
            // get the root app path programmatically
            var rootAppPath = Directory.GetCurrentDirectory();
            var csReportFile = Path.Combine(rootAppPath, "Reports", "Report1.cs");
            var report = new Report1();
            var reportSource = new Telerik.Reporting.InstanceReportSource();
            reportSource.ReportDocument = report;


            Telerik.Reporting.Processing.RenderingResult result = reportProcessor.RenderReport("pdf", reportSource, new System.Collections.Hashtable());

            if (!result.HasErrors)
            {
                string fileName = Path.GetFileNameWithoutExtension(csReportFile) + "." + result.Extension;
                string filePath = Path.Combine(rootAppPath, fileName);

                using (System.IO.FileStream fs = new System.IO.FileStream(filePath, System.IO.FileMode.Create))
                {
                    fs.Write(result.DocumentBytes, 0, result.DocumentBytes.Length);
                }
            }
            return true;
        }
        [HttpPost("CreateTrdpReport")]
        public bool CreateTrdpReport()
        {
            //RequiredFields requiredFields = RdlcReader("Reports\\RdlcReports\\Report.rdlc");
            RequiredFields requiredFields = new RequiredFields
            {
                HeaderFields = new List<Fields>
                {
                    new Fields { ColumnName = "ID", ColumnType = "Integer" },
                    new Fields { ColumnName = "ISIN", ColumnType = "String" },
                    new Fields { ColumnName = "IsActive", ColumnType = "Boolean" },
                    new Fields { ColumnName = "StartTime", ColumnType = "String" },
                    new Fields { ColumnName = "EndTime", ColumnType = "String" },
                    new Fields { ColumnName = "ResubmissionTime", ColumnType = "String" },
                    new Fields { ColumnName = "Location", ColumnType = "String" },
                    new Fields { ColumnName = "CreatedBy", ColumnType = "String" },
                    new Fields { ColumnName = "CreatedDate", ColumnType = "DateTime" },
                    new Fields { ColumnName = "UpdatedBy", ColumnType = "String" },
                    new Fields { ColumnName = "UpdatedDate", ColumnType = "DateTime" }
                }
            };
            Telerik.Reporting.Report rpt = new Telerik.Reporting.Report();

            //// Define the Web Service Data Source
            //var webServiceDataSource = new Telerik.Reporting.WebServiceDataSource
            //{
            //    ServiceUrl = @"http://localhost:7110/Report/GetTelerikJSONReportData/",
            //    Body = "{\r\n    \"AdditionalDetails\": null,\r\n    \"AppID\": 0,\r\n    \"AppName\": \"Alliance.MS.ALEXI\",\r\n    \"BusinessGroup\": null,\r\n    \"Entity\": \"{\\\"DataSetName\\\":\\\"dsISINConfig\\\",\\\"ConfigJsonReportName\\\":\\\"SampleDS\\\",\\\"ReportPackRunDetailsId\\\":43944,\\\"ReportPackRunId\\\":40142,\\\"Status\\\":\\\"InProgress\\\",\\\"ReportName\\\":\\\"ISINConfigReport\\\",\\\"AppName\\\":\\\"CHARM\\\",\\\"ReportParams\\\":\\\"{\\\\\\\"AccountNumber\\\\\\\":\\\\\\\"@AccountNumber\\\\\\\"}\\\"}\",\r\n    \"SID\": null,\r\n    \"UniqueRefId\": null\r\n}",
            //    Method = WebServiceRequestMethod.Post,
            //    // Optionally, set parameters or headers if needed
            //    Parameters = { new WebServiceParameter("Content-Type", WebServiceParameterType.Header, "application/json") },
            //    DataFormat = WebServiceResponseFormat.Json,
            //};
            var jsonDataSource = new Telerik.Reporting.JsonDataSource
            {
                Source = @"[
                  {
                      ""ID"": 4,
                      ""ISIN"": ""below 10"",
                      ""StartTime"": ""09:00:00"",
                      ""EndTime"": ""10:20:00"",
                      ""ResubmissionTime"": ""10:21:00"",
                      ""TimestampOverride"": ""09:35:00"",
                      ""Location"": ""Global"",
                      ""IsActive"": false,
                      ""CreatedBy"": ""ABC"",
                      ""CreatedDate"": ""2021-09-18T09:53:55.05"",
                      ""UpdatedBy"": ""XYZ"",
                      ""UpdatedDate"": ""2021-09-18T10:52:30.72"",
                      ""BusinessGroup"": ""222""
                  },
                  {
                      ""ID"": 5,
                      ""ISIN"": ""above 10 below 20"",
                      ""StartTime"": ""11:00:00"",
                      ""EndTime"": ""13:00:00"",
                      ""ResubmissionTime"": ""13:11:00"",
                      ""TimestampOverride"": ""13:00:00"",
                      ""Location"": ""Global"",
                      ""IsActive"": true,
                      ""CreatedBy"": ""LMN"",
                      ""CreatedDate"": ""2022-01-22T10:28:35.26"",
                      ""UpdatedBy"": ""PQR"",
                      ""UpdatedDate"": ""2022-09-08T15:23:28.103"",
                      ""BusinessGroup"": ""111""
                  },
                  {
                      ""ID"": 7,
                      ""ISIN"": ""above 20 above 20 above 20 above 20 above 20 above 20"",
                      ""StartTime"": ""17:00:00"",
                      ""EndTime"": ""22:00:00"",
                      ""ResubmissionTime"": ""22:11:00"",
                      ""TimestampOverride"": ""22:00:00"",
                      ""Location"": ""Global"",
                      ""IsActive"": true,
                      ""CreatedBy"": ""MNO"",
                      ""CreatedDate"": ""2022-02-18T10:39:40.84"",
                      ""UpdatedBy"": ""RST"",
                      ""UpdatedDate"": ""2022-03-28T08:42:46.21"",
                      ""BusinessGroup"": ""333""
                  }
                ]"
            };

            //rpt.DataSource = jsonDataSource;

            // Create a new table
            var table = new Telerik.Reporting.Table();
            table.DataSource = jsonDataSource;
            // Set the table's location and size
            table.Location = new PointU(Unit.Cm(1), Unit.Cm(1));
            table.Size = new SizeU(Unit.Cm(18), Unit.Cm(5));


            foreach (var field in requiredFields.HeaderFields)
            {


                var tableGroup = new TableGroup { Name = field.ColumnName };
                var columnHeaderTextBox = new TextBox
                {
                    Name = "headerTextBox",
                    Value = field.ColumnName,
                    //Size = new SizeU(Unit.Cm(6), Unit.Cm(1))
                    // how to use a random width between 6 and 2?
                    Size = new SizeU(Unit.Cm(2 + new Random().Next(5)), Unit.Cm(1))
                };

                tableGroup.ReportItem = columnHeaderTextBox;
                table.ColumnGroups.Add(tableGroup);
                //table.Body.Columns.Add(new TableBodyColumn(Unit.Cm(6)));

            }

            // Define a row group (if needed)
            var rowGroup = new TableGroup { Name = "rowGroup" };
            rowGroup.Groupings.Add(new Telerik.Reporting.Grouping(""));
            table.RowGroups.Add(rowGroup);
            // Add rows and cells to the table
            var row = new TableBodyRow(Unit.Cm(1));
            table.Body.Rows.Add(row);

            int columnIndex = 0;

            foreach (var field in requiredFields.HeaderFields)
            {
                // Create a text box for each field value
                var textBox = new TextBox
                {
                    Value = $"= Fields.{field.ColumnName}", // Bind to the field value
                    Size = new SizeU(Unit.Cm(6), Unit.Cm(1))
                };

                // Set the cell content in the table body
                table.Body.SetCellContent(0, columnIndex, textBox);

                // Increment the column index
                columnIndex++;
            }


            //Create Section
            DetailSection ds = new DetailSection();
            // Assume you have a reference to your table object here, e.g. 'var table = ...'

            // 1) Create a FormattingRule for alternate rows (RowNumber() modulo 2)
            var alternateRowRule = new Telerik.Reporting.Drawing.FormattingRule
            {
                Filters = {
                    new Telerik.Reporting.Filter("= RowNumber() % 2", Telerik.Reporting.FilterOperator.Equal, "=1")
                },
                Style = { BackgroundColor = System.Drawing.Color.LightGray }
            };

            var conditionalFormattingRuleBelow10 = new Telerik.Reporting.Drawing.FormattingRule
            {
                Filters =
                {
                    new Telerik.Reporting.Filter("=ReportItem.Value.Length", Telerik.Reporting.FilterOperator.LessThan, "=10")
                },
                Style =
                {
                    Font =
                    {
                        Size = Telerik.Reporting.Drawing.Unit.Point(16),
                    }
                }
            };

            var conditionalFormattingRuleAbove10Below20 = new Telerik.Reporting.Drawing.FormattingRule
            {
                Filters =
                {
                    new Telerik.Reporting.Filter("ReportItem.Value.Length", Telerik.Reporting.FilterOperator.GreaterOrEqual, "=10"),
                    new Telerik.Reporting.Filter("ReportItem.Value.Length", Telerik.Reporting.FilterOperator.LessThan, "=20")
                },
                Style =
                {
                    Font =
                    {
                        Size = Telerik.Reporting.Drawing.Unit.Point(14),
                    }
                }
            };

            var conditionalFormattingRuleAbove20 = new Telerik.Reporting.Drawing.FormattingRule
            {
                Filters =
                {
                    new Telerik.Reporting.Filter("ReportItem.Value.Length", Telerik.Reporting.FilterOperator.GreaterOrEqual, "=20")
                },
                Style =
                {
                    Font =
                    {
                        Size = Telerik.Reporting.Drawing.Unit.Point(12),
                    }
                }
            };

            foreach (var reportItem in table.Items)
            {
                reportItem.ConditionalFormatting.Add(conditionalFormattingRuleBelow10);
                reportItem.ConditionalFormatting.Add(conditionalFormattingRuleAbove10Below20);
                reportItem.ConditionalFormatting.Add(conditionalFormattingRuleAbove20);
            };




        ds.Items.Add(table);
            rpt.Items.Add(ds);
            var filePath = "C:\\Users\\petodoro\\ISINRDLCConversion.trdp";
            // Use ReportPackager to save the report as a TRDP file
            using (var fileStream = new FileStream(filePath, FileMode.Create))
            {
                ReportPackager reportPackager = new ReportPackager();
                reportPackager.Package(rpt, fileStream);
            }

            Console.WriteLine("Report saved as " + filePath);
            return true;
        }

        [HttpGet("RdlcReader")]
        public RequiredFields RdlcReader(string filePath)
        {
            RequiredFields fields = new RequiredFields();

            // Create an instance of XmlDocument
            XmlDocument xmlDoc = new XmlDocument();

            try
            {
                // Load the RDLC file
                xmlDoc.Load(filePath);

                // Access the root element
                XmlElement root = xmlDoc.DocumentElement;

                // Example: Print the root element name
                Console.WriteLine("Root Element: " + root.Name);

                // Example: Iterate through child nodes
                foreach (XmlNode node in root.ChildNodes)
                {
                    Console.WriteLine("Node Name: " + node.Name);

                    if (node.Name == "DataSets")
                    {
                        XmlNodeList dataSetNodes = xmlDoc.GetElementsByTagName("DataSet");

                        foreach (XmlNode dataSetNode in dataSetNodes)
                        {
                            foreach (XmlNode fieldNode in dataSetNode.ChildNodes)
                            {
                                if (fieldNode.Name == "Fields")
                                {
                                    foreach (XmlNode childNode in fieldNode.ChildNodes)
                                    {
                                        XmlAttribute attribute = childNode.Attributes["Name"];

                                        Fields fields1 = new Fields();

                                        if (childNode.LastChild.Name == "rd:TypeName")
                                        {
                                            fields1.ColumnType = childNode.LastChild.InnerText;
                                            Console.WriteLine("  TypeName: " + childNode.LastChild.InnerText);
                                        }

                                        fields1.ColumnName = attribute.Value;

                                        Console.WriteLine("Field Name: " + attribute.Value);

                                        fields.HeaderFields.Add(fields1);
                                    }
                                }
                            }
                        }
                        foreach (XmlNode dataSetNode in node.ChildNodes)
                        {
                            if (dataSetNode.Name == "DataSet")
                            {
                                XmlAttribute nameAttribute = dataSetNode.Attributes["Name"];
                                if (nameAttribute != null)
                                {
                                    fields.DataSets.Add(nameAttribute.Value);
                                    Console.WriteLine("Dataset Name: " + nameAttribute.Value);
                                }
                            }
                        }
                    }
                    if (node.Name == "ReportParameters")
                    {
                        foreach (XmlNode paramNode in node.ChildNodes)
                        {
                            if (paramNode.Name == "ReportParameter")
                            {
                                XmlAttribute nameAttribute = paramNode.Attributes["Name"];
                                if (nameAttribute != null)
                                {
                                    fields.Parameters.Add(nameAttribute.Value);
                                    Console.WriteLine("Report Parameter Name: " + nameAttribute.Value);
                                }
                            }
                        }
                    }
                    //if(node.Name == "ReportSections")
                    //{
                    //    foreach (XmlNode paramNode in node.ChildNodes)
                    //    {
                    //        if (paramNode.Name == "ReportSection")
                    //        {

                    //           fields.ReportParametersLayout.Add(paramNode.InnerText);
                    //           Console.WriteLine("Report Parameters Layout: " + paramNode.InnerText);

                    //        }
                    //    }
                    //}
                    if (node.Name == "ReportParametersLayout")
                    {
                        foreach (XmlNode paramNode in node.ChildNodes)
                        {
                            if (paramNode.Name == "GridLayoutDefinition")
                            {
                                fields.ReportParametersLayout.Add(paramNode.InnerText);
                                Console.WriteLine("Report Parameters Layout: " + paramNode.InnerText);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }

            return fields;
        }

        /// <summary>
        /// GetTelerikJSONReportData
        /// </summary>
        /// <param name="apiInput"></param>
        /// <returns></returns>
        [HttpPost("GetTelerikJSONReportData")]
        public IActionResult GetTelerikJSONReportData()
        {
            var data = System.IO.File.ReadAllText("DummyData.json");
            List<dynamic> lst = new List<dynamic>();
            lst.AddRange(JsonConvert.DeserializeObject<List<dynamic>>(data));
            return Ok(lst);

        }

    }

}



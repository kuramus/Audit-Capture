using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using AuditCapture.LoginWindow;

namespace AuditCapture
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            if (DesktopDirCheck.IsChecked.Value)
            {
                SaveLocation.IsEnabled = false;
            }
            else
            {
                SaveLocation.IsEnabled = true;
            }

            FetchXml.Text = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                                <entity name='msdyn_workorder'>
                                    <attribute name='msdyn_name'/>
                                    <filter type='and'>
                                        <condition attribute='msdyn_workorderid' operator='eq' value='3' />
                                    </filter>
                                </entity>
                            </fetch>";

            ColumnCount = 0;
            ColumnSet = new Dictionary<string, List<ColumnDetail>>();
            ExcelColumns = new Dictionary<int, string>();
            ColumnSetUp();
            AuditDetailsColl = new List<AttributeAuditDetail>();
        }

        #region Global properties
        public List<AttributeAuditDetail> AuditDetailsColl { get; set; }
        public int ColumnCount { get; set; }
        public Dictionary<string, List<ColumnDetail>> ColumnSet { get; set; }
        public Dictionary<int, string> ExcelColumns { get; set; }
        public class ColumnDetail
        {
            public int position { get; set; }
            public string Column { get; set; }
            public string AttributeName { get; set; }
        }
        public static int rowCount { get; set; }
        public Microsoft.Office.Interop.Excel.Application excel { get; set; }
        public OrganizationServiceProxy service { get; set; }
        #endregion

        #region Audit Retrieval
        public async Task<string> retrieveAudit()
        {
            await setLabelAsync(totalRecCount, "Connecting to CRM");
            string url = DynURL.Text;
            string userName = DynUserName.Text;
            string passwrd = DynPass.Password;
            if (ConnectionManager._service == null)
            {
                ConnectionManager.createCRMConnection(url, userName, passwrd);
            }

            var service = ConnectionManager._service;

            // Load Excel application
            excel = new Microsoft.Office.Interop.Excel.Application();

            // Create empty workbook
            excel.Workbooks.Add();

            // Create Worksheet from active sheet
            Microsoft.Office.Interop.Excel._Worksheet workSheet = excel.ActiveSheet;

            try
            {
                string fetchXml = FetchXml.Text;
                if (fetchXml != null)
                {
                    await setLabelAsync(totalRecCount, "Querying...");
                    FetchExpression fetchExp = new FetchExpression(fetchXml);
                    EntityCollection entColl = service.RetrieveMultiple(fetchExp);
                    if (entColl.Entities.Count > 0)
                    {
                        await setLabelAsync(totalRecCount, entColl.Entities.Count.ToString());
                        await setLabelAsync(processRecCount, "Fetching Audit History...");
                        int index = 0;

                        // Defining header cells
                        workSheet.Cells[1, "A"] = "Title";
                        workSheet.Cells[1, "B"] = "Entity Name";
                        workSheet.Cells[1, "C"] = "Action";
                        workSheet.Cells[1, "D"] = "Operation";
                        workSheet.Cells[1, "E"] = "Audit Record Created On";
                        workSheet.Cells[1, "F"] = "Attribute Name";
                        workSheet.Cells[1, "G"] = "Old Value";
                        workSheet.Cells[1, "H"] = "New Value";

                        rowCount = 2;

                        Dictionary<int, AuditDetailCollection> AuditDetailsFetched = new Dictionary<int, AuditDetailCollection>();

                        #region Execute Multiple with Results
                        // Create an ExecuteMultipleRequest object.
                        ExecuteMultipleRequest requestWithResults = new ExecuteMultipleRequest()
                        {
                            // Assign settings that define execution behavior: continue on error, return responses. 
                            Settings = new ExecuteMultipleSettings()
                            {
                                ContinueOnError = false,
                                ReturnResponses = true
                            },
                            // Create an empty organization request collection.
                            Requests = new OrganizationRequestCollection()
                        };

                        // Add a CreateRequest for each entity to the request collection.
                        foreach (Entity ent in entColl.Entities)
                        {
                            RetrieveRecordChangeHistoryRequest changeHistoryRequest = new RetrieveRecordChangeHistoryRequest();
                            changeHistoryRequest.Target = ent.ToEntityReference();
                            requestWithResults.Requests.Add(changeHistoryRequest);
                        }

                        // Execute all the requests in the request collection using a single web method call.
                        ExecuteMultipleResponse changeHistoryResponses =
                            (ExecuteMultipleResponse)service.Execute(requestWithResults);

                        // Display the results returned in the responses.
                        foreach (var changeHistoryResponse in changeHistoryResponses.Responses)
                        {
                            // A valid response.
                            if (changeHistoryResponse.Response != null)
                            {
                                AuditDetailCollection details = ((RetrieveRecordChangeHistoryResponse)(changeHistoryResponse.Response)).AuditDetailCollection;
                                AuditDetailsFetched.Add(changeHistoryResponse.RequestIndex, details);
                            }

                            // An error has occurred.
                            else if (changeHistoryResponse.Fault != null)
                            {
                                MessageBox.Show("Exception: There was a problem in extacting the Audit History. " + changeHistoryResponse.Fault.ToString());
                            }
                        }
                        #endregion

                        foreach (KeyValuePair<int, AuditDetailCollection> auditItem in AuditDetailsFetched)
                        {
                            index++;
                            if (index > 0)
                            {
                                await setLabelAsync(processRecCount, index.ToString());
                                var title = "Not Found";

                                var index_ = auditItem.Key;
                                if (entColl.Entities[auditItem.Key].Attributes.ContainsKey(TitleAttribute.Text))
                                {
                                    title = entColl.Entities[auditItem.Key].GetAttributeValue<string>(TitleAttribute.Text);
                                }
                                CaptureAuditDetails(auditItem.Value, workSheet, title);
                            }
                        }

                        // Apply some predefined styles for data to look nicely
                        workSheet.Range["A1"].AutoFormat(Microsoft.Office.Interop.Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic1);

                        string fileName = "";
                        // Define filename
                        if (DesktopDirCheck.IsChecked.Value)
                        {
                            fileName = string.Format(@"{0}\" + FileName.Text, Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));
                        }
                        else if (SaveLocation.Text != "" && FileName.Text != null)
                        {
                            fileName = string.Format(@"{0}\" + FileName.Text, SaveLocation.Text);
                        }

                        // Save this data as a file
                        if (fileName != "")
                        {
                            workSheet.SaveAs(fileName);
                        }

                        // Display SUCCESS message
                        MessageBox.Show(string.Format("The file '{0}' is saved successfully!", fileName));
                    }
                    else
                    {
                        MessageBox.Show("Connection Successful. No records retrieved.");
                    }
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Exception: There was a problem saving Excel file!\n" + exception.Message);
            }
            finally
            {
                // Quit Excel application
                excel.Quit();

                // Release COM objects (very important!)
                if (excel != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

                if (workSheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet);

                // Empty variables
                excel = null;
                workSheet = null;

                // Force garbage collector cleaning
                GC.Collect();
            }
            ConnectionManager._service = null;
            return "Success";
        }

        private void CaptureAuditDetails(AuditDetailCollection details, Microsoft.Office.Interop.Excel._Worksheet workSheet, string title)
        {
            string actionFilter = ((ComboBoxItem)AuditActionFilter.SelectedItem).Content.ToString();
            foreach (var detail in details.AuditDetails)
            {
                // Write out some of the change history information in the audit record. 
                var record = detail.AuditRecord;

                if (actionFilter == "All" || (actionFilter != "All" && actionFilter == record.FormattedValues["action"]))
                {
                    var ExportExcelObj = new AddAuditRecordToExcel()
                    {
                        EntityName = record.LogicalName,
                        Action = record.FormattedValues["action"],
                        Operation = record.FormattedValues["operation"],
                        CreatedOn = record.GetAttributeValue<DateTime>("createdon").ToLocalTime().ToString(),
                        Attributes = new List<Attribute>()
                    };

                    var detailType = detail.GetType();
                    if (detailType == typeof(AttributeAuditDetail))
                    {
                        var attributeDetail = (AttributeAuditDetail)detail;

                        // Display the old and new attribute values.
                        if (attributeDetail.NewValue != null)
                        {
                            foreach (KeyValuePair<String, object> attribute in attributeDetail.NewValue.Attributes)
                            {
                                String oldValue = "(no value)", newValue = "(no value)";

                                var type = attribute.Value.GetType().ToString();

                                // Display the lookup values of those attributes that do not contain strings.
                                if (attributeDetail.OldValue != null)
                                {
                                    if (attributeDetail.OldValue.Contains(attribute.Key))
                                    {
                                        switch (attribute.Value.GetType().ToString())
                                        {
                                            case "Microsoft.Xrm.Sdk.OptionSetValue":
                                                oldValue = attributeDetail.OldValue.GetAttributeValue<OptionSetValue>(attribute.Key).Value.ToString();
                                                break;

                                            case "Microsoft.Xrm.Sdk.EntityReference":
                                                var ref_ = attributeDetail.OldValue.GetAttributeValue<EntityReference>(attribute.Key);
                                                oldValue = ref_.Name + ", {" + ref_.Id + "}";
                                                break;

                                            default:
                                                oldValue = attributeDetail.OldValue[attribute.Key].ToString();
                                                break;
                                        }
                                    }
                                }

                                switch (attribute.Value.GetType().ToString())
                                {
                                    case "Microsoft.Xrm.Sdk.OptionSetValue":
                                        newValue = attributeDetail.NewValue.GetAttributeValue<OptionSetValue>(attribute.Key).Value.ToString();
                                        break;

                                    case "Microsoft.Xrm.Sdk.EntityReference":
                                        var ref_ = attributeDetail.NewValue.GetAttributeValue<EntityReference>(attribute.Key);
                                        newValue = ref_.Name + ", {" + ref_.Id + "}";
                                        break;

                                    default:
                                        newValue = attributeDetail.NewValue[attribute.Key].ToString();
                                        break;
                                }

                                ExportExcelObj.Attributes.Add(new Attribute()
                                {
                                    AttributeName = attribute.Key,
                                    OldValue = oldValue,
                                    NewValue = newValue
                                });
                            }
                        }

                        if (attributeDetail.OldValue != null)
                        {
                            foreach (KeyValuePair<String, object> attribute in attributeDetail.OldValue.Attributes)
                            {
                                String oldValue = "(no value)";

                                switch (attribute.Value.GetType().ToString())
                                {
                                    case "Microsoft.Xrm.Sdk.OptionSetValue":
                                        oldValue = attributeDetail.OldValue.GetAttributeValue<OptionSetValue>(attribute.Key).Value.ToString();
                                        break;

                                    case "Microsoft.Xrm.Sdk.EntityReference":
                                        var ref_ = attributeDetail.OldValue.GetAttributeValue<EntityReference>(attribute.Key);
                                        oldValue = ref_.Name + ", {" + ref_.Id + "}";
                                        break;

                                    default:
                                        oldValue = attributeDetail.OldValue[attribute.Key].ToString();
                                        break;
                                }

                                if (attributeDetail.NewValue != null)
                                {
                                    if (!attributeDetail.NewValue.Contains(attribute.Key))
                                    {
                                        ExportExcelObj.Attributes.Add(new Attribute()
                                        {
                                            AttributeName = attribute.Key,
                                            OldValue = oldValue,
                                            NewValue = "(no value)"
                                        });
                                    }
                                }
                                else
                                {
                                    ExportExcelObj.Attributes.Add(new Attribute()
                                    {
                                        AttributeName = attribute.Key,
                                        OldValue = oldValue,
                                        NewValue = "(no value)"
                                    });
                                }
                            }
                        }
                    }
                    rowCount = ExportExcelObj.ExportToExcel(workSheet, rowCount, title);
                }
            }
        }

        public async Task<string> PullDataFromAudit()
        {
            await setLabelAsync(totalRecCount, "Connecting to CRM");

            // Load Excel application
            excel = new Microsoft.Office.Interop.Excel.Application();

            // Create empty workbook
            excel.Workbooks.Add();

            // Create Worksheet from active sheet
            Microsoft.Office.Interop.Excel._Worksheet workSheet = excel.ActiveSheet;

            try
            {
                string fetchXml = FetchXml.Text;
                if (fetchXml != null)
                {
                    await setLabelAsync(totalRecCount, "Querying...");
                    FetchExpression fetchExp = new FetchExpression(fetchXml);
                    EntityCollection entColl = service.RetrieveMultiple(fetchExp);
                    if (entColl.Entities.Count > 0)
                    {
                        await setLabelAsync(totalRecCount, entColl.Entities.Count.ToString());
                        await setLabelAsync(processRecCount, "Fetching Audit History...");
                        int index = 0;

                        Dictionary<int, AuditDetailCollection> AuditDetailsFetched = new Dictionary<int, AuditDetailCollection>();

                        if (entColl.Entities.Count < 1000)
                        {
                            GetAuditDetails(service, entColl);
                        }
                        else
                        {
                            EntityCollection partialEntColl = new EntityCollection();

                            int rounds = (entColl.Entities.Count / 1000);
                            for (int i = 0; i < rounds; i++)
                            {
                                for (int num = 0; num < 1000; num++)
                                {
                                    partialEntColl.Entities.Add(entColl.Entities[((i * 1000) + num)]);
                                }

                                GetAuditDetails(service, partialEntColl);
                                partialEntColl = new EntityCollection();
                            }

                            for (int num = 0; num < (entColl.Entities.Count % 1000); num++)
                            {
                                partialEntColl.Entities.Add(entColl.Entities[((rounds * 1000) + num)]);
                            }

                            GetAuditDetails(service, partialEntColl);
                            partialEntColl = new EntityCollection();
                        }

                        // Creation of header cells
                        workSheet.Cells[1, "A"] = "Title";
                        rowCount = 2;
                        ColumnCount = 0;

                        foreach (AttributeAuditDetail auditItem in AuditDetailsColl)
                        {
                            index++;
                            if (index > 0)
                            {
                                await setLabelAsync(processRecCount, index.ToString());
                                PullDataFromAttributeAuditDetail(auditItem, workSheet, index.ToString());
                            }
                        }

                        // Apply some predefined styles for data to look nicely
                        workSheet.Range["A1"].AutoFormat(Microsoft.Office.Interop.Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic1);

                        string fileName = "";
                        // Define filename
                        if (DesktopDirCheck.IsChecked.Value)
                        {
                            fileName = string.Format(@"{0}\" + FileName.Text, Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));
                        }
                        else if (SaveLocation.Text != "" && FileName.Text != null)
                        {
                            fileName = string.Format(@"{0}\" + FileName.Text, SaveLocation.Text);
                        }

                        // Save this data as a file
                        if (fileName != "")
                        {
                            workSheet.SaveAs(fileName);
                        }

                        // Display SUCCESS message
                        MessageBox.Show(string.Format("The file '{0}' is saved successfully!", fileName));
                    }
                    else
                    {
                        MessageBox.Show("Connection Successful. No records retrieved.");
                    }
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Exception: There was a problem saving Excel file!\n" + exception.Message);
            }
            finally
            {
                // Quit Excel application
                excel.Quit();

                // Release COM objects (very important!)
                if (excel != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

                if (workSheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet);

                // Empty variables
                excel = null;
                workSheet = null;

                // Force garbage collector cleaning
                GC.Collect();
            }
            ConnectionManager._service = null;
            return "Success";
        }

        public void GetAuditDetails(IOrganizationService service, EntityCollection entColl)
        {
            #region Execute Multiple with Results
            // Create an ExecuteMultipleRequest object.
            ExecuteMultipleRequest requestWithResults = new ExecuteMultipleRequest()
            {
                // Assign settings that define execution behavior: continue on error, return responses. 
                Settings = new ExecuteMultipleSettings()
                {
                    ContinueOnError = false,
                    ReturnResponses = true
                },
                // Create an empty organization request collection.
                Requests = new OrganizationRequestCollection()
            };

            // Add a CreateRequest for each entity to the request collection.
            foreach (Entity ent in entColl.Entities)
            {
                var auditDetailsRequest = new RetrieveAuditDetailsRequest()
                {
                    AuditId = ent.Id
                };
                requestWithResults.Requests.Add(auditDetailsRequest);
            }

            // Execute all the requests in the request collection using a single web method call.
            ExecuteMultipleResponse AuditDetailResponses =
                (ExecuteMultipleResponse)service.Execute(requestWithResults);

            foreach (var AuditDetailResponse in AuditDetailResponses.Responses)
            {
                if (AuditDetailResponse.Response != null)
                {
                    if (((RetrieveAuditDetailsResponse)(AuditDetailResponse.Response)).Results != null)
                    {
                        if (((RetrieveAuditDetailsResponse)(AuditDetailResponse.Response)).AuditDetail != null)
                        {
                            AttributeAuditDetail AuditDetail_ = (AttributeAuditDetail)((RetrieveAuditDetailsResponse)(AuditDetailResponse.Response)).AuditDetail;
                            AuditDetailsColl.Add(AuditDetail_);
                        }
                    }
                }
                else if (AuditDetailResponse.Fault != null)
                {
                    MessageBox.Show("Exception: There was a problem in extacting the Audit History. " + AuditDetailResponse.Fault.ToString());
                }
            }
            #endregion
        }

        private void PullDataFromAttributeAuditDetail(AttributeAuditDetail attributeDetail, Microsoft.Office.Interop.Excel._Worksheet workSheet, string title)
        {
            var AttributeColl = attributeDetail.OldValue.Attributes;

            if (ColumnCount != AttributeColl.Count)
            {
                // Add Columns
                foreach (var ColumnName in AttributeColl.Keys)
                {
                    if (!ColumnSet.ContainsKey(ColumnName))
                    {
                        ColumnCount++;

                        string attributeType = AttributeColl[ColumnName].GetType().ToString();
                        if (attributeType == "Microsoft.Xrm.Sdk.EntityReference")
                        {
                            ColumnSet.Add(ColumnName, new List<ColumnDetail>() {
                                new ColumnDetail() { AttributeName = ColumnName, Column = ColumnName + "_GUID", position = ColumnCount },
                                new ColumnDetail() { AttributeName = ColumnName, Column = ColumnName + "_Name", position = (ColumnCount + 1) }
                            });
                            workSheet.Cells[1, ExcelColumns[ColumnCount + 1]] = ColumnName + "_GUID";
                            workSheet.Cells[1, ExcelColumns[ColumnCount + 2]] = ColumnName + "_Name";
                            ColumnCount++;
                        }
                        else
                        {
                            ColumnSet.Add(ColumnName, new List<ColumnDetail>() {
                                new ColumnDetail() { AttributeName = ColumnName, Column = ColumnName, position = ColumnCount }
                            });
                            workSheet.Cells[1, ExcelColumns[ColumnCount + 1]] = ColumnName;
                        }
                    }
                }
            }

            workSheet.Cells[rowCount, "A"] = title;

            foreach (KeyValuePair<string, List<ColumnDetail>> item in ColumnSet)
            {
                string cellValue = "";
                if (AttributeColl.ContainsKey(item.Key))
                {
                    var attribute = AttributeColl[item.Key];

                    switch (attribute.GetType().ToString())
                    {
                        case "Microsoft.Xrm.Sdk.OptionSetValue":
                            cellValue = ((OptionSetValue)attribute).Value.ToString();
                            break;

                        case "Microsoft.Xrm.Sdk.EntityReference":
                            var ref_ = (EntityReference)attribute;
                            workSheet.Cells[rowCount, ExcelColumns[item.Value[1].position + 1]] = ref_.Name;
                            cellValue = ref_.Id.ToString();
                            break;

                        default:
                            cellValue = attribute.ToString();
                            break;
                    }
                    workSheet.Cells[rowCount, ExcelColumns[item.Value[0].position + 1]] = cellValue;
                }
            }
            rowCount++;
        }
        #endregion

        #region Exception Handling
        public static void CreateExceptionString(Exception e)
        {
            MessageBox.Show("Exception thrown. Log is at C:\\schdLog\\testLog. Message: " + e.Message);
            StringBuilder sb = new StringBuilder();
            CreateExceptionString(sb, e, string.Empty);

            var dir = @"c:\schdLog\testLog";
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            StreamWriter log;
            if (!File.Exists(System.IO.Path.Combine(dir, "log.txt")))
                log = new StreamWriter(System.IO.Path.Combine(dir, "log.txt"));
            else
                log = File.AppendText(System.IO.Path.Combine(dir, "log.txt"));

            log.WriteLine(DateTime.Now);
            log.WriteLine(sb.ToString());
            log.WriteLine();
            log.Close();
        }

        private static void CreateExceptionString(StringBuilder sb, Exception e, string indent)
        {
            if (indent == null)
            {
                indent = string.Empty;
            }
            else if (indent.Length > 0)
            {
                sb.AppendFormat("{0}Inner ", indent);
            }

            sb.AppendFormat("Exception Found:\n{0}Type: {1}", indent, e.GetType().FullName);
            sb.AppendFormat("\n{0}Message: {1}", indent, e.Message);
            sb.AppendFormat("\n{0}Source: {1}", indent, e.Source);
            sb.AppendFormat("\n{0}Stacktrace: {1}", indent, e.StackTrace);

            if (e.InnerException != null)
            {
                sb.Append("\n");
                CreateExceptionString(sb, e.InnerException, indent + "  ");
            }
        }
        #endregion

        #region UI Commands
        /// <summary>
        /// Button to login to CRM and create a CrmService Client 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LoginBtn_Click(object sender, RoutedEventArgs e)
        {
            #region Login Control
            // Establish the Login control
            CrmLogin ctrl = new CrmLogin();
            // Wire Event to login response.
            ctrl.ConnectionToCrmCompleted += ctrl_ConnectionToCrmCompleted;
            // Show the dialog.
            ctrl.ShowDialog();

            // Handel return. 
            if (ctrl.CrmConnectionMgr != null && ctrl.CrmConnectionMgr.CrmSvc != null && ctrl.CrmConnectionMgr.CrmSvc.IsReady)
                MessageBox.Show("Good Connect");
            else
                MessageBox.Show("BadConnect");

            #endregion

            #region CRMServiceClient
            if (ctrl.CrmConnectionMgr != null && ctrl.CrmConnectionMgr.CrmSvc != null && ctrl.CrmConnectionMgr.CrmSvc.IsReady)
            {
                CrmServiceClient svcClient = ctrl.CrmConnectionMgr.CrmSvc;
                if (svcClient.IsReady)
                {
                    service = svcClient.OrganizationServiceProxy;
                }
            }
            #endregion
        }

        /// <summary>
        /// Raised when the login form process is completed.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ctrl_ConnectionToCrmCompleted(object sender, EventArgs e)
        {
            if (sender is CrmLogin)
            {
                this.Dispatcher.Invoke(() =>
                {
                    ((CrmLogin)sender).Close();
                });
            }
        }

        async Task setLabelAsync(System.Windows.Controls.Label totalRecCount, string value)
        {
            await Task.Delay(1);
            Application.Current.Dispatcher.Invoke(new Action(() => { totalRecCount.Content = value; }));
        }

        private async void GetAuditBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var result = await retrieveAudit();
            }
            catch (Exception ex)
            {
                CreateExceptionString(ex);
            }
        }

        private void DesktopDirCheck_Checked(object sender, RoutedEventArgs e)
        {
            if (DesktopDirCheck.IsChecked.Value)
            {
                SaveLocation.IsEnabled = false;
            }
            else
            {
                SaveLocation.IsEnabled = true;
            }
        }

        private async void PullAuditBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var result = await PullDataFromAudit();
            }
            catch (Exception ex)
            {
                CreateExceptionString(ex);
            }
        }
        #endregion
        
        #region Excel functions
        private void ColumnSetUp()
        {
            ExcelColumns.Add(1, "A");
            ExcelColumns.Add(2, "B");
            ExcelColumns.Add(3, "C");
            ExcelColumns.Add(4, "D");
            ExcelColumns.Add(5, "E");
            ExcelColumns.Add(6, "F");
            ExcelColumns.Add(7, "G");
            ExcelColumns.Add(8, "H");
            ExcelColumns.Add(9, "I");
            ExcelColumns.Add(10, "J");
            ExcelColumns.Add(11, "K");
            ExcelColumns.Add(12, "L");
            ExcelColumns.Add(13, "M");
            ExcelColumns.Add(14, "N");
            ExcelColumns.Add(15, "O");
            ExcelColumns.Add(16, "P");
            ExcelColumns.Add(17, "Q");
            ExcelColumns.Add(18, "R");
            ExcelColumns.Add(19, "S");
            ExcelColumns.Add(20, "T");
            ExcelColumns.Add(21, "U");
            ExcelColumns.Add(22, "V");
            ExcelColumns.Add(23, "W");
            ExcelColumns.Add(24, "X");
            ExcelColumns.Add(25, "Y");
            ExcelColumns.Add(26, "Z");
        }
        #endregion
    }
}

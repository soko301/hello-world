
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading;
using System.Text.RegularExpressions;

using Event = Bloomberglp.Blpapi.Event;
using Message = Bloomberglp.Blpapi.Message;
using Element = Bloomberglp.Blpapi.Element;
using Name = Bloomberglp.Blpapi.Name;
using Request = Bloomberglp.Blpapi.Request;
using Service = Bloomberglp.Blpapi.Service;
using Session = Bloomberglp.Blpapi.Session;
using SessionOptions = Bloomberglp.Blpapi.SessionOptions;
using EventHandler = Bloomberglp.Blpapi.EventHandler;
using CorrelationID = Bloomberglp.Blpapi.CorrelationID;
using DateTime = System.DateTime;
using DayOfWeek = System.DayOfWeek;
using System.Xml; 


// heres the list of procedures necessary to get bloomberg data
//1) createsession 
//2) processEvent
//3) processRequestDataEvent
//4) processBulkdata - if data retuned is an array

//void "get_BBG" should be called from other classes reference inputs 1) security 2) feilds 3) overrides 4) overrides inputs 5) xmlfilepath, 6) xmlfilename
//this should be usable for all bbg data types - output is saved into xmlfile which can be recalled for use. 

namespace Bloomberglp.Blpapi.Examples
{
    public class getBbgData
    {
        private static readonly Name EXCEPTIONS = new Name("exceptions");
        private static readonly Name FIELD_ID = new Name("fieldId");
        private static readonly Name REASON = new Name("reason");
        private static readonly Name CATEGORY = new Name("category");
        private static readonly Name DESCRIPTION = new Name("description");
        private static readonly Name ERROR_CODE = new Name("errorCode");
        private static readonly Name SOURCE = new Name("source");
        private static readonly Name SECURITY_ERROR = new Name("securityError");
        private static readonly Name MESSAGE = new Name("message");
        private static readonly Name RESPONSE_ERROR = new Name("responseError");
        private static readonly Name SECURITY_DATA = new Name("securityData");
        private static readonly Name FIELD_EXCEPTIONS = new Name("fieldExceptions");
        private static readonly Name ERROR_INFO = new Name("errorInfo");


        private SessionOptions d_sessionOptions;
        private Session d_session;
        public DataSet d_bulkData;
        public DataTable d_data;
        public DataTable bulkTable;
        private DataTable d_overRides;
        public DataTable export;
        private int d_numberOfReturnedSecurities = 0;
        private string xmlFilePath;
        private string xmlFileName;
        private string dtName;
       


        #region methods
        /// <summary>
        /// Initialize form controls
        /// </summary>
        private void initUI()
        {

            string serverHost = "localhost";
            int serverPort = 8194;

            // set sesson options
            d_sessionOptions = new SessionOptions();
            d_sessionOptions.ServerHost = serverHost;
            d_sessionOptions.ServerPort = serverPort;
            // initialize UI controls

            // add columns to data table
            if (d_data == null)
                d_data = new DataTable();
            d_data.Columns.Add("security");
            d_data.AcceptChanges();

            // add columns to data table
            if (bulkTable == null)
                bulkTable = new DataTable();
      
         

            if (d_overRides == null)
                d_overRides = new DataTable();
            d_overRides.Columns.Add("override");
            d_overRides.Columns.Add("value");
            d_overRides.AcceptChanges();

            // set grid data source
            //   dataGridViewData.DataSource = d_data;
        }

        /// <summary>
        /// Add securities to grid
        /// </summary>
        /// <param name="securities"></param>
        private void addSecurities(string securities)
        {
            // Tokenize the string into what (we hope) are Security strings
            char[] sep = { '\r', '\n', '\t', ',' };
            string[] words = securities.Split(sep);
            foreach (string security in words)
            {
                if (security.Trim().Length > 0)
                {
                    // add fields
                    d_data.Rows.Add(security.Trim());
                }
            }

        }

        private void addOverrides(string overide, string value)
        {
            if (overide.Trim().Length > 0)
            {

                char[] sep = { '\r', '\n', '\t', ',' };
                string[] over = overide.Split(sep);
                string[] val = value.Split(sep);


                for (int i = 0; i < over.Length; i++)
                {
                    d_overRides.Rows.Add(over[i], val[i]);
                }

            }

        }
        /// <summary>
        /// Add fields
        /// </summary>
        /// <param name="fields"></param>
        private void addFields(string fields)
        {
            // Tokenize the string into what (we hope) are Security strings
            char[] sep = { '\r', '\n', '\t', ',' };
            string[] words = fields.Split(sep);
            foreach (string field in words)
            {
                if (field.Trim().Length > 0)
                {
                    // add fields
                    if (!d_data.Columns.Contains(field.Trim()))
                    {
                        d_data.Columns.Add(field.Trim());

                    }
                }
            }

        }

      

        /// <summary>
        /// Create session
        /// </summary>
        /// <returns></returns>
        private bool createSession()
        {

            d_session = new Session(d_sessionOptions, new EventHandler(processEvent));
            
            return d_session.Start();
        }



        /// <summary>
        /// Clear security data
        /// </summary>
        private void clearData()
        {
            // clear security count
            d_numberOfReturnedSecurities = 0;

            // clear bulk data
            if (d_bulkData != null)
            {
                d_bulkData.Clear();
                d_bulkData.AcceptChanges();
            }

            if (d_data != null)
            {
                foreach (DataRow row in d_data.Rows)
                {
                    for (int index = 1; index < d_data.Columns.Count; index++)
                    {
                        row[index] = DBNull.Value;
                    }
                }
                d_data.AcceptChanges();
            }

        }

        /// <summary>
        /// Remove all fields from grid
        /// </summary>
        private void clearFields()
        {
            for (int index = d_data.Columns.Count - 1; index > 0; index--)
            {
                d_data.Columns.RemoveAt(index);
            }
            d_data.AcceptChanges();

        }

        /// <summary>
        /// Remove all securities and fields from grid
        /// </summary>
        private void clearAll()
        {
            if (d_bulkData != null)
            {
                d_bulkData.Clear();
                d_bulkData.AcceptChanges();
            }
            clearFields();
            d_data.Rows.Clear();
            d_data.AcceptChanges();
      

        }
        #endregion end methods

        #region Control Events



        public void get_BBG(string sec, string field, string over, string ovrVal, string xmlFilePathIn, string xmlfileNameIn, string dtNameIn)
        {
            clearData();
            initUI();
            addSecurities(sec);
            addFields(field);

            xmlFilePath = xmlFilePathIn;
            xmlFileName = xmlfileNameIn;
            dtName = dtNameIn;


            if (over != null)
            {
                addOverrides(over, ovrVal);
            }
            
            // create session
            if (!createSession())
            {

                return;
            }
            // open reference data service
            if (!d_session.OpenService("//blp/refdata"))
            {

                return;
            }

            Service refDataService = d_session.GetService("//blp/refdata");
            // create reference data request
            Request request = refDataService.CreateRequest("ReferenceDataRequest");
            // set request parameters
            Element securities = request.GetElement("securities");
            Element fields = request.GetElement("fields");
            Element requestOverrides = request.GetElement("overrides");
            request.Set("returnEids", true);
            // populate security
            foreach (DataRow secRow in d_data.Rows)
            {
                securities.AppendValue(secRow["security"].ToString());
            }
            // populate fields
            for (int fieldIndex = 1; fieldIndex < d_data.Columns.Count; fieldIndex++)
                fields.AppendValue(d_data.Columns[fieldIndex].ColumnName);


            if (d_overRides.Rows.Count > 0)
            {
                // populate overrides
                foreach (DataRow row in d_overRides.Rows)
                {
                    Element ovr = requestOverrides.AppendElement();
                    ovr.SetElement(FIELD_ID, row["override"].ToString());
                    ovr.SetElement("value", row["value"].ToString());
                }
            }
            // create correlation id            
            CorrelationID cID = new CorrelationID(1);
            d_session.Cancel(cID);
            // send request
            d_session.SendRequest(request, cID);

           
         
        }
        #endregion

        #region Bloomberg API Events
        /// <summary>
        /// Data event
        /// </summary>
        /// <param name="eventObj"></param>
        /// <param name="session"></param>
        private void processEvent(Event eventObj, Session session)
        {
            

                switch (eventObj.Type)
                {
                    case Event.EventType.RESPONSE:
                        // process final respose for request
                        processRequestDataEvent(eventObj, session);


                        break;
                    case Event.EventType.PARTIAL_RESPONSE:
                        // process partial response
                        processRequestDataEvent(eventObj, session);
                        break;
                    default:
                        processMiscEvents(eventObj, session);
                        break;
                }

            
        }

        /// <summary>
        /// Process subscription data
        /// </summary>
        /// <param name="eventObj"></param>
        /// <param name="session"></param>
        private void processRequestDataEvent(Event eventObj, Session session)
        {
            if (d_numberOfReturnedSecurities == 0)
             
            d_data.BeginLoadData();
            // process message
            foreach (Message msg in eventObj)
            {
                // get message correlation id
                int cId = (int)msg.CorrelationID.Value;
                if (msg.MessageType.Equals(Bloomberglp.Blpapi.Name.GetName("ReferenceDataResponse")))
                {
                    // process security data
                    Element secDataArray = msg.GetElement(SECURITY_DATA);
                    int numberOfSecurities = secDataArray.NumValues;
                    for (int index = 0; index < numberOfSecurities; index++)
                    {
                        Element secData = secDataArray.GetValueAsElement(index);
                        Element fieldData = secData.GetElement("fieldData");
                        d_numberOfReturnedSecurities++;
                        // get security index
                        int rowIndex = secData.GetElementAsInt32("sequenceNumber");
                        if (d_data.Rows.Count > rowIndex)
                        {
                            // get security record
                            DataRow row = d_data.Rows[rowIndex];
                            // check for field error
                            if (secData.HasElement(FIELD_EXCEPTIONS))
                            {
                                // process error
                                Element error = secData.GetElement(FIELD_EXCEPTIONS);
                                for (int errorIndex = 0; errorIndex < error.NumValues; errorIndex++)
                                {
                                    Element errorException = error.GetValueAsElement(errorIndex);
                                    string field = errorException.GetElementAsString(FIELD_ID);
                                    Element errorInfo = errorException.GetElement(ERROR_INFO);
                                    string message = errorInfo.GetElementAsString(MESSAGE);
                                    row[field] = message;
                                }
                            }
                            // check for security error
                            if (secData.HasElement(SECURITY_ERROR))
                            {
                                Element error = secData.GetElement(SECURITY_ERROR);
                                string errorMessage = error.GetElementAsString(MESSAGE);
                                row[1] = errorMessage;
                            }
                            // process data

                          try
                         {


                                foreach (DataColumn col in d_data.Columns)
                                {
                                    String dataValue = string.Empty;
                                    if (fieldData.HasElement(col.ColumnName))
                                    {
                                        Element item = fieldData.GetElement(col.ColumnName);
                                        if (item.IsArray)
                                        {
                                            // bulk field
                                            dataValue = "Bulk Data...";
                                            processBulkData(secData.GetElementAsString("security"), item);
                                        }
                                        else
                                        {
                                            dataValue = item.GetValueAsString();
                                        }
                                         row[col.ColumnName] = dataValue;
                                    }
                                }


                             
                          }
                         catch { }
                        }
                    }
                }
            }
            d_data.EndLoadData();

            if (d_numberOfReturnedSecurities >= d_data.Rows.Count)
            {
                saveXmlDataSet();
            }
            // check if we are done

        }


        private void saveXmlDataSet()
        {
            // filepath and name are passed from get_bbg function used in other class

            string myXMLfile = xmlFilePath + xmlFileName + ".xml";
            DataSet ds = new DataSet();


            //check if current file exist and if so load tables into dataset
            if (File.Exists(myXMLfile))
            {
                // Create new FileStream with which to read the schema.
                System.IO.FileStream fsReadXml = new System.IO.FileStream
                    (myXMLfile, System.IO.FileMode.Open);
                try
                {
                    ds.ReadXml(fsReadXml);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {
                    fsReadXml.Close();
                }
            }
           

          //table name is also passed from get_bbg

            d_data.TableName = dtName;

            //check to see if there is a table for current ccy - if so delete old table and add new data else just at new table
            if (ds.Tables.Contains(dtName))
            {
                ds.Tables.Remove(dtName);

                ds.Tables.Add(d_data);
            }
            else
            {          
                ds.Tables.Add(d_data);
            }

            //write xml file.
       
            ds.WriteXml(myXMLfile);
            

        }

        /// <summary>
        /// Process bulk data
        /// </summary>
        /// <param name="security"></param>
        /// <param name="data"></param>
        public void processBulkData(string security, Element data)
        {
        
            d_data.BeginLoadData();

            //DataTable bulkTable = null;

         
            Element bulk = data.GetValueAsElement(0);
           
            // create columns in data table for bulk data
            foreach (Element item in bulk.Elements)
            {
               
                    d_data.Columns.Add(new DataColumn(item.Name.ToString(), typeof(String)));
            }
            // populate bulk
            int count = 0;
            for (int index = 0; index < data.NumValues; index++)
            {
                bulk = data.GetValueAsElement(index);
                object[] dataArray = new object[bulk.NumElements + 2];
                dataArray[0] = security;
                dataArray[1] = count;
                int dataIndex = 2;
                foreach (Element item in bulk.Elements)
                {
                    dataArray[dataIndex] = item.GetValueAsString();
                    dataIndex++;
                }
                d_data.Rows.Add(dataArray);
                count++;
            }


            d_data.EndLoadData();

       
           
                saveXmlDataSet();
            
     

       

        }

        private void processBulkDataOLD(string security, Element data)
        {
            DataTable bulkTable = null;
            // bulk data dataset
            if (d_bulkData == null)
                d_bulkData = new DataSet();
            // get bulk data
            Element bulk = data.GetValueAsElement(0);
            if (d_bulkData.Tables.Contains(bulk.Name.ToString()))
            {
                // get existing bulk data table
                bulkTable = d_bulkData.Tables[bulk.Name.ToString()];
            }
            else
            {
                // create new bulk data table
                bulkTable = d_bulkData.Tables.Add(bulk.Name.ToString());
            }
            // create column if not already exist
            if (!bulkTable.Columns.Contains("security"))
            {
                bulkTable.Columns.Add(new DataColumn("security", typeof(String)));
                bulkTable.Columns.Add(new DataColumn("Id", typeof(int)));
            }
            // create columns in data table for bulk data
            foreach (Element item in bulk.Elements)
            {
                if (!bulkTable.Columns.Contains(item.Name.ToString()))
                    bulkTable.Columns.Add(new DataColumn(item.Name.ToString(), typeof(String)));
            }
            // populate bulk
            int count = 0;
            for (int index = 0; index < data.NumValues; index++)
            {
                bulk = data.GetValueAsElement(index);
                object[] dataArray = new object[bulk.NumElements + 2];
                dataArray[0] = security;
                dataArray[1] = count;
                int dataIndex = 2;
                foreach (Element item in bulk.Elements)
                {
                    dataArray[dataIndex] = item.GetValueAsString();
                    dataIndex++;
                }
                bulkTable.Rows.Add(dataArray);
                count++;
            }
        }

        private void processMiscEvents(Event eventObj, Session session)
        {
            foreach (Message msg in eventObj)
            {
                switch (msg.MessageType.ToString())
                {
                    case "SessionStarted":
                        // "Session Started"
                        break;
                    case "SessionTerminated":
                    case "SessionStopped":
                        // "Session Terminated"
                        break;
                    case "ServiceOpened":
                        // "Reference Service Opened"
                        break;
                    case "RequestFailure":
                        Element reason = msg.GetElement(REASON);
                        string message = string.Concat("Error: Source-", reason.GetElementAsString(SOURCE),
                            ", Code-", reason.GetElementAsString(ERROR_CODE), ", category-", reason.GetElementAsString(CATEGORY),
                            ", desc-", reason.GetElementAsString(DESCRIPTION));

                        break;
                    default:

                        break;
                }
            }
        }
        #endregion
    }
}
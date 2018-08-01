using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using CsvHelper;						//downloaded from https://github.com/joshclose/csvhelper; 
										//	installed using package manager
using Interop.QBXMLRP2Lib;				//added ref from qb sdk
using System.Web.Script.Serialization;  //added ref System.Web.Extensions from .net framework


namespace Quicken_Proof_of_Concept
{
	class Query
	{
		private string _queryObject;
		public string queryObject
		{
			get
			{
				return _queryObject;
			}
			set
			{
				_queryObject = value;
				updateFileName();
			}
		}
			
		public List<string> attributes { get; set; }
		public string fileName { get; set; }

		public Query()
		{
			queryObject = "";
			attributes = new List<string>();
			fileName = "";
		}
		public Query(string call, List<string> atts)
		{
			queryObject = call;
			attributes = atts;
			updateFileName();
		}

		public void setQueryObject(string name)
		{
			queryObject = name;
			updateFileName();
		}

		public void updateFileName()
		{
			fileName = "Quickbooks" + queryObject + "Results.csv";
		}
		
	}

	class JsonRoot
	{
		public List<Query> listOfQueries { get; set; }
	}

	class Program
	{
		//name of config file to be read
		public const string CONFIG = "config.json";

		static void Main(string[] args)
		{
			bool testing = false; //prints log statements to terminal
			bool debug = false; //more detailed logging
			string jsonFromFile;
			JavaScriptSerializer ser = new JavaScriptSerializer();
			JsonRoot inputRoot;
			List<Query> qList;
			string requestXML;
			string response;
			List<String> filesWrittenTo;

			//parse config file
			jsonFromFile = File.ReadAllText(CONFIG);
			inputRoot = ser.Deserialize<JsonRoot>(jsonFromFile);
			qList = inputRoot.listOfQueries;

			//check what queries were read in
			if (debug) printQueries(qList);

			//create qbxml request
			requestXML = constructRequest(qList);

			//connect to quickbooks
			response = queryQuickBooks(requestXML, testing);
			
			//write query results to csv files
			filesWrittenTo = writeResults(qList, response, testing);

			//report names of files written
			if (testing) Console.WriteLine("\n");
			Console.WriteLine("Wrote to files: ");
			foreach (string name in filesWrittenTo)
			{
				Console.WriteLine(name);
			}

			//wait for user
			Console.ReadKey();
		}
		
		//given list of queries, constructs XML document with headers, returns string
		//parameters: msgList, list of Queries; stopOnError, optl bool that decides error behavior
		static private string constructRequest(List<Query> msgList, bool stopOnError = true)
		{
			XmlDocument doc = new XmlDocument();
			string errorBehavior;
			string messageLanguage = "QBXML";
			string messageSet = "QBXMLMsgsRq";
			List<string> addedQueries = new List<string>();
			int queryIndex = 1;
			string headers = "<?xml version=\"1.0\"?><?qbxml version=\"8.0\">?";
			string requestXml;

			//top nodes: language and message set
			XmlElement language = doc.CreateElement(messageLanguage);
			doc.AppendChild(language);
			XmlElement messageWrapper = doc.CreateElement(messageSet);
			language.AppendChild(messageWrapper);

			//define error behavior ("stopOnError" or "continueOnError")
			if (stopOnError) errorBehavior = "stopOnError";
			else errorBehavior = "continueOnError";
			messageWrapper.SetAttribute("onError", errorBehavior);

			//add each query to message set
			foreach (Query msg in msgList)
			{
				//if duplicate query, skip
				if (addedQueries.Contains(msg.queryObject))
					continue;

				//add message node
				addMessage(doc, ref messageWrapper, msg.queryObject, queryIndex);
				queryIndex++;
				addedQueries.Add(msg.queryObject);
			}
			
			//add headers, convert to string
			requestXml = headers + doc.OuterXml;

			return requestXml;
		}

		//helper for constructRequest: adds a single message to the set
		//parameters: an xml document used to construct elements, the element to append to, 
		//	the query to add, and the index of the query
		static private void addMessage(XmlDocument doc, ref XmlElement msgSet, string msg, int index)
		{
			XmlElement message = doc.CreateElement(msg + "QueryRq");
			msgSet.AppendChild(message);
			message.SetAttribute("requestID", index.ToString());
		}

		//connects to quickbooks
		//parameters: request, qbxml string; log, optl bool that prints completed steps to console if true
		//returns string containing xml document on success, empty string on error
		static private string queryQuickBooks(string request, bool log = false)
		{
			RequestProcessor2 rp = null;
			string ticket = null;
			string response = "";
			
			try
			{
				rp = new RequestProcessor2();

				rp.OpenConnection2("", "Proof of Concept", QBXMLRPConnectionType.localQBD);
				if (log) Console.WriteLine("Log: Connection opened");

				ticket = rp.BeginSession("", QBFileMode.qbFileOpenDoNotCare);
				if (log) Console.WriteLine("Log: Session started");

				response = rp.ProcessRequest(ticket, request);
				if (log) Console.WriteLine("Log: Request processed");
			}
			catch (System.Runtime.InteropServices.COMException ex)
			{
				Console.WriteLine("ERROR Connection problem: " + ex.Message);
			}
			finally
			{
				if (ticket != null)
				{
					rp.EndSession(ticket);
					if (log) Console.WriteLine("Log: Session ended");
				}
				if (rp != null)
				{
					rp.CloseConnection();
					if (log) Console.WriteLine("Log: Connection closed");
				}
			}

			return response;
		}

		//processes results, writes each query to a csv file
		//parameters: reqList, a list of the queries to be written; docString, qbxml results; log, optl bool controls logging
		static private List<String> writeResults(List<Query> reqList, string docString, bool log = false)
		{
			List<String> filesWritten = new List<String>();

			//check if empty
			if (String.IsNullOrEmpty(docString))
			{
				if(log) Console.WriteLine("No results");
				return filesWritten;
			}

			XmlDocument output = new XmlDocument();
			output.LoadXml(docString);

			foreach (Query req in reqList)
			{
				//get all results for this query
				XmlNodeList resultSet = output.GetElementsByTagName(req.queryObject + "QueryRs");

				//check only one result was returned
				if (resultSet.Count == 1)
				{
					XmlNode node = resultSet[0];
					string fileName = req.fileName;
					List<string> attsRqstd = req.attributes;
					XmlNodeList attsFound = node.ChildNodes;

					//get status of query
					if (log)
					{
						XmlAttributeCollection resultAtts = node.Attributes;
						string statusCode = resultAtts.GetNamedItem("statusCode").Value;
						string statusSeverity = resultAtts.GetNamedItem("statusSeverity").Value;
						string statusMessage = resultAtts.GetNamedItem("statusMessage").Value;

						Console.WriteLine("\n\nQuery: " + req.queryObject + "s\n" +
										  "Status code: " + statusCode + "\n" +
										  "Status severity: " + statusSeverity + "\n" +
										  "Status message: " + statusMessage);
					}

					//write results found to new csv file
					fileName = freshName(fileName, filesWritten);
					writeCsv(fileName, attsRqstd, attsFound, log);
					filesWritten.Add(fileName);
				}
				else if (resultSet.Count > 1)
				{
					if (log) Console.WriteLine("Log: repeated query, stopping processing");
				}
				else
				{
					if (log) Console.WriteLine("Log: no results");
				}
			}

			return filesWritten;
		}
		
		//helper for writeResults, recursively checks if name is free, returns fresh name
		static private string freshName(string desiredName, List<string> taken, int index = 0)
		{
			string name = desiredName;

			if (index > 0)
				name = desiredName.Substring(0, desiredName.Length - 4)+ index.ToString() + ".csv";

			if (taken.Contains(name))
				name = freshName(desiredName, taken, ++index);

			return name;
		}

		//helper for writeResults, writes a single query to a csv
		//parameters: file, string name of csv to write; 
		// desiredAttributes, list of attributes to return; 
		// results, XmlNodeList to unpack,
		// log, optl bool controlling logging
		static private void writeCsv(string file, List<String> desiredAttributes, XmlNodeList results, bool log = false)
		{
			using (TextWriter writer = new StreamWriter(file))
			{
				var csv = new CsvWriter(writer);

				//build header
				foreach (string att in desiredAttributes)
				{
					csv.WriteField(att);
				}
				csv.NextRecord();


				//walk results
				foreach (XmlNode item in results)
				{
					List<string> fieldsToWrite = new List<string>();
					bool hasText = false;
					XmlNodeList returnedAttributes = item.ChildNodes;

					if (log)
						Console.WriteLine("\n" + item.Name);

					foreach (string att in desiredAttributes)
					{
						//search result attributes for desired
						XmlDocument dummyDoc = new XmlDocument();
						XmlNode desiredEntry = dummyDoc.CreateNode("text", "dummy", "");
						string foundAttribute;

						foreach (XmlNode aNode in returnedAttributes)
						{
							if (aNode.Name.Equals(att))
							{
								desiredEntry = aNode;
								break;
							}
						}

						//if found, save attribute and set flag, otherwise save blank field marker
						foundAttribute = desiredEntry.InnerText;
						fieldsToWrite.Add(foundAttribute);
						if (!(foundAttribute.Equals("")))
							hasText = true;
						if (log)
							Console.WriteLine(desiredEntry.Name + ": " + foundAttribute);

					}

					//only write if attributes were found
					if (hasText)
					{
						foreach (string field in fieldsToWrite)
						{
							csv.WriteField(field);
						}
						csv.NextRecord();
					}
				}
			}
		}

		// debug logging method
		// prints all queries (queryObject and all attributes)
		static private void printQueries(List<Query> list)
		{
			Console.WriteLine("List of queries found: ");
			foreach (Query q in list)
			{
				Console.WriteLine(q.queryObject);
				foreach (string att in q.attributes)
				{
					Console.WriteLine("\t" + att);
				}
			}
			Console.WriteLine("All queries printed\n");
		}
	}
}
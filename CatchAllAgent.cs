/*
<config>
    <redirect pattern="^john+\.[a-z]+@domain.com$" address="john@domain.com" />
    <redirect pattern="^jane+\.[a-z]+@domain.com$" address="jane.doe@gmail.com" />
</config>
*/

namespace DEA_MTA
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using System.IO;
    using System.Xml;
    using System.Globalization;
    using System.Diagnostics;
    using System.Threading;
    using System.Text.RegularExpressions;
    using Microsoft.Exchange.Data.Transport;
    using Microsoft.Exchange.Data.Transport.Smtp;

    public sealed class CatchAllFactory : SmtpReceiveAgentFactory
    {
        private CatchAllConfig catchAllConfig = new CatchAllConfig();

        public override SmtpReceiveAgent CreateAgent(SmtpServer server)
        {
            return new CatchAllAgent(catchAllConfig, server.AddressBook);
        }
    }

    public class Logger
    {
        private static EventLog logger = null;

        public static void LogError(string message, int id = 0)
        {
            LogEntry(message, id, EventLogEntryType.Information);
        }

        private static void LogEntry(string message, int id, EventLogEntryType logType)
        {
            if (logger == null)
            {
                logger = new EventLog();
                logger.Source = "Exchange DEA";
            }
            logger.WriteEntry(message, logType, id);
        }
    }

    public class CatchAllAgent : SmtpReceiveAgent
    {
        private AddressBook addressBook;
        private CatchAllConfig catchAllConfig;

        public CatchAllAgent(CatchAllConfig catchAllConfig, AddressBook addressBook)
        {
            this.addressBook = addressBook;
            this.catchAllConfig = catchAllConfig;

            this.OnRcptCommand += new RcptCommandEventHandler(this.RcptToHandler);
        }

        private void RcptToHandler(ReceiveCommandEventSource source, RcptCommandEventArgs rcptArgs)
		{
			// Get the recipient address as a lowercase string.
            string strRecipientAddress = rcptArgs.RecipientAddress.ToString().ToLower();

            // For each pair of regexps to email addresses
            foreach (var pair in catchAllConfig.AddressMap)
            {
                // Create the regular expression and the routing address from the dictionary.
                Regex emailPattern = new Regex(pair.Value);
                RoutingAddress emailAddress = new RoutingAddress(pair.Key);

                // If the recipient address matches the regular expression.
                if (emailPattern.IsMatch(strRecipientAddress))
                {
                    // And if the recipient is NOT in the address book.
                    if ((this.addressBook != null) && (this.addressBook.Find(rcptArgs.RecipientAddress) == null))
                    {
                        // Redirect the recipient to the other address.
                        rcptArgs.RecipientAddress = emailAddress;

                        Logger.LogError(String.Format("matched DEA entry ({0} -> {1})", pair.Key, pair.Value));

                        // No further processing.
                        return;
                    }
                }
                else
                {
                    Logger.LogError(String.Format("unmatched DEA entry ({0} -> {1})", pair.Key, pair.Value));
                }
            }
		}		
    }
    
    public class CatchAllConfig
    {
        private static readonly string configFileName = "config.xml";
        private string configDirectory;
        private FileSystemWatcher configFileWatcher;
        private Dictionary<string, string> addressMap;
        private int reLoading = 0;

        private static bool IsValidRegex(string pattern)
        {
            if (string.IsNullOrEmpty(pattern)) return false;

            try
            {
                Regex.Match(String.Empty, pattern);
            }
            catch (ArgumentException)
            {
                return false;
            }

            return true;
        }

        public CatchAllConfig()
        {
            // Setup a file system watcher to monitor the configuration file
            this.configDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            this.configFileWatcher = new FileSystemWatcher(this.configDirectory);
            this.configFileWatcher.NotifyFilter = NotifyFilters.LastWrite;
            this.configFileWatcher.Filter = "config.xml";
            this.configFileWatcher.Changed += new FileSystemEventHandler(this.OnChanged);

            // Create an initially empty map
            this.addressMap = new Dictionary<string, string>();

            // Load the configuration
            this.Load();

            // Now start monitoring
            this.configFileWatcher.EnableRaisingEvents = true;
        }

        public Dictionary<string, string> AddressMap
        {
            get { return this.addressMap; }
        }

        private void OnChanged(object source, FileSystemEventArgs e)
        {
            // Ignore if load ongoing
            if (Interlocked.CompareExchange(ref this.reLoading, 1, 0) != 0)
            {
                Logger.LogError("load ongoing: ignore");
                return;
            }

            // (Re) Load the configuration
            this.Load();

            // Reset the reload indicator
            this.reLoading = 0;
        }
        
        private void Load()
        {
            // Load the configuration
            XmlDocument doc = new XmlDocument();
            bool docLoaded = false;
            string fileName = Path.Combine(this.configDirectory, CatchAllConfig.configFileName);

            try
            {
                doc.Load(fileName);
                docLoaded = true;
            }
            catch (FileNotFoundException)
            {
                Logger.LogError(String.Format("configuration file not found: {0}", fileName));
            }
            catch (XmlException e)
            {
                Logger.LogError(String.Format("XML error: {0}", e.Message));
            }
            catch (IOException e)
            {
                Logger.LogError(String.Format("IO error: {0}", e.Message));
            }

            // If a failure occured, ignore and simply return
            if (!docLoaded || doc.FirstChild == null)
            {
                Logger.LogError("configuration error: either no file or an XML error");
                return;
            }

            // Create a dictionary to hold the mappings
            Dictionary<string, string> map = new Dictionary<string, string>(100);

            // Track whether there are invalid entries
            bool invalidEntries = false;

            // Validate all entries and load into a dictionary
            foreach (XmlNode node in doc.FirstChild.ChildNodes)
            {   
                // We have one node - redirect
                if (string.Compare(node.Name, "redirect", true, CultureInfo.InvariantCulture) != 0)
                {
                    continue;
                }

                XmlAttribute pattern = node.Attributes["pattern"];
                XmlAttribute address = node.Attributes["address"];

                // Validate the data
                if (pattern == null || address == null)
                {
                    invalidEntries = true;
                    Logger.LogError("reject configuration due to an incomplete entry. (Either or both pattern and address missing.)");
                    break;
                }

                if (!IsValidRegex(pattern.Value))
                {
                    invalidEntries = true;
                    Logger.LogError("reject configuration due to an invalid pattern regex.");
                    break;
                }

                if (!RoutingAddress.IsValidAddress(address.Value))
                {
                    invalidEntries = true;
                    Logger.LogError(String.Format("reject configuration due to an invalid address ({0}).", address));
                    break;
                }

                // Add the new entry
                map[address.Value.ToLower()] = pattern.Value;

                Logger.LogError(String.Format("added DEA entry ({0} -> {1}).", pattern.Value, address.Value));
            }

            // If there are no invalid entries, swap in the map
            if (!invalidEntries)
            {
                Interlocked.Exchange<Dictionary<string, string>>(ref this.addressMap, map);
                Logger.LogError("accepted configuration.");
            }
            else
            {
                Logger.LogError("configuration not accepted due to errors.");
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Text;
using Microsoft.Exchange.WebServices;
using Microsoft.Exchange.WebServices.Data;

namespace RoomBook
{
    internal static class ExchangeWebServiceWrapper
    {
        #region Private Members

        private static string mExchangeWebServiceUrl;

        #endregion

        /// <summary>
        /// Initialise the Exchange web service URL
        /// </summary>
        /// <exception cref="ConfigurationErrorsException">Web service URL missing or blank.</exception>
        static ExchangeWebServiceWrapper()
        {
            try
            {
                // Load the Web Service URL from config
                mExchangeWebServiceUrl = ConfigurationManager.AppSettings["ExchangeWebServiceUrl"];

                // It was there, but blank - error
                if (String.IsNullOrEmpty(mExchangeWebServiceUrl))
                {
                    throw new Exception();
                }                
            }
            catch
            {
                // Something went wrong...
                throw new ConfigurationErrorsException("Could not load the Exchange web service URL from the configuration file.");
            }
        }

        /// <summary>
        /// Retrieve an instance of the Exchange web service.
        /// </summary>
        /// <param name="Credentials">Network credentials to connect with.</param>
        /// <returns></returns>
        public static ExchangeService GetWebServiceInstance(NetworkCredential Credentials)
        {
            // Initialise web service instance.
            ExchangeService WebService = new ExchangeService(ExchangeVersion.Exchange2010);

            // Set the credentials to use.
            WebService.Credentials = Credentials;

            // Set the URL to use.
            WebService.Url = new Uri(mExchangeWebServiceUrl);

            // Done.
            return WebService;
        }

        /// <summary>
        /// Retrieve an instance of the Exchange web service.
        /// </summary>
        /// <param name="Domain">Domain of the account to connect with</param>
        /// <param name="Username">Username of the account to connect with</param>
        /// <param name="Password">Password of the account to connect with</param>
        /// <returns></returns>
        public static ExchangeService GetWebServiceInstance(string Username, string Password, string Domain)
        { 
            return GetWebServiceInstance(new NetworkCredential(Username, Password, Domain));
        }

        /// <summary>
        /// Retrieve an instance of the Exchange web service.
        /// </summary>
        /// <returns></returns>
        public static ExchangeService GetWebServiceInstance()
        {
            return GetWebServiceInstance(CredentialCache.DefaultNetworkCredentials);
        }

        /// <summary>
        /// Return the FolderID of a sub-folder within a given parent folder.
        /// </summary>
        /// <param name="WebService">Exchange Web Service instance to use</param>
        /// <param name="BaseFolderId">FolderID of the prent folder</param>
        /// <param name="FolderName">Name of the folder to find</param>
        /// <returns>FolderID of the matching folder (or null if not found)</returns>
        public static FolderId FindFolder(ExchangeService WebService, FolderId BaseFolderId, string FolderName)
        {
            // Retrieve a list of folders (paged by 10)
            FolderView folderView = new FolderView(10, 0);
            // Set base point of offset.
            folderView.OffsetBasePoint = OffsetBasePoint.Beginning;
            // Set the properties that will be loaded on returned items (display namme & folder ID)
            folderView.PropertySet = new PropertySet(FolderSchema.DisplayName, FolderSchema.Id);
            
            FindFoldersResults folderResults;
            do
            {
                // Get the folder at the base level.
                folderResults = WebService.FindFolders(BaseFolderId, folderView);

                // Iterate over the folders, until we find one that matches what we're looking for.
                foreach (Folder folder in folderResults)
                {
                    // Found it - return
                    if (String.Compare(folder.DisplayName, FolderName, StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        return folder.Id;
                    }
                }

                // If there's more folders, get them ...
                if (folderResults.NextPageOffset.HasValue)
                {
                    folderView.Offset = folderResults.NextPageOffset.Value;
                }
            }
            while (folderResults.MoreAvailable);

            // If we got here, we didn't find it, so return null.
            return null;
        }

        /// <summary>
        /// Extract a calendar folder from the specified folder.
        /// </summary>
        /// <param name="WebService">Exchange Web Service instance to use</param>
        /// <param name="RootFolder">The location of the calendar to retrieve.</param>
        /// <param name="FolderPath">Path of the calendar folder within the Public Folder folder to find.</param>
        /// <exception cref="System.Exception">Could not find folder, or folder found is not a calendar.</exception>
        /// <returns>CalenderFolder object corresponding to the requested calendar (or null if not found).</returns>
        public static CalendarFolder GetCalendarFolder(ExchangeService WebService, WellKnownFolderName RootFolder, string FolderPath)
        {
            if (RootFolder != WellKnownFolderName.Calendar && RootFolder != WellKnownFolderName.PublicFoldersRoot)
            {
                throw new ArgumentException("FolderRoot must be either Calendar or PublicFoldersRoot", "FolderRoot");
            }

            // Get the root well-known folder to look in.
            Folder FolderRoot = Folder.Bind(WebService, RootFolder);
            // Split the path of the folder we want to find on "\", and load it into an array (effectively of sub-folders)
            string[] SplitFolderPath = FolderPath.Split('\\');
            // Get the FolderID of the well-known root folder.
            FolderId fId = FolderRoot.Id;

            // For each of the folders in the path, extract it
            foreach (string subFolderName in SplitFolderPath)
            {
                // Find the FolderID of the folder we're looking for.
                fId = FindFolder(WebService, fId, subFolderName);

                // If we couldn't find it, error...
                if (fId == null)
                {
                    throw new Exception(string.Format("Can't find public folder {0}", FolderPath));
                }
            }

            // Get the FolderID of the folder we settled on.
            Folder folderFound = Folder.Bind(WebService, fId);

            // Check the folder we found is a calendar - if not, error
            if (String.Compare(folderFound.FolderClass, "IPF.Appointment", StringComparison.Ordinal) != 0)
            {
                throw new Exception(string.Format("Public folder {0} is not a Calendar", FolderPath));
            }

            // Bind the folder we found as a calendar folder
            CalendarFolder Calendar = CalendarFolder.Bind(WebService, fId, BasePropertySet.FirstClassProperties);

            // Done.
            return Calendar;
        }

        /// <summary>
        /// Extract a calendar folder from the Public Folders folder.
        /// </summary>
        /// <param name="WebService">Exchange Web Service instance to use</param>        
        /// <param name="FolderPath">Path of the calendar folder within the Public Folder folder to find.</param>
        /// <exception cref="System.Exception">Could not find folder, or folder found is not a calendar.</exception>
        /// <returns>CalenderFolder object corresponding to the requested calendar (or null if not found).</returns>
        public static CalendarFolder GetCalendarFolder(ExchangeService WebService, string FolderPath)
        {
            return GetCalendarFolder(WebService, WellKnownFolderName.PublicFoldersRoot, FolderPath);
        }
    }
}

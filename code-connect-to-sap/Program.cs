using System;
using System.Runtime.InteropServices;

namespace code_connect_to_sap
{
    internal class Program
    {
        /// <summary>
        /// Main Test Method
        /// </summary>
        /// <param name="args">Not used</param>
        static void Main(string[] args)
        {
            /* Declare the company variable - the connection */
            SAPbobsCOM.Company company = null;

            Console.WriteLine("Connecting to SAP...");

            /* Connect returns if connection has been established as bool */
            var isConnected = Connect(ref company);

            Console.WriteLine(Environment.NewLine + "Connected; disconnecting now...");

            /* Disconnect also released the held memory */
            Disconnect(ref company);

            Console.WriteLine(Environment.NewLine + "Disconnected. Press any key to exit.");
            Console.ReadKey();
        }

        /// <summary>
        /// Connects to the provided company.
        /// </summary>
        /// <param name="company">Provide uninstantiated</param>
        /// <returns>True if connection was extablished and False if connection could not be done</returns>
        static bool Connect(ref SAPbobsCOM.Company company)
        {
            if (company == null)
            {
                company = new SAPbobsCOM.Company();
            }

            if (!company.Connected)
            {
                /* Server connection details */
                company.Server = "sql-server-name";
                company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016;
                company.DbUserName = "sa";
                company.DbPassword = "sql-password";
                company.UseTrusted = false;

                /* SAP connection details: DB, SAP User and SAP Password */
                company.CompanyDB = "SBODEMOUS";
                company.UserName = "sap-user";
                company.Password = "sap-password";

                /* In case the SAP license server is kept in a different location (in most cases can be left empty) */
                company.LicenseServer = "";

                var isSuccessful = company.Connect() == 0;

                return isSuccessful;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Disconnects and releases the held memory (RAM)
        /// </summary>
        /// <param name="company"></param>
        static void Disconnect(ref SAPbobsCOM.Company company)
        {
            if (company != null
                && company.Connected)
            {
                company.Disconnect();

                Release(ref company);
            }
        }

        /// <summary>
        /// Re-usable method for releasing COM-held memory
        /// </summary>
        /// <typeparam name="T">Type of object to be released</typeparam>
        /// <param name="obj">The instance to be released</param>
        static void Release<T>(ref T obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                }
            }
            catch (Exception) { }
            finally
            {
                obj = default(T);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }
}

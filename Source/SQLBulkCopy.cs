using System;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data;
using System.Data.Odbc;

namespace SQLBulkCopy
{
    /// <summary>
    /// This class uses SQL bulk copy to copy records from a source table, which is accessed either using
    /// OLE DB or ODBC, to a SQL Server destination table.
    /// </summary>
    public class SQLBulkCopy
    {
        /// <summary>
        /// The connection string for the source database.
        /// </summary>
        public string SourceConnectionString;

        /// <summary>
        /// The connection string for the destination database.
        /// </summary>
        public string DestinationConnectionString;

        /// <summary>
        /// Copy data from the source table to the destination table.
        /// </summary>
        /// <param name="sourceTable">
        /// The name of the source table.
        /// </param>
        /// <param name="destinationTable">
        /// The name of the destination table.
        /// </param>
        /// <remarks>
        /// The column positions in the source data reader must match the column positions in 
        /// the destination table.
        /// </remarks>
        public void LoadData(string sourceTable, string destinationTable)
        {
            // If the source connection string contains "driver=" or "dsn=", use ODBC.
            bool useODBC = SourceConnectionString.ToLower().Contains("driver=") ||
                SourceConnectionString.ToLower().Contains("dsn=");

            // Create a connection object for the source database.
            IDbConnection sourceConnection;
            if (useODBC)
            {
                sourceConnection = new OdbcConnection(SourceConnectionString);
            }
            else
            {
                sourceConnection = new OleDbConnection(SourceConnectionString);
            }
            using (sourceConnection)
            {
                // Open the source connection.
                sourceConnection.Open();

                // Create the command objects we'll need, one for getting the record count and one for reading from
                // the source table.
                string countSQL = "select count(*) from " + sourceTable;
                string readSQL = "select * from " + sourceTable;
                IDbCommand commandSourceData;
                if (useODBC)
                {
                    OdbcConnection conn = (OdbcConnection)sourceConnection;
                    commandSourceData = new OdbcCommand(readSQL, conn);
                }
                else
                {
                    OleDbConnection conn = (OleDbConnection)sourceConnection;
                    commandSourceData = new OleDbCommand(readSQL, conn);
                    OleDbCommand cmd = new OleDbCommand("SET DELETED OFF", conn);
                    cmd.ExecuteNonQuery();
                }

                // Get data from the source table as a DataReader.
                IDataReader reader = commandSourceData.ExecuteReader();

                // Set up the bulk copy object. 
                //using (SqlBulkCopy bulkCopy = new SqlBulkCopy(destinationConnection))
                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(DestinationConnectionString, SqlBulkCopyOptions.TableLock))
                {
                    bulkCopy.DestinationTableName = destinationTable;
                    bulkCopy.BulkCopyTimeout = 0;   // no timeout

                    // Copy from the source to the destination and perform a final count on the destination table.
                    try
                    {
                        bulkCopy.WriteToServer(reader);
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                    finally
                    {
                        // Close the DataReader. The SqlBulkCopy and connection objects are automatically closed
                        // at the end of the using block.
                        reader.Close();
                    }
                }
            }
        }

        /// <summary>
        /// Executes a SQL statement.
        /// </summary>
        /// <param name="statement">
        /// The SQL statement to execute.
        /// </param>
        public void ExecuteStatement(string statement)
        {
            using (SqlConnection destinationConnection = new SqlConnection(DestinationConnectionString))
            {
                destinationConnection.Open();
                SqlCommand command = new SqlCommand(statement, destinationConnection);
                try
                {
                    command.ExecuteNonQuery();
                }
                catch (Exception)
                {
                    throw;
                }
            }
        }
    }
}

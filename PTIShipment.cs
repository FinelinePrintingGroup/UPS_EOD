using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using UPS_EODprocessing;
using System.Collections;

namespace UPS_EOD
{
    class PTIShipment
    {
        public string EOD_ID { get; set; }
        public string PTI_LineItemID { get; set; }
        public int numPkgs { get; set; }
        public int numItems { get; set; }
        public decimal charges { get; set; }


        /// <summary>
        /// Get charges for the UPS shipment.
        /// </summary>
        public void GetShipmentCharges()
        {
            string q_getCharges = @"select top 1 ex.totalShipmentcharge
                                        from printable.dbo.CT_UPS_EODexport ex inner join printable.dbo.CT_UPS_EODitems it
                                        on ex.id = it.exportID
                                        where it.PTI_lineItemID  = " + PTI_LineItemID;
            string sqlda_cost = "";

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_logicConnString))
                {
                    SqlCommand command = new SqlCommand(q_getCharges, conn);
                    try
                    {
                        command.Connection.Open();
                        try
                        {
                            sqlda_cost = command.ExecuteScalar().ToString();
                        }
                        catch (Exception e)
                        {
                            sqlda_cost = "0";
                        }

                        Console.WriteLine(sqlda_cost);

                        conn.Close();

                        if ((sqlda_cost == "") || (sqlda_cost == null))
                        {
                            //If no charges exist, make them 0
                            charges = 0;
                        }
                        else
                        {
                            //If charges exist, convert them to decimal and round.
                            charges = Convert.ToDecimal(sqlda_cost);
                        }
                    }
                    catch (Exception e)
                    {
                        errorLog("9", "calculateItemCharges", e.ToString().Substring(0, 250));
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("10", "calculateItemCharges", e.ToString().Substring(0, 250));
                Console.WriteLine(e.ToString());
            }
        }

        /// <summary>
        /// Get per item UPS charge.
        /// </summary>
        /// <returns>Item charge</returns>
        public decimal GetChargesPerItem()
        {
            charges = charges / getPTIlines();
            charges = Math.Round(charges, 2);
            return charges;
        }

        /// <summary>
        /// Gets the number of items on the shipment
        /// </summary>
        /// <returns></returns>
        public int getPTIlines()
        {
            int numLines = 0;

            string q = @"SELECT COUNT(*)
                         FROM pLogic.dbo.ShipmentItems
                         WHERE ShipmentNumber = (SELECT TOP 1 ShipmentNum
                                                 FROM CT_UPS_EODitems
                                                 WHERE id = " + EOD_ID + @")";

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        numLines = Convert.ToInt32(command.ExecuteScalar());

                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                        errorLog("getPTInum-1", e.ToString(), "Get PTI #lines Error");
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                errorLog("getPTInum-2", e.ToString(), "Get PTI #lines Error");
            }

            return numLines;
        }

        /// <summary>
        /// Check if Cummins customer for free shipping
        /// </summary>
        /// <returns></returns>
        public bool freeShippingByZip_PTI()
        {
            string zip = "";
            ArrayList al = new ArrayList();

            string queryZip = @"select zip
                                from vw_UPS_EODitemsDetail  
                                where ShipmentNumber = (select shipmentNum
                                                        from CT_UPS_EODitems
                                                        where id = " + EOD_ID + @")
                                and LineN = (select shipmentLineNum
                                                        from CT_UPS_EODitems
                                                        where id = " + EOD_ID + ")";
            //RETRIEVE ZIP CODE FROM THE VIEW
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(queryZip, conn);
                    try
                    {
                        command.Connection.Open();
                        SqlDataReader reader = command.ExecuteReader();

                        while (reader.Read())
                        {
                            object[] values = new object[reader.FieldCount];
                            reader.GetValues(values);
                            al.Add(values);
                        }

                        reader.Close();
                        conn.Close();

                        foreach (object[] row in al)
                        {
                            foreach (object column in row)
                            {
                                zip = column.ToString();
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        errorLog("26", "freeShippingByZip", e.ToString().Substring(0, 250));
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("27", "freeShippingByZip", e.ToString().Substring(0, 250));
                Console.WriteLine(e.ToString());
            }

            //SEE IF ZIP CODE MEETS REQUIREMENTS
            //// Cummins Columbus - 4720_
            bool zipMatch = false;

            if (String.Compare(zip, 0, "4720", 0, 4) == 0)
                zipMatch = true;
            else if (String.Compare(zip, 0, "46282", 0, 5) == 0)
                zipMatch = true;
            else
                zipMatch = false;

            return zipMatch;
        }

        public int getJobN()
        {
            int jobN = 0;

            string q = @"SELECT TOP 1 Logic_jobNum
                         FROM printable.dbo.CT_UPS_EODitems
                         WHERE PTI_LineItemID = " + PTI_LineItemID;

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();

                        jobN = Convert.ToInt32(command.ExecuteScalar() ?? 0);

                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        errorLog("XX", "getJobN", e.ToString().Substring(0, 250));
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("XX", "getJobN", e.ToString().Substring(0, 250));
                Console.WriteLine(e.ToString());
            }

            return jobN;
        }

        public int getFGNum()
        {
            int fgNum = 0;

            string q = @"SELECT FGorder
                        FROM printable.dbo.vw_UPS_EoditemsDetail
                        WHERE comments = '" + PTI_LineItemID + "'";

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();

                        fgNum = Convert.ToInt32(command.ExecuteScalar() ?? 0);

                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        errorLog("XX", "getJobN", e.ToString().Substring(0, 250));
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("XX", "getJobN", e.ToString().Substring(0, 250));
                Console.WriteLine(e.ToString());
            }

            return fgNum;
        }

        public decimal getPrevCharges()
        {
            decimal prevCost = 0;
            ArrayList al = new ArrayList();

            string getPrevQuery = @"select ISNULL(ActualCharge,0)
                                    from ShipmentItems
                                    where ShipmentNumber = (select shipmentNum
                                                            from printable.dbo.CT_UPS_EODitems
                                                            where id = " + EOD_ID + @")
                                    and LineN = (select shipmentLineNum
                                                            from printable.dbo.CT_UPS_EODitems
                                                            where id = " + EOD_ID + ")";
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_logicConnString))
                {
                    SqlCommand command = new SqlCommand(getPrevQuery, conn);
                    try
                    {
                        command.Connection.Open();
                        SqlDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            object[] values = new object[reader.FieldCount];
                            reader.GetValues(values);
                            al.Add(values);
                        }

                        reader.Close();
                        conn.Close();

                        foreach (object[] row in al)
                        {
                            foreach (object column in row)
                            {
                                prevCost = Convert.ToDecimal(column.ToString());
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        errorLog("24", "getPrevCharges", e.ToString().Substring(0, 250));
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("25", "getPrevCharges", e.ToString().Substring(0, 250));
                Console.WriteLine(e.ToString());
            }

            return prevCost;
        }

        public void UpdateItemCharges()
        {
            int jobN = getJobN();
            int FGNum = getFGNum();
            decimal cost = GetChargesPerItem();
            decimal prevCost = 0.00M;
            decimal totalCost = 0.00M;

            prevCost = getPrevCharges();
            totalCost = prevCost + cost;

            string updateQuery = @"UPDATE CT_UPS_EODitems
                                   SET itemCharges = '" + cost + @"'
                                   WHERE PTI_lineItemID = (SELECT PTI_lineItemID
                                                           FROM printable.dbo.CT_UPS_EODitems
                                                           WHERE id = " + EOD_ID + @")";
            //UPDATE EODitems TABLE
            Console.WriteLine("Updating EODitems Table");
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(updateQuery, conn);
                    try
                    {
                        int rowsUpdated = 0;
                        command.Connection.Open();
                        rowsUpdated = command.ExecuteNonQuery();
                        Console.WriteLine("EODitems Table Updated");
                        command.Dispose();
                        command = null;
                    }
                    catch (Exception e)
                    {
                        errorLog("14", "updateItemCharges", e.ToString().Substring(0, 250));
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("15", "updateItemCharges", e.ToString().Substring(0, 250));
                Console.WriteLine(e.ToString());
            }

            if (jobN != 0)
            {
                Console.WriteLine("Updating JOB costing...");
                //UPDATE ShipmentItems TABLE via JOBN
                string updateQuery2 = @"UPDATE ShipmentItems
                                        SET ActualCharge = " + cost + @"
                                        WHERE MainReference = " + jobN;
                try
                {
                    using (SqlConnection conn = new SqlConnection(Globals.get_logicConnString))
                    {
                        SqlCommand command = new SqlCommand(updateQuery2, conn);
                        try
                        {
                            int rowsUpdated = 0;
                            command.Connection.Open();
                            rowsUpdated = command.ExecuteNonQuery();
                            Console.WriteLine("ShipmentItems Table Updated\n");
                            command.Dispose();
                            command = null;
                            errorLog("CHRG-" + EOD_ID, "ItemCharges updated for id:" + EOD_ID, "Prev: $" + prevCost + ", New: $" + totalCost);
                        }
                        catch (Exception e)
                        {
                            errorLog("16", "updateItemCharges", e.ToString().Substring(0, 250));
                            Console.WriteLine(e.ToString());
                        }
                    }
                }
                catch (Exception e)
                {
                    errorLog("17", "updateItemCharges", e.ToString().Substring(0, 250));
                    Console.WriteLine(e.ToString());
                }
            }
            else if (FGNum != 0)
            {
                Console.WriteLine("Updating FG Order costing...");
                //UPDATE ShipmentItems TABLE via FGNum
                string updateQuery2 = @"update plogic.dbo.shipmentitems
                                        set actualcharge = " + cost + @"
                                        where shipmentNumber = (select shipmentnumber
						                                         from plogic.dbo.shipments
						                                         where FGOrder = " + FGNum + ")";
                try
                {
                    using (SqlConnection conn = new SqlConnection(Globals.get_logicConnString))
                    {
                        SqlCommand command = new SqlCommand(updateQuery2, conn);
                        try
                        {
                            int rowsUpdated = 0;
                            command.Connection.Open();
                            rowsUpdated = command.ExecuteNonQuery();
                            Console.WriteLine("ShipmentItems Table Updated\n");
                            command.Dispose();
                            command = null;
                            errorLog("CHRG-FG" + EOD_ID, "ItemCharges updated for id:" + EOD_ID, "Prev: $" + prevCost + ", New: $" + cost );
                        }
                        catch (Exception e)
                        {
                            errorLog("16", "updateItemCharges", e.ToString().Substring(0, 250));
                            Console.WriteLine(e.ToString());
                        }
                    }
                }
                catch (Exception e)
                {
                    errorLog("17", "updateItemCharges", e.ToString().Substring(0, 250));
                    Console.WriteLine(e.ToString());
                }
            }
        }

        //Error logging function
        public void errorLog(string pErrNum, string pErrDesc, string pErrMsg)
        {
            string errNum = pErrNum;
            string errDesc = pErrDesc;
            string errMsg = pErrMsg;
            string errDate = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss");

            string queryString = @"INSERT INTO CT_EOD_errorLog (errNum, errDesc, errMsg, errDate)
                                    VALUES ('" + errNum + "','" + errDesc + "','" + errMsg + "','" + errDate + "')";

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(queryString, conn);
                    try
                    {
                        int rowsInserted = 0;
                        command.Connection.Open();
                        rowsInserted = command.ExecuteNonQuery();
                        command.Dispose();
                        command = null;
                        Console.WriteLine(rowsInserted.ToString() + " error logged - Error #" + errNum);
                    }
                    catch (Exception e)
                    {
                        errorLog("28", "errorLog", e.ToString().Substring(0, 800));
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("27", "errorLog", e.ToString().Substring(0, 250));
                Console.WriteLine(e.ToString());
            }
        }
    }
     
}

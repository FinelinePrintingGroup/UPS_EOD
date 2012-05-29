using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;
using System.Net;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Collections;
using System.Net.Mail;

namespace UPS_EODprocessing
{
    class methods
    {
        //Function that returns an ArrayList of IDs that need to be processed. They're EOD items will be added for processing.
        public ArrayList getShipmentsForProcessing()
        {
            Console.WriteLine("");
            Console.WriteLine("=========================");
            Console.WriteLine("=   GETTING SHIPMENTS   =");
            Console.WriteLine("=========================");

            string queryString = @"SELECT id
                                   FROM CT_UPS_EODexport
                                   WHERE EOD_processed = 0
                                   AND ISNUMERIC(shipmentNum) = 1";

            ArrayList al = new ArrayList();

            //Querying the DB
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(queryString, conn);
                    try
                    {
                        command.Connection.Open();
                        SqlDataReader reader = command.ExecuteReader();

                        //Populating the ArrayList
                        while (reader.Read())
                        {
                            object[] values = new object[reader.FieldCount];
                            reader.GetValues(values);
                            al.Add(values);
                        }

                        reader.Close();
                        conn.Close();

                        return al;
                    }
                    catch (Exception e)
                    {
                        errorLog("1", "getShipmentsForProcessing", e.ToString().Substring(0,250));
                        Console.WriteLine(e.ToString());
                        return al;
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("2", "getShipmentsForProcessing", e.ToString().Substring(0,250));
                Console.WriteLine(e.ToString());
                return al;
            }
            
        }

        // Function that gets internal shipment#'s that were shipped UPS but never reached the EOD export table. 
        // This can be caused by FG orders that were shipped in the same package as jobs but weren't able to be 
        // combined into the same shipment.
        public ArrayList getMissedUPSShipments()
        {
            Console.WriteLine("");
            Console.WriteLine("=========================");
            Console.WriteLine("=    GETTING MISSED     =");
            Console.WriteLine("=========================");

            string queryString = @"SELECT ShipmentNumber
                                   FROM vw_UPS_EODmissedUPS
                                   WHERE ShipDate > GETDATE()-7
                                   AND ShipDate < GETDATE()";

            ArrayList al = new ArrayList();

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(queryString, conn);
                    try
                    {
                        command.Connection.Open();
                        SqlDataReader reader = command.ExecuteReader();

                        while (reader.Read())
                        {
                            al.Add(Convert.ToInt32(reader[0].ToString()));
                        }

                        reader.Close();
                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        errorLog("1", "getMissedUPSShipments", e.ToString().Substring(0, 250));
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("2", "getMissedUPSShipments", e.ToString().Substring(0, 250));
                Console.WriteLine(e.ToString());
            }

            return al;
        }

        //Function that processes ID's with no Shipment Numbers
        public void processNoShipNum()
        {
            string queryString = @"select id
                                   from CT_UPS_EODexport
                                   where EOD_processed = 0
                                   and shipmentNum is NULL
                                   or shipmentNum = '0'";

            ArrayList al = new ArrayList();

            //Querying the DB
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(queryString, conn);
                    try
                    {
                        command.Connection.Open();
                        SqlDataReader reader = command.ExecuteReader();

                        //Populating the ArrayList
                        while (reader.Read())
                        {
                            object[] values = new object[reader.FieldCount];
                            reader.GetValues(values);
                            al.Add(values);
                        }

                        reader.Close();
                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        errorLog("X", "processNoShipNum", e.ToString().Substring(0, 250));
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("X", "processNoShipNum", e.ToString().Substring(0, 250));
                Console.WriteLine(e.ToString());
            }

            Console.WriteLine("");
            Console.WriteLine("=========================");
            Console.WriteLine("= NULL SHIPMENT NUMBERS =");
            Console.WriteLine("=========================");
            

            foreach (object[] row in al)
            {
                foreach (object column in row)
                {
                    Console.Write(column.ToString() + ", ");
                    errorLog("9999", "No shipment#", "No shipment number for ID:" + column.ToString());
                    processEOD_EODexport(column.ToString());
                }
            }

        }

        //Function that processes non-numeric shipment numbers when Shipping puts bad data in WorldShip
        public void processBadShipNum()
        {
            string queryString = @"SELECT shipmentNum
                                   FROM CT_UPS_EODexport
                                   WHERE EOD_processed = 0
                                   AND shipmentNum IS NOT NULL
                                   AND ISNUMERIC(ShipmentNum) = 0";

            ArrayList al = new ArrayList();

            //Querying the DB
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(queryString, conn);
                    try
                    {
                        command.Connection.Open();
                        SqlDataReader reader = command.ExecuteReader();

                        //Populating the ArrayList
                        while (reader.Read())
                        {
                            object[] values = new object[reader.FieldCount];
                            reader.GetValues(values);
                            al.Add(values);
                        }

                        reader.Close();
                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        errorLog("X", "processBadShipNum", e.ToString().Substring(0, 250));
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("X", "processBadShipNum", e.ToString().Substring(0, 250));
                Console.WriteLine(e.ToString());
            }

            Console.WriteLine("");
            Console.WriteLine("========================");
            Console.WriteLine("= BAD SHIPMENT NUMBERS =");
            Console.WriteLine("========================");


            foreach (object[] row in al)
            {
                foreach (object column in row)
                {
                    sendEmail("ShawnH@FineLink.com", column.ToString() + " - Bad Shipment Number");
                    errorLog("9999", "No shipment#", "Bad shipment number: " + column.ToString());
                }
            }

        }

        //Function that returns an ArrayList of EOD IDs that need to be processed
        public ArrayList getShipmentItemsForProcessing()
        {
            Console.WriteLine("");
            Console.WriteLine("==========================");
            Console.WriteLine("= GETTING SHIPMENT ITEMS =");
            Console.WriteLine("==========================");

            string queryString = @"select id
                                   from CT_UPS_EODitems
                                   where EOD_processed = 0";

            ArrayList al = new ArrayList();

            //Querying the DB
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(queryString, conn);
                    try
                    {
                        command.Connection.Open();
                        SqlDataReader reader = command.ExecuteReader();

                        //Populating the DB
                        while (reader.Read())
                        {
                            object[] values = new object[reader.FieldCount];
                            reader.GetValues(values);
                            al.Add(values);
                        }

                        reader.Close();
                        conn.Close();

                        return al;
                    }
                    catch (Exception e)
                    {
                        errorLog("3", "getShipmentItemsForProcessing", e.ToString().Substring(0,250));
                        Console.WriteLine(e.ToString());
                        return al;
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("4", "getShipmentItemsForProcessing", e.ToString().Substring(0,250));
                Console.WriteLine(e.ToString());
                return al;
            }
        }

        //Takes a shipment ID and populates the EODitems table
        public void insertEODitems(string pID)
        {
            Console.WriteLine("Inserting EOD Items for ID:" + pID);

            string id = pID;

            // There was a big issue with this query and view originally because FGIOrderDetl.Comments was chosen to contain the OrderDetails.OrderDetail_ID but it also was capable of 
            // containing garbage that CSRs entered. This caused a hangup because of the nonnumeric data messing with things. I pushed all the OD_ID's over to FGIOrderDetl.UserItemDesc
            // and changed the view to look there instead (under the alias of comments, to prevent breaking anything that relied on it). I also changed the incoming PHP so that it     
            // writes the OD_ID to both Comments and UserItemDesc, just in case.                                                                                                        
            string queryString = @"INSERT INTO CT_UPS_EODitems (exportID, shipmentNum, shipmentLineNum, Logic_jobNum, Logic_FGNum, PTI_lineItemID, itemCharges)
                                   SELECT DISTINCT '" + id + @"', ShipmentNumber, LineN, MainReference, FGItemNum, ISNULL(OrderDetail_ID, Comments), ISNULL(Price_Cost_Shipping, 0)
	                               FROM vw_UPS_EODitemsDetail
                                   WHERE ShipmentNumber = (SELECT shipmentNum
                                                           FROM CT_UPS_EODexport
                                                           WHERE id = " + id + @")";
            
            //Inserting into the DB
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
                    }
                    catch (Exception e)
                    {
                        errorLog("5", "insertEODitems", e.ToString().Substring(0,250));
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("6", "insertEODitems", e.ToString().Substring(0,250));
                Console.WriteLine(e.ToString());
            }
        }

        public void insertEODitems_nonUPS()
        {
            Console.WriteLine("");
            Console.WriteLine("=========================================");
            Console.WriteLine("= Inserting EOD Items for NON-EOD Items =");
            Console.WriteLine("=========================================");

            string shipNum = "";
            string lineNum = "";
            ArrayList list = new ArrayList();
            list.Clear();

            // WHERE CLAUSES ARE USED TO STRIP OUT OLD UNUSED DATA
            string querySelect = @"select distinct ShipmentNumber, shipmentLineNum
                                   from printable.dbo.vw_PTI_integration
                                   where not exists (select shipmentNum, shipmentLineNum
                                                     from printable.dbo.CT_UPS_EODitems
                                                     where printable.dbo.CT_UPS_EODitems.shipmentNum = printable.dbo.vw_PTI_integration.ShipmentNumber)
                                   and CarrierName != 'UPS'
                                   and isnumeric(OrderDetail_ID) = 1
                                   and OrderDetail_ID is not null
                                   and OrderDetail_ID not like '$%'
                                   and ShipmentStatus = 1
                                   order by ShipmentNumber, shipmentLineNum";

            //GET SHIPNUM AND LINENUM
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(querySelect, conn);
                    try
                    {
                        command.Connection.Open();
                        SqlDataReader reader = command.ExecuteReader();

                        while (reader.Read())
                        {
                            object[] values = new object[reader.FieldCount];
                            reader.GetValues(values);
                            list.Add(values);
                        }

                        reader.Close();
                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            foreach (object[] row in list)
            {
                
                Console.WriteLine("   -" + row[0].ToString());
                Console.WriteLine("   -" + row[1].ToString());
                
                shipNum = row[0].ToString();
                lineNum = row[1].ToString();

                string queryString = @"insert into printable.dbo.CT_UPS_EODitems (exportID, shipmentNum, shipmentLineNum, Logic_jobNum, Logic_FGNum, PTI_lineItemID, itemCharges)
                                       select top 1 0, ShipmentNumber, LineN, MainReference, FGItemNum, isnull(OrderDetail_ID, Comments) as ID, isnull(Price_Cost_Shipping,0) as cost
	                                   from printable.dbo.vw_UPS_EODitemsDetail
                                       where ShipmentNumber = " + shipNum + @"
                                       and LineN = " + lineNum + @"
                                       and (Comments not in (select PTI_lineItemID from CT_UPS_EODitems where ShipmentNum = " + shipNum + @")
                                            or
                                            OrderDetail_ID not in (select PTI_lineItemID from CT_UPS_EODitems where ShipmentNum = " + shipNum + @"))";
                //INSERT EOD ITEMS
                try
                {
                    using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                    {
                        SqlCommand command = new SqlCommand(queryString, conn);
                        try
                        {
                            Console.WriteLine("Inserting");
                            int rowsInserted = 0;
                            command.Connection.Open();
                            rowsInserted = command.ExecuteNonQuery();
                            command.Dispose();
                            command = null;
                        }
                        catch (Exception e)
                        {
                            errorLog("5", "insertEODitemsNONUPSins", e.ToString().Substring(0, 250));
                            Console.WriteLine(e.ToString());
                        }
                    }
                }
                catch (Exception e)
                {
                    errorLog("6", "insertEODitemsNONUPSins", e.ToString().Substring(0, 250));
                    Console.WriteLine(e.ToString());
                }

                //UPDATE CHARGES ACROSS LOGIC
                string updateQuery2 = @"update pLogic.dbo.ShipmentItems
                                        set ActualCharge = (select top 1 isnull(Price_Cost_Shipping, 0)
                                                            from printable.dbo.vw_UPS_EODitemsDetail
                                                            where ShipmentNumber = " + shipNum + @"
                                                            and LineN = " + lineNum + @"),
                                            EstimatedChrg = (select top 1 isnull(Price_Cost_Shipping, 0)
                                                            from printable.dbo.vw_UPS_EODitemsDetail
                                                            where ShipmentNumber = " + shipNum + @"
                                                            and LineN = " + lineNum + @")
                                        where ShipmentNumber = " + shipNum + @"
                                        and LineN = " + lineNum;

                try
                {
                    using (SqlConnection conn = new SqlConnection(Globals.get_logicConnString))
                    {
                        SqlCommand command = new SqlCommand(updateQuery2, conn);
                        try
                        {
                            Console.WriteLine("Updating");
                            int rowsUpdated = 0;
                            command.Connection.Open();
                            rowsUpdated = command.ExecuteNonQuery();
                            Console.WriteLine(rowsUpdated.ToString() + " SI rows updated");
                            command.Dispose();
                            command = null;
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

                //PROCESS EOD ITEMS
                string updateQuery = @"update printable.dbo.CT_UPS_EODitems
                                       set EOD_processed = 1
                                       where ShipmentNum = '" + shipNum + @"'
                                       and shipmentLineNum = '" + lineNum + "'";

                try
                {
                    using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                    {
                        SqlCommand command = new SqlCommand(updateQuery, conn);
                        try
                        {
                            int rowsUpdated2 = 0;
                            command.Connection.Open();
                            rowsUpdated2 = command.ExecuteNonQuery();
                            Console.WriteLine(rowsUpdated2.ToString() + " EOD rows updated");
                            command.Dispose();
                            command = null;
                        }
                        catch (Exception e)
                        {
                            errorLog("18", "processEOD_EODitems", e.ToString().Substring(0, 250));
                            Console.WriteLine(e.ToString());
                        }
                    }
                }
                catch (Exception e)
                {
                    errorLog("19", "processEOD_EODitems", e.ToString().Substring(0, 250));
                    Console.WriteLine(e.ToString());
                }
                
            }
        }

        public void insertEODitems_UPSmissedEOD(int shipmentNumber)
        {
            Console.WriteLine("Inserting EOD Items for Missed UPS Shipment:" + shipmentNumber);

            string queryString = @"INSERT INTO CT_UPS_EODitems (exportID, shipmentNum, shipmentLineNum, Logic_jobNum, Logic_FGNum, PTI_lineItemID, itemCharges, EOD_processed)
                                    SELECT DISTINCT '0', ShipmentNumber, LineN, MainReference, FGItemNum, ISNULL(OrderDetail_ID, Comments), ISNULL(Price_Cost_Shipping, 0), 1
	                                FROM vw_UPS_EODitemsDetail
                                    WHERE ShipmentNumber = " + shipmentNumber;

            //Inserting into the DB
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
                    }
                    catch (Exception e)
                    {
                        errorLog("1", "insertEODitems_UPSmissedEOD", e.ToString().Substring(0, 250));
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("2", "insertEODitems_UPSmissedEOD", e.ToString().Substring(0, 250));
                Console.WriteLine(e.ToString());
            }
        }

        //Function that calculates the EOD item charges per EOD ID.
        public void calculateItemCharges(string pID)
        {
            string id = pID;
            string pti_ID = "0";
            string exportID = getExportID(id);
            decimal charges = 0;
            decimal chargesTotal = 0;
            decimal numPkgs = 0;

            string queryString = @"select PTI_lineItemID
                                   from printable.dbo.CT_UPS_EODitems
                                   where id = " + id;

            ArrayList al = new ArrayList();

            //Looks for a PTI ID so to process it as a Logic job or Printable Job
            #region Accessing PTI ID
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(queryString, conn);
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
                                if ((column.ToString() == "") || (column.ToString() == null))
                                {
                                    pti_ID = "0";
                                }
                                else
                                {
                                    pti_ID = column.ToString();
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        errorLog("7", "calculateItemCharges", e.ToString().Substring(0,250));
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("8", "calculateItemCharges", e.ToString().Substring(0,250));
            }
            #endregion

            #region Printable Cost Area
            if (Convert.ToDecimal(pti_ID) > 0)
            {
                Console.WriteLine("Processing ShipmentItem ID:" + id + " as a Printable order");
                Console.WriteLine("OrderDetail_ID:" + pti_ID);
                string sqlda_cost = "";

                string q_getCharges = @"SELECT TOP 1 ISNULL(Price_Cost_Shipping, 0)
                                        FROM printable.dbo.OrderDetails
                                        WHERE OrderDetail_ID = " + pti_ID;

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
                                
                                charges = charges / getPTIpkgs(pti_ID, id);
                                charges = charges / getPTIlines(pti_ID, id);
                                charges = Math.Round(charges, 2);
                            }
                        }
                        catch (Exception e)
                        {                           
                            errorLog("9", "calculateItemCharges", e.ToString().Substring(0,250));
                            Console.WriteLine(e.ToString());                            
                        }
                    }
                }
                catch (Exception e)
                {
                    errorLog("10", "calculateItemCharges", e.ToString().Substring(0,250));
                    Console.WriteLine(e.ToString());
                }
                        

                //Update charges on CT_UPS_EODitems table and possibly ShipmentItems

                if (freeShippingByZip_PTI(id)) //Free shipping by zip code check!!!
                {
                    charges = 0.00M;
                    Console.WriteLine("Cummins - Columbus PTI Order: No charge shipping");
                    errorLog("NOCHRG", "No charge Cummins", pti_ID + " - No charge shipping $0 - " + id);
                    updateItemCharges(0.00M, id);
                    updateItemCharges_PTI(0.00M, id, pti_ID);
                    updateItemCharges_OrderDetails(0.00M, pti_ID);
                }

                Console.WriteLine("Charges: $" + charges);

                if (Convert.ToDecimal(pti_ID) > 0)
                {
                }
                else
                {
                    updateItemCharges(charges, id);
                }
            }
            #endregion

            #region Logic cost area
            else
            {
                Console.WriteLine("Processing ShipmentItem ID:" + id + " as a Logic order");

                string queryCharges = @"select ex.shipmentCharges
                                        from CT_UPS_EODexport ex
                                            left outer join CT_UPS_EODitems it
                                                on ex.id = it.exportID
                                        where it.id = " + id;

                string queryNumPackages = @"select isnull(count(id),0)
                                            from CT_UPS_EODitems
                                            where shipmentNum = (select shipmentNum
                                                                 from CT_UPS_EODitems
                                                                 where id = " + id + @"
                                                                 and exportID = " + exportID + @")
                                            and exportID = " + exportID;

                try
                {
                    using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                    {
                        SqlCommand command = new SqlCommand(queryCharges, conn); //get the charges command
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
                                    if ((column.ToString() == "") || (column.ToString() == null))
                                    {
                                        charges = 0;
                                    }
                                    else
                                    {
                                        charges = Convert.ToDecimal(column.ToString());
                                        charges = Math.Round(charges, 2);
                                    }
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            errorLog("11", "calculateItemCharges", e.ToString().Substring(0,250));
                            Console.WriteLine(e.ToString());
                        }
                    }

                    using (SqlConnection conn2 = new SqlConnection(Globals.get_printableConnString))
                    {
                        SqlCommand command2 = new SqlCommand(queryNumPackages, conn2); //get the number of packages command
                        try
                        {
                            command2.Connection.Open();
                            SqlDataReader reader2 = command2.ExecuteReader();

                            while (reader2.Read())
                            {
                                object[] values2 = new object[reader2.FieldCount];
                                reader2.GetValues(values2);
                                al.Add(values2);
                            }

                            reader2.Close();
                            conn2.Close();

                            foreach (object[] row in al)
                            {
                                foreach (object column in row)
                                {
                                    if ((column.ToString() == "") || (column.ToString() == null))
                                    {
                                        numPkgs = 0;
                                    }
                                    else
                                    {
                                        numPkgs = Convert.ToDecimal(column.ToString());
                                    }
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            errorLog("12", "calculateItemCharges", e.ToString().Substring(0,250));
                            Console.WriteLine(e.ToString());
                        }
                    }

                    //Calculates "per package" cost for multipackage shipments
                    //double addSomeMore = Convert.ToDouble(charges) * 0.2;
                    //chargesTotal = charges + Convert.ToDecimal(addSomeMore);

                    chargesTotal = charges;
                    if (numPkgs > 0)
                    {
                        charges = chargesTotal / numPkgs;
                    }


                    Console.WriteLine("Packages in shipment: " + numPkgs);
                    Console.WriteLine("Charges for shipment: $" + chargesTotal);
            
                }
                catch (Exception e)
                {
                    errorLog("13", "calculateItemCharges", e.ToString().Substring(0,250));
                    Console.WriteLine(e.ToString());
                }

                //Update charges on CT_UPS_EODitems table and possibly ShipmentItems

                Console.WriteLine("Charges for item: $" + charges);
                updateItemCharges(charges, id);
            }
            #endregion
        }

        private int getPTIpkgs(string ptiID, string id)
        {
            int numPkgs = 0;

//            string q = @"SELECT COUNT(*)
//                         FROM CT_UPS_EODitems
//                         WHERE PTI_LineItemID = " + ptiID + @"
//                         AND ShipmentNum = (SELECT TOP 1 ShipmentNum
//                                            FROM CT_UPS_EODitems
//                                            WHERE id = " + id + @")";

            string q = @"SELECT COUNT(*)
                         FROM CT_UPS_EODitems
                         WHERE PTI_LineItemID = " + ptiID;

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        numPkgs = Convert.ToInt32(command.ExecuteScalar());

                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                        errorLog("getPTInum-1", e.ToString(), "Get PTI #pkgs Error");
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                errorLog("getPTInum-2", e.ToString(), "Get PTI #pkgs Error");
            }

            return numPkgs;
        }

        private int getPTIlines(string ptiID, string id)
        {
            int numLines = 0;

            string q = @"SELECT COUNT(*)
                         FROM pLogic.dbo.ShipmentItems
                         WHERE ShipmentNumber = (SELECT TOP 1 ShipmentNum
                                                 FROM CT_UPS_EODitems
                                                 WHERE id = " + id + @")";

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

        //Function that updates the Item Charges per EOD ID.
        public void updateItemCharges(decimal pCost, string pID)
        {
            decimal cost = pCost;
            decimal totalCost = 0;
            decimal prevCost = 0;
            string id = pID;
            ArrayList al = new ArrayList();

            prevCost = getPrevCharges(id);
            totalCost = prevCost + cost;

            string updateQuery = @"update CT_UPS_EODitems
                                   set itemCharges = '" + cost + @"'
                                   where id = " + id;

//            string updateQuery2 = @"update ShipmentItems
//                                    set ActualCharge = " + totalCost + @",
//                                        EstimatedChrg = " + totalCost + @"
//                                    where ShipmentNumber = (select shipmentNum
//                                                            from printable.dbo.CT_UPS_EODitems
//                                                            where id = " + id + @")
//                                    and LineN = (select shipmentLineNum
//                                                            from printable.dbo.CT_UPS_EODitems
//                                                            where id = " + id + ")";

            string updateQuery2 = @"update ShipmentItems
                                    set ActualCharge = " + totalCost + @"
                                    where ShipmentNumber = (select shipmentNum
                                                            from printable.dbo.CT_UPS_EODitems
                                                            where id = " + id + @")
                                    and LineN = (select shipmentLineNum
                                                            from printable.dbo.CT_UPS_EODitems
                                                            where id = " + id + ")";

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
                        errorLog("14", "updateItemCharges", e.ToString().Substring(0,250));
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("15", "updateItemCharges", e.ToString().Substring(0,250));
                Console.WriteLine(e.ToString());
            }

            //UPDATE ShipmentItems TABLE
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
                        errorLog("CHRG-"+id, "ItemCharges updated for id:" + id, "Prev: $" + prevCost + ", New: $" + totalCost);
                    }
                    catch (Exception e)
                    {
                        errorLog("16", "updateItemCharges", e.ToString().Substring(0,250));
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("17", "updateItemCharges", e.ToString().Substring(0,250));
                Console.WriteLine(e.ToString());
            }
        }

        public void updateItemCharges_PTI(decimal pCost, string pID, string PTI_id)
        {
            decimal cost = pCost;
            decimal totalCost = 0;
            decimal prevCost = 0;
            string id = pID;
            ArrayList al = new ArrayList();

            int jobN = getJobN(Convert.ToInt32(PTI_id));

            prevCost = getPrevCharges(id);
            totalCost = prevCost + cost;

            string updateQuery = @"UPDATE CT_UPS_EODitems
                                   SET itemCharges = '" + cost + @"'
                                   WHERE PTI_lineItemID = (SELECT PTI_lineItemID
                                                           FROM printable.dbo.CT_UPS_EODitems
                                                           WHERE id = " + pID + @")";

            //            string updateQuery2 = @"update ShipmentItems
            //                                    set ActualCharge = " + totalCost + @",
            //                                        EstimatedChrg = " + totalCost + @"
            //                                    where ShipmentNumber = (select shipmentNum
            //                                                            from printable.dbo.CT_UPS_EODitems
            //                                                            where id = " + id + @")
            //                                    and LineN = (select shipmentLineNum
            //                                                            from printable.dbo.CT_UPS_EODitems
            //                                                            where id = " + id + ")";

            string updateQuery2 = @"UPDATE ShipmentItems
                                    SET ActualCharge = " + totalCost + @"
                                    WHERE MainReference = " + jobN;

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

            //UPDATE ShipmentItems TABLE
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
                        errorLog("CHRG-" + id, "ItemCharges updated for id:" + id, "Prev: $" + prevCost + ", New: $" + totalCost);
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

        public void updateItemCharges_OrderDetails(decimal pCost, string PTI_id)
        {
            string q = @"UPDATE printable.dbo.OrderDetails
                         SET Price_Cost_Shipping = '" + pCost + @"'
                         WHERE OrderDetail_ID = " + PTI_id;

            Console.WriteLine("Updating OrderDetails Table");
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        int rowsUpdated = 0;
                        command.Connection.Open();
                        rowsUpdated = command.ExecuteNonQuery();
                        Console.WriteLine("OrderDetails Table Updated");
                        command.Dispose();
                        command = null;
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

        //Changes the EODitems_processed from 0 to 1
        public void processEOD_EODitems(string pID)
        {
            string id = pID;

            string updateQuery = @"update CT_UPS_EODitems
                                   set EOD_processed = 1
                                   where id = " + id;

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
                        command.Dispose();
                        command = null;
                    }
                    catch (Exception e)
                    {
                        errorLog("18", "processEOD_EODitems", e.ToString().Substring(0, 250));
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("19", "processEOD_EODitems", e.ToString().Substring(0, 250));
                Console.WriteLine(e.ToString());
            }
        }

        //Changes the EODexport_processed from 0 to 1
        public void processEOD_EODexport(string pID)
        {
            string id = pID;

            string updateQuery = @"update CT_UPS_EODexport
                                   set EOD_processed = 1
                                   where id = " + id;

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
                        command.Dispose();
                        command = null;
                    }
                    catch (Exception e)
                    {
                        errorLog("20", "processEOD_EODitems", e.ToString().Substring(0, 250));
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("21", "processEOD_EODitems", e.ToString().Substring(0, 250));
                Console.WriteLine(e.ToString());
            }
        }

        //Helper function that retrieves the exportID based on the EODitem ID
        public string getExportID(string pID)
        {
            string id = pID;
            string exportID = "";

            string queryString = @"select exportID
                                   from CT_UPS_EODitems
                                   where id = " + id;

            ArrayList al = new ArrayList();

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(queryString, conn);
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
                                exportID = column.ToString();
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        errorLog("22", "getExportID", e.ToString().Substring(0, 250));
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("23", "getExportID", e.ToString().Substring(0, 250));
                Console.WriteLine(e.ToString());
            }

            return exportID;
        }

        //Helper function that retrieves the previous shipping charges applied to the CT Job table
        public decimal getPrevCharges(string pID)
        {
            decimal prevCost = 0;
            string id = pID;
            ArrayList al = new ArrayList();

            string getPrevQuery = @"select ActualCharge
                                    from ShipmentItems
                                    where ShipmentNumber = (select shipmentNum
                                                            from printable.dbo.CT_UPS_EODitems
                                                            where id = " + id + @")
                                    and LineN = (select shipmentLineNum
                                                            from printable.dbo.CT_UPS_EODitems
                                                            where id = " + id + ")";
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

        //Helper method that guarantees free shipping by Zip Code
        public bool freeShippingByZip_PTI(string pID)
        {
            string id = pID;
            string zip = "";
            ArrayList al = new ArrayList();

            string queryZip = @"select zip
                                from vw_UPS_EODitemsDetail  
                                where ShipmentNumber = (select shipmentNum
                                                        from CT_UPS_EODitems
                                                        where id = " + id + @")
                                and LineN = (select shipmentLineNum
                                                        from CT_UPS_EODitems
                                                        where id = " + id + ")";
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

        public int getJobN(int pti_lineItemID)
        {
            int jobN = 0;

            string q = @"SELECT TOP 1 Logic_jobNum
                         FROM printable.dbo.CT_UPS_EODitems
                         WHERE PTI_LineItemID = " + pti_lineItemID;

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



        //                                                                                                  
        // ALL NEW METHODS                                                                                  
        //                                                                                                  

        public ArrayList getInternationalShipments()
        {
            ArrayList al = new ArrayList();

            string q = @"SELECT COUNT(*), ShipmentNum, TotalShipmentCharge
                         FROM printable.dbo.CT_UPS_EODexport
                         WHERE (shipmentType LIKE 'Worldwide%'
                                OR countryCode NOT IN ('United States', 'US', 'USA', 'U.S.A.', 'U.S.'))
                         AND International_processed = 0
                         GROUP BY ShipmentNum, TotalShipmentCharge";

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
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
                    }
                    catch (Exception e)
                    {
                        errorLog("getInternationalShipments-1", e.ToString(), "getInternationalShipments Error");
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("getInternationalShipments-2", e.ToString(), "getInternationalShipments Error");
            }

            return al;
        }

        public int countShipmentLineItems(int shipmentNumber)
        {
            int temp = 1;

            string q = @"SELECT COUNT(*)
                         FROM ShipmentItems
                         WHERE ShipmentNumber = " + shipmentNumber;

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();

                        temp = Convert.ToInt32(command.ExecuteScalar());

                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        errorLog("XX", "countShipmentLineItems", e.ToString().Substring(0, 250));
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("XX", "countShipmentLineItems", e.ToString().Substring(0, 250));
                Console.WriteLine(e.ToString());
            }

            return temp;
        }

        public void updateIntlShipmentCharges(int shipmentNumber, decimal charges)
        {
            int rowsUpdated = 0;

            string q = @"UPDATE pLogic.dbo.ShipmentItems
                         SET ActualCharge = " + charges + @"
                         WHERE ShipmentNumber = " + shipmentNumber;

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();

                        rowsUpdated = command.ExecuteNonQuery();

                        if (rowsUpdated > 0)
                        {
                            processIntlShipment(shipmentNumber);
                        }

                        Console.WriteLine(rowsUpdated + " rows updated.");

                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        errorLog("XX", "updateIntlShipmentCharges", e.ToString().Substring(0, 250));
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("XX", "updateIntlShipmentCharges", e.ToString().Substring(0, 250));
                Console.WriteLine(e.ToString());
            }
        }

        private void processIntlShipment(int shipmentNumber)
        {
            int rowsProcessed = 0;

            string q = @"UPDATE printable.dbo.CT_UPS_EODexport
                         SET International_Processed = 1
                         WHERE ShipmentNum = " + shipmentNumber;

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();

                        rowsProcessed = command.ExecuteNonQuery();

                        Console.WriteLine(rowsProcessed + " rows processed.");

                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        errorLog("XX", "processIntlShipment", e.ToString().Substring(0, 250));
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                errorLog("XX", "processIntlShipment", e.ToString().Substring(0, 250));
                Console.WriteLine(e.ToString());
            }
        }

        //                                                                                                  
        //                                                                                                  
        //                                                                                                  


        //SMTP Connection Settings
        private static SmtpClient smtpClient = new SmtpClient("192.168.240.27");

        public void sendEmail(string recipID, string msgBody)
        {
            ArrayList msgList = new ArrayList();

            msgList.Clear();
            msgList.Add(recipID);
                foreach (string item in msgList)
                {
                    try
                    {
                        MailMessage message = new MailMessage();
                        message.To.Add(item);
                        message.Subject = "EOD Processing";
                        message.From = new MailAddress("FPG_Automation@Finelink.com");
                        message.Body = msgBody;
                        message.ReplyTo = new MailAddress("GVreeman@Finelink.com");
                        message.IsBodyHtml = true;
                        System.Net.Mail.SmtpClient smtp = smtpClient;
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.Write("Sending email to " + message.To.ToString());
                        smtp.Send(message);
                        Console.WriteLine(" - Success");
                        Console.ResetColor();
                        message.Dispose();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("SendEmail Error - RecipID:" + recipID + " - Receiver:" + item + " - Msg:" + msgBody);
                        Console.WriteLine(e.ToString()); 
                        Console.Beep();
                    }
                }
            
        }
    }
}

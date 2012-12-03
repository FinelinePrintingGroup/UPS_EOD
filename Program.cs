using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;
using System.Net;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Collections;

namespace UPS_EODprocessing
{
    class Program
    {
        static void Main(string[] args)
        {
            //Declare new helper methods instance
            methods me = new methods();

            //Create an array for EOD shipments to process
            ArrayList al = new ArrayList();
            me.processNoShipNum();
            me.processBadShipNum();
            al = me.getShipmentsForProcessing();
            
            try
            {
                #region BY SHIPMENT
                //Process each Shipment ID accordingly
                foreach (object[] row in al)
                {
                    foreach (object column in row)
                    {
                        //Insert ShipmentItems into EODitems table for processing
                        me.insertEODitems(column.ToString());

                        //Process EOD status
                        me.processEOD_EODexport(column.ToString());
                    }
                }
                #endregion

                #region BY SHIPMENT ITEM
                //Create and array for EOD shipment items to process
                ArrayList al2 = new ArrayList();
                al2 = me.getShipmentItemsForProcessing();

                //Process each Shipment Item ID accordingly
                foreach (object[] row in al2)
                {
                    foreach (object column in row)
                    {
                        //calculate and update respective item shipping charges
                        me.calculateItemCharges(column.ToString());

                        //Process EOD status
                        me.processEOD_EODitems(column.ToString());
                    }
                }
                #endregion

                #region NON-UPS ITEMS
                //INSERT AND PROCESS ALL NON UPS EOD ITEMS
                me.insertEODitems_nonUPS();
                #endregion

                #region INTERNATIONAL
                //                                                                  
                // Process International Shipments Accordingly                      
                //                                                                  

                Console.WriteLine(Environment.NewLine + "===============================");
                Console.WriteLine("=   INTERNATIONAL SHIPMENTS   =");
                Console.WriteLine("===============================");

                int lineCount = 0;
                int shipmentNum = 0;
                decimal totalCharge = 0.00M;
                decimal lineCharge = 0.00M;

                ArrayList al3 = new ArrayList();
                al3 = me.getInternationalShipments();

                foreach (object[] row in al3)
                {
                    if (Int32.TryParse(row[1].ToString(), out shipmentNum))
                    {
                        shipmentNum = Convert.ToInt32(row[1].ToString());

                        totalCharge = Convert.ToDecimal(row[2].ToString());

                        lineCount = me.countShipmentLineItems(shipmentNum);

                        lineCharge = totalCharge / lineCount;

                        Console.WriteLine(Environment.NewLine + "Shipment:" + shipmentNum + " - $" + totalCharge);
                        Console.WriteLine("   Lines:" + lineCount + " - $" + lineCharge + Environment.NewLine);

                        me.updateIntlShipmentCharges(shipmentNum, lineCharge);
                    }
                    else
                    {
                        me.sendEmail("keenan@finelink.com", row[1].ToString() + " Shipment: Bad Conversion");
                    }
                    
                }
                //                                                                  
                //                                                                  
                //                                                                  
                    #endregion

                #region MISSED UPS SHIPMENTS
                foreach (int shipment in me.getMissedUPSShipments())
                {
                    me.insertEODitems_UPSmissedEOD(shipment);
                }
                #endregion
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }
    }
}
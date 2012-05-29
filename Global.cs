using System;
using System.Collections.Generic;
using System.Text;

namespace UPS_EODprocessing
{
    //This class is used to define GLOBAL variables only. 
    class Globals
    {
        //Connection String Declarations
        //private static string logicConnString = "Data Source=PLM;Initial Catalog=devLogic;User ID=FPGwebservice;Password=kissmygrits";
        //private static string printableConnString = "Data Source=PLM;Initial Catalog=printable;User ID=FPGwebservice;Password=kissmygrits";

        private static string logicConnString = "Data Source=SQL1;Initial Catalog=pLogic;User ID=FPGwebservice;Password=kissmygrits";
        private static string logicConnString_TO = "Data Source=SQL1;Initial Catalog=pLogic;User ID=FPGwebservice;Password=kissmygrits;Timeout=10";
        private static string printableConnString = "Data Source=SQL1;Initial Catalog=printable;User ID=FPGwebservice;Password=kissmygrits";


        //Accessor Methods
        //Use these for access "private" global variables for data protection
        public static string get_logicConnString
        {
            get { return Globals.logicConnString; }
        }

        public static string get_printableConnString
        {
            get { return Globals.printableConnString; }
        }

        public static string get_logicConnString_TO
        {
            get { return Globals.logicConnString_TO; }
        }
    }
}

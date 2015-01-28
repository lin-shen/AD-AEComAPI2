using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using MYOB.Tax.TaxMan.Client.ComApi;
using MYOB.Tax.TaxMan.Client.BusinessLogic;
using MYOB.Tax.TaxMan.Client.ComApi.Dto;
//using MYOB.Tax.TaxMan.Client.ComApi.Properties;
using MYOB.Tax.TaxMan.Model.EntityClasses;
using MYOB.AF.Core.DataObjects.Company;



namespace MYOB.Tax.TaxMan.Client.AEComAPI
{
    static class AEComAPI_old
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        //[STAThread]
        static void Main_backup(string[] args)
        {
            // MM before any API can be used, Initialise must be run first
            string param1 = Convert.ToString(args[0]); // This is the API to be run, for example GetAgencyForClient
            string param2 = Convert.ToString(args[1]); // For Initialise API. This is aoClassicDataPath (in this case the SOL64 folder)
            string param3 = Convert.ToString(args[2]); // For initialise API. This is aoEmployeeIdentifier (Employee client code for AE)
            string param4; // Params required by API or output file name
            string param5; // Params required by API or output file name
            string param6; // Params required by API or output file name

      
            switch (param1)
            {
                 case "GetAgencyForClient":

                    param4 = Convert.ToString(args[3]); // aoContactIdentifier
                    param5 = Convert.ToString(args[4]); // Output file name and path
 
                    TaxManagerApi init = new TaxManagerApi();
                    init.Initialise(param2, param3);

                    AgencyComDto temp = init.GetAgencyForClient(param3);
                    temp = init.GetAgencyForClient(param4);
                    string agentDets = temp.AgencyName;
                    if (agentDets == "")
                    { 
                       System.IO.File.WriteAllText(param5, "No agent details found"); 
                    }
                    else
                    {
                       System.IO.File.WriteAllText(param5, agentDets);
                    }                  

                    break;

                 default:
                    //
                    break;
            }
            
          
        }
    }
}

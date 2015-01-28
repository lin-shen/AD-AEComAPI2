using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using MYOB.Tax.TaxMan.Client.ComApi;
using MYOB.Tax.TaxMan.Client.BusinessLogic;
using MYOB.Tax.TaxMan.Client.ComApi.Dto;
using MYOB.Tax.TaxMan.Model.EntityClasses;
using MYOB.AF.Core.DataObjects.Company;
using System.Reflection;

namespace MYOB.Tax.TaxMan.Client.AEComAPI
{
    class AEComAPI
    {
        public static TaxManagerApi aoComAPI;
        public static string[] paramAry;
        [STAThread]
        static void Main(string[] args)
        {
            
            string methodName = args[0];
            paramAry = args;
            ComApiWrapper.initialise();

            for (MethodSelected enum1 = MethodSelected.Initialise; enum1 <= MethodSelected.GetFormTypeCode; enum1++)
            {
                if (enum1.ToString() == methodName)
                {
                    Assembly assembly = Assembly.GetExecutingAssembly();
                    Type type = assembly.GetType("MYOB.Tax.TaxMan.Client.AEComAPI.ComApiWrapper", false, true);
                    MethodInfo method = type.GetMethod(methodName);
                    method.Invoke(null, null);
                }
            }
        }
    }
}

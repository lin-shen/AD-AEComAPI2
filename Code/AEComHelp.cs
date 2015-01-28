using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.Serialization;

namespace MYOB.Tax.TaxMan.Client.AEComAPI
{
    static class AEComHelp
    {
        public static void ShowMsg(string msg)
        {
            MessageBox.Show("It's in ShowMsg() and message is - " + msg);
        }
    }

    enum MethodSelected
    {
        Initialise,
        GetAgencies,
        GetAgencyForClient,
        GetPracticeDefaultSettings,
        GetTaxManagerStartYear,
        GetLosses,
        UpdateLosses,
        GetTaxTypes,
        GetAssessedDate,
        HasEstimateForYear,
        GetTaxTransactions,
        OpenTaxDocument,
        SaveTaxDocument,
        DeleteTaxDocument,
        CloseTaxDocument,
        GetTaxDocumentId,
        UpdateTaxDocumentDetails,
        UpdateTaxDocumentStatus,
        AddTaxDocumentTransferTransaction,
        AddTaxDocumentTransaction,
        ReplaceExistingTaxDocumentTransactions,
        GetProvisionalTaxDetails,
        GetInterimTaxDetails,
        GetFormTypeCode
    }
}

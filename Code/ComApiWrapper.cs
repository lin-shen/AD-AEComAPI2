using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MYOB.Tax.TaxMan.Client.ComApi;
//using MYOB.Tax.TaxMan.Client.BusinessLogic;
using MYOB.Tax.TaxMan.Client.ComApi.Dto;
using MYOB.Tax.TaxMan.Model.EntityClasses;
//using MYOB.AF.Core.DataObjects.Company;
using MYOB.Common;
using System.Reflection;

namespace MYOB.Tax.TaxMan.Client.AEComAPI
{
    public static class ComApiWrapper
    {

        public static void initialise()
        {
            try
            {
                string aePath = AEComAPI.paramAry[1];
                string aeEmployeeId = AEComAPI.paramAry[2];
                AEComAPI.aoComAPI = new TaxManagerApi();
                AEComAPI.aoComAPI.Initialise(aePath, aeEmployeeId);
            }
            catch (Exception ex)
            {
            }
        }

        public static void GetAgencies()
        {
            //Params: GetAgencies c:\MYOBAO\DataPM _3NG12GEZU c:\temp\sltest.txt        
            if (AEComAPI.paramAry.Length == 4)
            {
                string outPathName = AEComAPI.paramAry[3]; // Output file name and path
                string tempString = string.Empty;
                string outString = string.Empty;
                AgencyComDto[] agencies = AEComAPI.aoComAPI.GetAgencies();

                //With returned agency array, only write agant name to file for now
                if (agencies.Count() == 0)
                    System.IO.File.WriteAllText(outPathName, "GetAgencies record not found");
                else
                {
                    foreach (AgencyComDto agency in agencies)
                    {
                        tempString = string.Empty;
                        Type agentType = agency.GetType();

                        foreach (PropertyInfo agentInfo in agentType.GetProperties())
                            tempString = tempString + ',' + agentInfo.GetValue(agency, null);

                        tempString = tempString.Substring(1);
                        outString = outString + Environment.NewLine + tempString;
                    }
                    outString = outString.Substring(1);
                    System.IO.File.WriteAllText(outPathName, outString);
                }
            }
        }

        public static void GetAgencyForClient()
        {
            //Params: GetAgencyForClient c:\MYOBAO\DataPM _3NG12GEZU AAAAA_T98H c:\temp\sltest.txt  
            if (AEComAPI.paramAry.Length == 5)
            {
                string aoContactId = AEComAPI.paramAry[3]; ; // aoContactIdentifier
                string outPathName = AEComAPI.paramAry[4]; // Output file name and path
                string outString = string.Empty;

                AgencyComDto agency = AEComAPI.aoComAPI.GetAgencyForClient(aoContactId);

                //With returned agency array, only write agant name to file for now
                if (string.IsNullOrEmpty(agency.ToString()))
                {
                    System.IO.File.WriteAllText(outPathName, "GetAgencyForClient record not found");
                }
                else
                {
                    Type agencyType = agency.GetType();
                    foreach (PropertyInfo agentInfo in agencyType.GetProperties())
                        outString = outString + ',' + agentInfo.GetValue(agency, null);

                    outString = outString.Substring(1);
                    System.IO.File.WriteAllText(outPathName, outString);
                }
            }
        }

        public static void GetPracticeDefaultSettings()
        {
            //Params: GetPracticeDefaultSettings c:\MYOBAO\DataPM _3NG12GEZU c:\temp\sltest.txt        
            if (AEComAPI.paramAry.Length == 4)
            {
                string outPathName = AEComAPI.paramAry[3]; // Output file name and path

                PracticeDefaultSettingsComDto settings = AEComAPI.aoComAPI.GetPracticeDefaultSettings();

                //With returned settings array, write settings and agency name for now
                if (string.IsNullOrEmpty(settings.ToString()))
                    System.IO.File.WriteAllText(outPathName, "GetPracticeDefaultSettings record not found");
                else
                {
                    string returnDetails = string.Empty;
                    string agencyDetails = string.Empty;
                    Type settingsType = settings.GetType();
                    foreach (PropertyInfo settingsInfo in settingsType.GetProperties())
                    {
                        if (settingsInfo.Name == "DefaultAgency")
                        {
                            Type defaultAgencyType = settings.DefaultAgency.GetType();
                            foreach (PropertyInfo defaultAgencyInfo in defaultAgencyType.GetProperties())
                                agencyDetails = agencyDetails + ',' + defaultAgencyInfo.GetValue(settings.DefaultAgency, null);
                            agencyDetails = Environment.NewLine + agencyDetails.Substring(1);
                        }
                        else
                            returnDetails = returnDetails + ',' + settingsInfo.GetValue(settings, null);
                    }
                    returnDetails = returnDetails.Substring(1) + agencyDetails;
                    System.IO.File.WriteAllText(outPathName, returnDetails);
                }
            }
        }

        public static void GetTaxManagerStartYear()
        {
            //Params: GetTaxManagerStartYear c:\MYOBAO\DataPM _3NG12GEZU c:\temp\sltest.txt        
            if (AEComAPI.paramAry.Length == 4)
            {
                string outPathName = AEComAPI.paramAry[3]; // Output file name and path

                int startYear = AEComAPI.aoComAPI.GetTaxManagerStartYear();

                if (string.IsNullOrEmpty(startYear.ToString()))
                    System.IO.File.WriteAllText(outPathName, "GetTaxManagerStartYear record not found");
                else
                    System.IO.File.WriteAllText(outPathName, startYear.ToString());
            }
        }

        public static void GetLosses()
        {
            //Params: GetTaxManagerStartYear c:\MYOBAO\DataPM _3NG12GEZU AAAAA_T98H 2013 c:\temp\sltest.txt        
            if (AEComAPI.paramAry.Length == 6)
            {
                string aoContactId = AEComAPI.paramAry[3];
                int Year = int.Parse(AEComAPI.paramAry[4]);
                string outPathName = AEComAPI.paramAry[5]; // Output file name and path

                LossesComDto lossDetails = AEComAPI.aoComAPI.GetLosses(aoContactId, Year);

                System.IO.File.WriteAllText(outPathName, lossDetails.LossesCarriedForward.ToString() + "," + lossDetails.ImputationCreditsCarriedForward.ToString());
            }
        }

        public static void UpdateLosses()
        {
            //Params: UpdateLosses c:\MYOBAO\DataPM _3NG12GEZU AAAAA_T98H 2013 I 123.45 23.45 c:\temp\sltest.txt        
            if (AEComAPI.paramAry.Length == 9)
            {
                string aoContactId = AEComAPI.paramAry[3];
                int year = int.Parse(AEComAPI.paramAry[4]);
                string taxTypeCode = AEComAPI.paramAry[5];
                decimal lossBfd = decimal.Parse(AEComAPI.paramAry[6]);
                decimal impBfd = decimal.Parse(AEComAPI.paramAry[7]);
                string outPathName = AEComAPI.paramAry[8]; // Output file name and path

                //Currently only INC tax type can be updated back to TM - ignore other types
                if (taxTypeCode != "I")
                    return;

                LossesComDto lossDetails = new LossesComDto();
                lossDetails.LossesCarriedForward = lossBfd;
                lossDetails.ImputationCreditsCarriedForward = impBfd;

                TaxTypeComDto[] taxTypes = AEComAPI.aoComAPI.GetTaxTypes();
                TaxTypeComDto taxType = taxTypes.Where(o => o.Code.Substring(0, taxTypeCode.Length) == taxTypeCode).FirstOrDefault();
                if (taxType == null)
                    System.IO.File.WriteAllText(outPathName, "UpdateLosses record not found");
                else
                {
                    AEComAPI.aoComAPI.UpdateLosses(aoContactId, year, taxType.TaxTypeId, lossDetails);

                    System.IO.File.WriteAllText(outPathName, "UpdateLosses was processed");
                }
            }
        }

        public static void GetTaxTypes()
        {
            //Params: GetTaxTypes c:\MYOBAO\DataPM _3NG12GEZU c:\temp\sltest.txt  
            string outPathName = AEComAPI.paramAry[3]; // Output file name and path
            string outString = string.Empty;
            string tempString = string.Empty;

            if (AEComAPI.paramAry.Length == 4)
            {
                TaxTypeComDto[] taxTypes = AEComAPI.aoComAPI.GetTaxTypes();
                if (taxTypes.Count() == 0)
                {
                    System.IO.File.WriteAllText(outPathName, "GetTaxTypes record not find");
                }
                else
                {
                    foreach (TaxTypeComDto taxType in taxTypes)
                    {
                        if (outString != string.Empty)
                            outString = outString + Environment.NewLine;

                        tempString = string.Empty;

                        Type taxTypeInfo = taxType.GetType();
                        foreach (PropertyInfo propertyInfo in taxTypeInfo.GetProperties())
                        {
                            tempString = tempString + ',' + propertyInfo.GetValue(taxType, null);
                        }
                        outString = outString + tempString.Substring(1);
                    }

                    System.IO.File.WriteAllText(outPathName, outString);
                }
            }
        }

        public static void GetAssessedDate()
        {
            //Params: GetAssessedDate c:\MYOBAO\DataPM _3NG12GEZU AAAAA_T98H 2013 I c:\temp\sltest.txt  
            string aoContactId = AEComAPI.paramAry[3];
            int year = int.Parse(AEComAPI.paramAry[4]);
            string taxTypeCode = AEComAPI.paramAry[5];
            string outPathName = AEComAPI.paramAry[6];   // Output file name and path

            if (AEComAPI.paramAry.Length == 7)
            {
                TaxTypeComDto[] taxTypes = AEComAPI.aoComAPI.GetTaxTypes();
                TaxTypeComDto taxType = taxTypes.Where(o => o.Code.Substring(0, taxTypeCode.Length) == taxTypeCode).FirstOrDefault();
                if (taxType == null)
                    System.IO.File.WriteAllText(outPathName, "GetAssessedDate record not find");
                else
                {
                    DateTime assessedDate = AEComAPI.aoComAPI.GetAssessedDate(aoContactId, year, taxType.TaxTypeId);
                    System.IO.File.WriteAllText(outPathName, assessedDate.ToString());
                }
            }
        }

        public static void HasEstimateForYear()
        {
            //TODO check this method existance in COMAPI !!!
            //Params: HasEstimateForYear c:\MYOBAO\DataPM _3NG12GEZU AAAAA_T98H 2013 I c:\temp\sltest.txt  
            string aoContactId = AEComAPI.paramAry[3];
            int year = int.Parse(AEComAPI.paramAry[4]);
            string taxTypeCode = AEComAPI.paramAry[5];
            string outPathName = AEComAPI.paramAry[6];   // Output file name and path

            if (AEComAPI.paramAry.Length == 7)
            {
                TaxTypeComDto[] taxTypes = AEComAPI.aoComAPI.GetTaxTypes();
                TaxTypeComDto taxType = taxTypes.Where(o => o.Code.Substring(0, taxTypeCode.Length) == taxTypeCode).FirstOrDefault();

                if (taxType == null)
                    System.IO.File.WriteAllText(outPathName, "HasEstimateForYear record not find");
                else
                {
                    //TODO check this method existance in COMAPI !!!
                    // bool assessedDate = AEComAPI.aoComAPI.HasEstimateForYear(aoContactId, year, taxTypeId);
                    // System.IO.File.WriteAllText(outPathName, assessedDate.ToString());
                }
            }
        }
        public static void GetTaxTransactions()
        {
            //Params: GetTaxTransactions c:\MYOBAO\DataPM _3NG12GEZU AAAAA_T98H 2013 c:\temp\sltest.txt  
            string aoContactId = AEComAPI.paramAry[3];
            int year = int.Parse(AEComAPI.paramAry[4]);
            string outPathName = AEComAPI.paramAry[5];   // Output file name and path
            string outString = string.Empty;
            string tempString = string.Empty;

            if (AEComAPI.paramAry.Length == 6)
            {
                TaxTransactionComDto[] taxTransactions = AEComAPI.aoComAPI.GetTaxTransactions(aoContactId, year);

                if (taxTransactions.Count() == 0)
                    System.IO.File.WriteAllText(outPathName, "GetTaxTransactions record not find");
                else
                {
                    foreach (TaxTransactionComDto taxTransaction in taxTransactions)
                    {
                        if (outString != string.Empty)
                            outString = outString + Environment.NewLine;

                        Type taxTransactionType = taxTransaction.GetType();
                        foreach (PropertyInfo taxTransactionInfo in taxTransactionType.GetProperties())
                        {
                            tempString = tempString + "," + taxTransactionInfo.GetValue(taxTransaction, null);
                        }
                        outString = outString + tempString.Substring(1);
                    }

                    System.IO.File.WriteAllText(outPathName, outString);
                }
            }
        }
        public static void OpenTaxDocument()
        {
            //Params: OpenTaxDocument c:\MYOBAO\DataPM _3NG12GEZU AAAAA_T98H 2013 IR3 True c:\temp\sltest.txt  
            string aoContactId = AEComAPI.paramAry[3];
            short year = short.Parse(AEComAPI.paramAry[4]);
            string formTypeCode = AEComAPI.paramAry[5];
            bool isPrimary = Boolean.Parse(AEComAPI.paramAry[6]);
            string outPathName = AEComAPI.paramAry[7];   // Output file name and path

            if (AEComAPI.paramAry.Length == 8)
            {
                AEComAPI.aoComAPI.OpenTaxDocument(aoContactId, year, formTypeCode, isPrimary);
                System.IO.File.WriteAllText(outPathName, "OpenTaxDocument was processed.");
            }
        }
        public static void SaveTaxDocument()
        {
            //Params: SaveTaxDocument c:\MYOBAO\DataPM _3NG12GEZU c:\temp\sltest.txt 
            string outPathName = AEComAPI.paramAry[3];   // Output file name and path
            if (AEComAPI.paramAry.Length == 4)
            {
                AEComAPI.aoComAPI.SaveTaxDocument();
                System.IO.File.WriteAllText(outPathName, "SaveTaxDocument was processed.");
            }
        }
        public static void DeleteTaxDocument()
        {
            //Params: DeleteTaxDocument c:\MYOBAO\DataPM _3NG12GEZU c:\temp\sltest.txt 
            string outPathName = AEComAPI.paramAry[3];   // Output file name and path
            if (AEComAPI.paramAry.Length == 4)
            {
                AEComAPI.aoComAPI.DeleteTaxDocument();
                System.IO.File.WriteAllText(outPathName, "DeleteTaxDocument was processed.");
            }
        }
        public static void CloseTaxDocument()
        {
            //Params: DeleteTaxDocument c:\MYOBAO\DataPM _3NG12GEZU c:\temp\sltest.txt  
            string outPathName = AEComAPI.paramAry[3];   // Output file name and path
            if (AEComAPI.paramAry.Length == 4)
            {
                AEComAPI.aoComAPI.CloseTaxDocument();
                System.IO.File.WriteAllText(outPathName, "CloseTaxDocument was processed.");
            }
        }
        public static void GetTaxDocumentId()
        {
            //Params: GetTaxDocumentId c:\MYOBAO\DataPM _3NG12GEZU c:\temp\sltest.txt  
            string outPathName = AEComAPI.paramAry[3];   // Output file name and path
            if (AEComAPI.paramAry.Length == 4)
            {
                int taxDocumentId = AEComAPI.aoComAPI.GetTaxDocumentId();
                System.IO.File.WriteAllText(outPathName, taxDocumentId.ToString());
            }
        }
        public static void UpdateTaxDocumentDetails()
        {
            //Params: UpdateTaxDocumentDetails c:\MYOBAO\DataPM _3NG12GEZU AAAAA_T98H 2013 IR3 True 0 0 0 0 0 0 0 0 0 0 01/01/2015 0 01/01/2015 0 c:\temp\sltest.txt 
            string aoContactId = AEComAPI.paramAry[3];
            short year = short.Parse(AEComAPI.paramAry[4]);
            string formTypeCode = AEComAPI.paramAry[5];
            bool isPrimary = Boolean.Parse(AEComAPI.paramAry[6]);
            decimal? rit = decimal.Parse(AEComAPI.paramAry[7]);
            decimal? taxableIncome = decimal.Parse(AEComAPI.paramAry[8]);
            decimal? taxThereon = decimal.Parse(AEComAPI.paramAry[9]);
            decimal? taxCredits = decimal.Parse(AEComAPI.paramAry[10]);
            decimal? familyCredits = decimal.Parse(AEComAPI.paramAry[11]);
            decimal? studentLoan = decimal.Parse(AEComAPI.paramAry[12]);
            int? provisionalBasis = int.Parse(AEComAPI.paramAry[13]);
            decimal? provAmount = decimal.Parse(AEComAPI.paramAry[14]);
            int? provInstallments = int.Parse(AEComAPI.paramAry[15]);
            decimal? uomiAmount = decimal.Parse(AEComAPI.paramAry[16]);
            DateTime? uomiDate = DateTime.Parse(AEComAPI.paramAry[17]);
            decimal? studentLoanProv = decimal.Parse(AEComAPI.paramAry[18]);
            DateTime? returnEffectiveDate = DateTime.Parse(AEComAPI.paramAry[19]);
            decimal? beneficiaryTax = decimal.Parse(AEComAPI.paramAry[20]);
            string outPathName = AEComAPI.paramAry[21];   // Output file name and path

            if (AEComAPI.paramAry.Length == 22)
            {
                TaxDocumentDetailsComDto taxDocumentDetails = new TaxDocumentDetailsComDto()
                    {
                        Rit = rit,
                        TaxableIncome = taxableIncome,
                        TaxThereon = taxThereon,
                        TaxCredits = taxCredits,
                        FamilyCredits = familyCredits,
                        StudentLoan = studentLoan,
                        ProvisionalBasis = provisionalBasis,
                        ProvAmount = provAmount,
                        ProvInstallments = provInstallments,
                        UomiAmount = uomiAmount,
                        UomiDate = uomiDate,
                        StudentLoanProv = studentLoanProv,
                        ReturnEffectiveDate = returnEffectiveDate,
                        BeneficiaryTax = beneficiaryTax
                    };

                AEComAPI.aoComAPI.OpenTaxDocument(aoContactId, year, formTypeCode, isPrimary);
                AEComAPI.aoComAPI.UpdateTaxDocumentDetails(taxDocumentDetails);
                AEComAPI.aoComAPI.SaveTaxDocument();
                System.IO.File.WriteAllText(outPathName, "UpdateTaxDocumentDetails was processed.");
            }
        }

        public static void UpdateTaxDocumentStatus()
        {
            //Params: DeleteTaxDocument c:\MYOBAO\DataPM _3NG12GEZU 01/01/2015 01/01/2015 01/01/2015 01/01/2015 01/01/2015 c:\temp\sltest.txt 
            TaxDocumentStatusDateDto inProgressDate = new TaxDocumentStatusDateDto() { Date = DateTime.Parse(AEComAPI.paramAry[3]) };
            TaxDocumentStatusDateDto approvedDate = new TaxDocumentStatusDateDto() { Date = DateTime.Parse(AEComAPI.paramAry[4]) };
            TaxDocumentStatusDateDto sentToEFileDate = new TaxDocumentStatusDateDto() { Date = DateTime.Parse(AEComAPI.paramAry[5]) };
            TaxDocumentStatusDateDto filedDate = new TaxDocumentStatusDateDto() { Date = DateTime.Parse(AEComAPI.paramAry[6]) };
            TaxDocumentStatusDateDto assessedDate = new TaxDocumentStatusDateDto() { Date = DateTime.Parse(AEComAPI.paramAry[7]) };
            string outPathName = AEComAPI.paramAry[8];   // Output file name and path

            if (AEComAPI.paramAry.Length == 9)
            {
                TaxDocumentStatusComDto taxDocumentStatus = new TaxDocumentStatusComDto()
                    {
                        InProgressDate = (TaxDocumentStatusDateDto)inProgressDate,
                        ApprovedDate = approvedDate,
                        SentToEFileDate = sentToEFileDate,
                        FiledDate = filedDate,
                        AssessedDate = assessedDate
                    };
                AEComAPI.aoComAPI.UpdateTaxDocumentStatus(taxDocumentStatus);
                System.IO.File.WriteAllText(outPathName, "UpdateTaxDocumentStatus was processed.");
            }
        }
        public static void AddTaxDocumentTransferTransaction()
        {
            //Params: AddTaxDocumentTransferTransaction c:\MYOBAO\DataPM _3NG12GEZU I 100 01/01/2015 advance 069038638 31/03/2015 S c:\temp\sltest.txt
            string taxTypeCode = AEComAPI.paramAry[3];
            decimal amount = Decimal.Parse(AEComAPI.paramAry[4]);
            DateTime effectiveDate = DateTime.Parse(AEComAPI.paramAry[5]);
            string comment = AEComAPI.paramAry[6];
            string otherPartyIrdNo = AEComAPI.paramAry[7];
            DateTime otherPartyTaxPeriodEnd = DateTime.Parse(AEComAPI.paramAry[8]);
            string otherPartyTaxTypeCode = AEComAPI.paramAry[9];
            string outPathName = AEComAPI.paramAry[10];   // Output file name and path

            if (AEComAPI.paramAry.Length == 11 && !string.IsNullOrEmpty(otherPartyIrdNo.ToString()))
            {
                TaxTypeComDto[] taxTypes = AEComAPI.aoComAPI.GetTaxTypes();
                TaxTypeComDto taxType = taxTypes.Where(o => o.Code.Substring(0, taxTypeCode.Length) == taxTypeCode).FirstOrDefault();
                TaxTypeComDto otherPartyTaxType = taxTypes.Where(o => o.Code.Substring(0, taxTypeCode.Length) == otherPartyTaxTypeCode).FirstOrDefault();
                if (taxType == null || otherPartyTaxType == null)
                    System.IO.File.WriteAllText(outPathName, "AddTaxDocumentTransferTransaction record not found.");
                else
                {
                    TaxDocumentTransferTransactionComDto taxDocumentTransferTransaction = new TaxDocumentTransferTransactionComDto()
                        {
                            TaxTypeId = taxType.TaxTypeId,
                            Amount = amount,
                            EffectiveDate = effectiveDate,
                            Comment = comment,
                            OtherPartyIrdNo = otherPartyIrdNo,
                            OtherPartyTaxPeriodEnd = otherPartyTaxPeriodEnd,
                            OtherPartyTaxTypeId = otherPartyTaxType.TaxTypeId
                        };

                    AEComAPI.aoComAPI.AddTaxDocumentTransferTransaction(taxDocumentTransferTransaction);
                    System.IO.File.WriteAllText(outPathName, "AddTaxDocumentTransferTransaction was processed.");
                }
            }
        }
        public static void AddTaxDocumentTransaction()
        {
            //Params: AddTaxDocumentTransaction c:\MYOBAO\DataPM _3NG12GEZU AAAAA_T98H 2013 IR3 True I 0830 12.34 01/01/2015 payment c:\temp\sltest.txt
            string aoContactId = AEComAPI.paramAry[3];
            short year = short.Parse(AEComAPI.paramAry[4]);
            string formTypeCode = AEComAPI.paramAry[5];
            bool isPrimary = Boolean.Parse(AEComAPI.paramAry[6]);
            string taxTypeCode = AEComAPI.paramAry[7];
            string taxTransactionType = AEComAPI.paramAry[8];
            decimal amount = Decimal.Parse(AEComAPI.paramAry[9]);
            DateTime effectiveDate = DateTime.Parse(AEComAPI.paramAry[10]);
            string comment = AEComAPI.paramAry[11];
            string outPathName = AEComAPI.paramAry[12];   // Output file name and path

            if (AEComAPI.paramAry.Length == 13)
            {
                TaxTypeComDto[] taxTypes = AEComAPI.aoComAPI.GetTaxTypes();
                TaxTypeComDto taxType = taxTypes.Where(o => o.Code.Substring(0, taxTypeCode.Length) == taxTypeCode).FirstOrDefault();
                if (taxType == null)
                    System.IO.File.WriteAllText(outPathName, "AddTaxDocumentTransaction record not found.");
                else
                {
                    TaxDocumentTransactionComDto taxDocumentTransaction = new TaxDocumentTransactionComDto()
                    {
                        TaxTypeId = taxType.TaxTypeId,
                        TaxTransactionType = taxTransactionType,
                        Amount = amount,
                        EffectiveDate = effectiveDate,
                        Comment = comment
                    };

                    AEComAPI.aoComAPI.OpenTaxDocument(aoContactId, year, formTypeCode, isPrimary);
                    AEComAPI.aoComAPI.AddTaxDocumentTransaction(taxDocumentTransaction);
                    AEComAPI.aoComAPI.SaveTaxDocument();
                    System.IO.File.WriteAllText(outPathName, "AddTaxDocumentTransaction was processed.");
                }
            }
        }
        public static void ReplaceExistingTaxDocumentTransactions()
        {
            //Params: ReplaceExistingTaxDocumentTransactions c:\MYOBAO\DataPM _3NG12GEZU c:\temp\sltest.txt
            string outPathName = AEComAPI.paramAry[3];   // Output file name and path

            if (AEComAPI.paramAry.Length == 4)
            {
                AEComAPI.aoComAPI.ReplaceExistingTaxDocumentTransactions();
                System.IO.File.WriteAllText(outPathName, "ReplaceExistingTaxDocumentTransactions was processed.");
            }
        }
        public static void GetProvisionalTaxDetails()
        {
            //Params: GetProvisionalTaxDetails c:\MYOBAO\DataPM _3NG12GEZU AAAAA_T98H 2014 I c:\temp\sltest.txt     
            string aoContactId = AEComAPI.paramAry[3];
            int year = int.Parse(AEComAPI.paramAry[4]);
            string taxTypeCode = AEComAPI.paramAry[5];
            string outPathName = AEComAPI.paramAry[6];
            string outString = string.Empty;

            if (AEComAPI.paramAry.Length == 7)
            {
                TaxTypeComDto[] taxTypes = AEComAPI.aoComAPI.GetTaxTypes();
                TaxTypeComDto taxType = taxTypes.Where(o => o.Code.Substring(0, taxTypeCode.Length) == taxTypeCode).FirstOrDefault();
                if (taxType == null)
                    System.IO.File.WriteAllText(outPathName, "GetProvisionalTaxDetails record not found.");
                else
                {
                    ProvisionalTaxDetailsComDto provisionalTaxDetails = AEComAPI.aoComAPI.GetProvisionalTaxDetails(aoContactId, year, taxType.TaxTypeId);

                    Type provisionalTaxType = provisionalTaxDetails.GetType();
                    foreach (PropertyInfo provisionalTaxInfo in provisionalTaxType.GetProperties())
                    {
                        outString = outString + "," + provisionalTaxInfo.GetValue(provisionalTaxDetails, null);
                    }
                    outString = outString.Substring(1);
                    System.IO.File.WriteAllText(outPathName, outString);
                }
            }
        }
        public static void GetInterimTaxDetails()
        {
            //Params: GetInterimTaxDetails c:\MYOBAO\DataPM _3NG12GEZU AAAAA_T98H 2014 I c:\temp\sltest.txt     
            string aoContactId = AEComAPI.paramAry[3];
            int year = int.Parse(AEComAPI.paramAry[4]);
            string taxTypeCode = AEComAPI.paramAry[5];
            string outPathName = AEComAPI.paramAry[6];
            string outString = string.Empty;

            if (AEComAPI.paramAry.Length == 7)
            {
                TaxTypeComDto[] taxTypes = AEComAPI.aoComAPI.GetTaxTypes();
                TaxTypeComDto taxType = taxTypes.Where(o => o.Code.Substring(0, taxTypeCode.Length) == taxTypeCode).FirstOrDefault();

                if (taxType == null)
                    System.IO.File.WriteAllText(outPathName, "GetInterimTaxDetails record not found.");
                else
                {
                    InterimTaxDetailsComDto interimTaxDetails = AEComAPI.aoComAPI.GetInterimTaxDetails(aoContactId, year, taxType.TaxTypeId);

                    Type interimTaxDetailType = interimTaxDetails.GetType();
                    foreach (PropertyInfo provisionalTaxInfo in interimTaxDetailType.GetProperties())
                    {
                        outString = outString + "," + provisionalTaxInfo.GetValue(interimTaxDetails, null);
                    }
                    outString = outString.Substring(1);
                    System.IO.File.WriteAllText(outPathName, outString);
                }
            }
        }
        public static void GetFormTypeCode()
        {
            //Params: GetFormTypeCode c:\MYOBAO\DataPM _3NG12GEZU AAAAA_T98H 2014 true c:\temp\sltest.txt     
            string aoContactId = AEComAPI.paramAry[3];
            short year = short.Parse(AEComAPI.paramAry[4]);
            bool isPrimary = Boolean.Parse(AEComAPI.paramAry[5]);
            string outPathName = AEComAPI.paramAry[6];

            if (AEComAPI.paramAry.Length == 7)
            {
                string formTypeCode = AEComAPI.aoComAPI.GetFormTypeCode(aoContactId, year, isPrimary);
                System.IO.File.WriteAllText(outPathName, formTypeCode.ToString());
            }
        }
    }
}

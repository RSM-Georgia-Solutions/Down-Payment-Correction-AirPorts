using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using DownPaymentLogic.Classes;
using jo0urnaltest;
using SAPApi;
using SAPbobsCOM;
using SAPbouiCOM;

namespace DownPaymentLogic
{
    public static class DownPaymentLogic
    {
        /// <summary>
        ///   
        /// </summary>
        /// <param name="downPaymentToDrow"> down payment - is forma gaxsnili invoisidan (A/R ; A/P) </param>
        /// <param name="data"></param>
        /// <param name="_comp"> SAPbobsCOM company </param>
        /// <param name="formType"> invoisis(mshobeli) formis tipi (A/R ; A/P) </param>
        /// <param name="docCurrency">invoisis(mshobeli) formis valuta </param>
        /// <param name="totalInv"> Total Befor Discounts damatebuli Tax-i invoisis pormidan </param>
        /// <param name="isRateCalculated"> tu isRateCalculated true daabruna eseigi invoisis BP Currency velshi  unda cahvsvat  globalRate </param>
        /// <param name="globalRate"> tu isRateCalculated true daabruna eseigi invoisis BP Currency velshi  unda cahvsvat  globalRate </param>
        /// <param name="ratInv">invoisis(mshobeli) formis valuta </param>
        /// 
        /// 
        public static jo0urnaltest.SimpleLogger _Logger = new jo0urnaltest.SimpleLogger();
        public static void ExchangeRateCorrectionUi(DataForCalculationRate data, SAPbobsCOM.Company _comp)
        {

            data.GlobalRate = "1.0000";

            decimal paidAmountDpLc = 0m; // A/R DownPayment - ში არჩეული თანხა ლოკალურ ვალუტაში //Net AmountFC To Drow * Rate
            decimal paidAmountDpFc = 0m; //  A/R DownPayment - ში არჩეული თანხა FC //Net AmountFC To Drow 

            foreach (var downpayment in data.NetAmountsForDownPayment)
            {
                if (data.FormTypex == "133")
                {
                    //string ORCTDocEntrys = string.Empty;

                    Recordset recSet2 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                    recSet2.DoQuery("SELECT DocEntry FROM ODPI WHERE DocNum = '" + downpayment.First().Key + "'");
                    var dpDocEntry = recSet2.Fields.Item("DocEntry").Value.ToString();

                    string ORCTDocEntrys =
                        "select ORCT.DocEntry from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where RCT2.DocEntry = '" +
                        dpDocEntry + "' and InvType = 203 and ORCT.Canceled = 'N'"; // ეს არის Incoming Paymentebis docentry -ები



                    recSet2.DoQuery("select   ORCT.DocEntry, SUM(RCT2.AppliedFC) as 'AppliedFC', ORCT.DocDate from ORCT inner join RCT2 on " +
                           "ORCT.DocEntry = RCT2.DocNum inner join ORTT on ORCT.DocDate = ORTT.RateDate where ORCT.DocEntry in (" + ORCTDocEntrys + ") and ORTT.Currency = '" + data.DocCurrency + "' group by ORCT.DocEntry, ORCT.DocDate  ");


                    List<Tuple<int, DateTime, decimal>> sumPayments = new List<Tuple<int, DateTime, decimal>>();
                    while (!recSet2.EoF)
                    {
                        int OCRTDocEntry = int.Parse(recSet2.Fields.Item("DocEntry").Value.ToString());
                        DateTime DocDate = DateTime.Parse(recSet2.Fields.Item("DocDate").Value.ToString());
                        decimal appliedAmountFc = decimal.Parse(recSet2.Fields.Item("AppliedFC").Value.ToString());
                        sumPayments.Add(new Tuple<int, DateTime, decimal>(OCRTDocEntry, DocDate, appliedAmountFc));
                        recSet2.MoveNext();
                    }



                    decimal weightedAvarageForPayment = sumPayments.Sum(doc => (decimal)UiManager.GetCurrencyRate(data.DocCurrency, doc.Item2, _comp) * doc.Item3) / sumPayments.Sum(doc => doc.Item3);

                    //decimal WeightedRate = LCSum / FCSum;

                    paidAmountDpLc += weightedAvarageForPayment * decimal.Parse(downpayment.First().Value);
                    paidAmountDpFc += decimal.Parse(downpayment.First().Value);


                }
                else
                {


                    Recordset recSet2 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                    recSet2.DoQuery("SELECT DocEntry FROM ODPO WHERE DocNum = '" + downpayment.First().Key + "'");
                    var dpDocEntry = recSet2.Fields.Item("DocEntry").Value.ToString();

                    string ORCTDocEntrys =
                        "select OVPM.DocEntry from OVPM inner join VPM2 on OVPM.DocEntry = VPM2.DocNum where VPM2.DocEntry = '" +
                        dpDocEntry + "' and InvType = 204 and OVPM.Canceled = 'N'"; // ეს არის Incoming Paymentebis docentry -ები


                    recSet2.DoQuery("select   OVPM.DocEntry, SUM(VPM2.AppliedFC) as 'AppliedFC', OVPM.DocDate from OVPM inner join VPM2 on " +
                                    "OVPM.DocEntry = VPM2.DocNum inner join ORTT on OVPM.DocDate = ORTT.RateDate where OVPM.DocEntry in (" + ORCTDocEntrys + ") and ORTT.Currency = '" + data.DocCurrency + "' group by OVPM.DocEntry, OVPM.DocDate  ");

                    List<Tuple<int, DateTime, decimal>> sumPayments = new List<Tuple<int, DateTime, decimal>>();
                    while (!recSet2.EoF)
                    {
                        int OCRTDocEntry = int.Parse(recSet2.Fields.Item("DocEntry").Value.ToString());
                        DateTime DocDate = DateTime.Parse(recSet2.Fields.Item("DocDate").Value.ToString());
                        decimal appliedAmountFc = decimal.Parse(recSet2.Fields.Item("AppliedFC").Value.ToString());
                        sumPayments.Add(new Tuple<int, DateTime, decimal>(OCRTDocEntry, DocDate, appliedAmountFc));
                        recSet2.MoveNext();
                    }


                    decimal weightedAvarageForPayment = sumPayments.Sum(doc => (decimal)UiManager.GetCurrencyRate(data.DocCurrency, doc.Item2, _comp) * doc.Item3) / sumPayments.Sum(doc => doc.Item3);

                    //decimal WeightedRate = LCSum / FCSum;

                    paidAmountDpLc += weightedAvarageForPayment * decimal.Parse(downpayment.First().Value);
                    paidAmountDpFc += decimal.Parse(downpayment.First().Value);



                }


            }

            _Logger.Info($"Before Calculation Weighted Rate  paidAmountDpLc = {paidAmountDpLc} paidAmountDpFc = {paidAmountDpFc} totalInv = {data.TotalInv}" +
                         $"GlobalRate = {data.GlobalRate} is Calculated  = {data.IsCalculated}");
            CalculateWaightedRate(data, paidAmountDpLc, paidAmountDpFc);
            _Logger.Info($"After Calculation Weighted Rate  paidAmountDpLc = {paidAmountDpLc} paidAmountDpFc = {paidAmountDpFc} totalInv = {data.TotalInv}" +
                         $"GlobalRate = {data.GlobalRate} is Calculated  = {data.IsCalculated}");



        }

        private static void CalculateWaightedRate(decimal totalInvFc, /*ref bool isRateCalculated,*/ ref string globalRate,
            decimal ratInv, ref decimal paidAmountDpLc, decimal paidAmountDpFc)
        {
            if (totalInvFc == paidAmountDpFc)
            {
                var rate = paidAmountDpLc / totalInvFc;
                globalRate = rate.ToString();
                //isRateCalculated = true;
            }
            else if (totalInvFc > paidAmountDpFc)
            {
                var dif = (totalInvFc - paidAmountDpFc) * ratInv; //invocie Open AmountFC
                paidAmountDpLc += dif;
                var rate = paidAmountDpLc / totalInvFc;
                //isRateCalculated = true;

                globalRate = Math.Round(rate, 6).ToString();
            }
        }
        private static void CalculateWaightedRate(DataForCalculationRate data,
             decimal paidAmountDpLc, decimal paidAmountDpFc)
        {
            if (data.TotalInv == paidAmountDpFc)
            {
                decimal rate = paidAmountDpLc / data.TotalInv;
                data.GlobalRate = Math.Round(rate, 6).ToString();
                data.IsCalculated = true;
            }
            else if (data.TotalInv > paidAmountDpFc)
            {
                var dif = (data.TotalInv - paidAmountDpFc) * data.RateInv; //invocie Open AmountFC
                paidAmountDpLc += dif;
                decimal rate = paidAmountDpLc / data.TotalInv;
                data.GlobalRate = Math.Round(rate, 6).ToString();
                data.IsCalculated = true;
            }
        }

        public static decimal ExchangeRateCorrectionDi(decimal netAmountToDrow, decimal totalInv, decimal ratInv,
            int downPaymentDocEntry, string docCurrency, SAPbobsCOM.Company _comp)
        {
            decimal paidAmountDpLc = 0m; // A/R DownPayment - ში არჩეული თანხა ლოკალურ ვალუტაში //Net AmountFC To Drow * Rate
            decimal paidAmountDpFc = 0m; //  A/R DownPayment - ში არჩეული თანხა FC //Net AmountFC To Drow 

            var recSetTransferDocEntry =
                (SAPbobsCOM.Recordset)_comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            var recSerTranferRate =
                (SAPbobsCOM.Recordset)_comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            recSetTransferDocEntry.DoQuery(
                "select ORCT.DocEntry from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum " +
                "where RCT2.DocEntry = '" + downPaymentDocEntry + "' and InvType = 203 and ORCT.Canceled = 'N'");

            decimal sendRate = 0;
            if (recSetTransferDocEntry.RecordCount == 0)
            {
                return 0;
            }
            if (recSetTransferDocEntry.RecordCount == 1)
            {

                var ORCTDocEntry = recSetTransferDocEntry.Fields.Item("DocEntry").Value.ToString();

                recSerTranferRate.DoQuery(
                    "select ORCT.TrsfrSum , RCT2.AppliedFC, RCT2.DocRate from ORCT inner join RCT2 on " +
                    "ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry = '" + ORCTDocEntry + "'");

                decimal transferSumLc = decimal.Parse(recSerTranferRate.Fields.Item("TrsfrSum").Value.ToString());

                decimal appliedAmountFcSum = 0;

                while (!recSerTranferRate.EoF)
                {
                    appliedAmountFcSum += decimal.Parse(recSerTranferRate.Fields.Item("AppliedFC").Value.ToString());
                    recSerTranferRate.MoveNext();
                }

                if (appliedAmountFcSum == 0)
                {
                    return 0;
                }

                sendRate = transferSumLc / appliedAmountFcSum;
            }

            else
            {
                string ORCTDocEntrys = string.Empty;
                Recordset recSet2 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                ORCTDocEntrys =
                    "select ORCT.DocEntry from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where RCT2.DocEntry = '" +
                    downPaymentDocEntry +
                    "' and InvType = 203 and ORCT.Canceled = 'N'"; // ეს არის Incoming Paymentebis docentry -ები



                recSet2.DoQuery(
                    "select ORCT.DocEntry, avg(ORCT.TrsfrSum) as 'TrsfrSum' , SUM(RCT2.AppliedFC) as 'AppliedFC' from ORCT inner join RCT2 on " +
                    "ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry in (" + ORCTDocEntrys +
                    ") group by ORCT.DocEntry");
                // აქ მოგვაქვს ინფორმაცია გადახდების მიხედვით სრუტლი თანხა LC - ში დოკუმენტის ნომერი და გადახდილი თანხა უცხოურ ვალუტაში

                List<Tuple<int, decimal, decimal>> sumPayments = new List<Tuple<int, decimal, decimal>>();
                while (!recSet2.EoF)
                {
                    int OCRTDocEntry = int.Parse(recSet2.Fields.Item("DocEntry").Value.ToString());
                    decimal appliedAmountLc = decimal.Parse(recSet2.Fields.Item("TrsfrSum").Value.ToString());
                    decimal appliedAmountFc = decimal.Parse(recSet2.Fields.Item("AppliedFC").Value.ToString());
                    sumPayments.Add(new Tuple<int, decimal, decimal>(OCRTDocEntry, appliedAmountLc, appliedAmountFc));
                    recSet2.MoveNext();
                }








                //აქ გვაქვს  ლოკალულ ვალუტაში გატარებილი დოკუმენტების ჯამი და ნომერი რომელიც უნდა გამოაკლდეს გადახდის სრულ თანხას LC- ში

                Recordset recSet4 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                recSet4.DoQuery(
                    "select DocEntry, SUM(LcPrices) as SumLCPayments from ( select  ORCT.DocEntry as 'DocEntry',  SUM(case when AppliedFC = 0 then RCT2.SumApplied else 0 end ) as 'LcPrices' from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry in (" +
                    ORCTDocEntrys +
                    ") group by  RCT2.SumApplied , ORCT.DocEntry ) LcPricesTable group by DocEntry");

                Dictionary<string, decimal> DocumentLcPriceSums = new Dictionary<string, decimal>();
                while (!recSet4.EoF)
                {
                    string OCRTDocEntry = recSet4.Fields.Item("DocEntry").Value.ToString();
                    decimal SumLCPayments = decimal.Parse(recSet4.Fields.Item("SumLCPayments").Value.ToString());
                    DocumentLcPriceSums.Add(OCRTDocEntry, SumLCPayments);
                    recSet4.MoveNext();

                    // აქ არის ლოკალურ ვალუტაში გატარებულ დოკუმენტებზე გადახდილი სტრული თანხა  
                }


                Dictionary<string, decimal> rateByDocuments = new Dictionary<string, decimal>();

                List<XContainer> DocsWithRateAndValue = new List<XContainer>();

                foreach (var tuple in sumPayments)
                {
                    var rate = (tuple.Item2 - DocumentLcPriceSums[tuple.Item1.ToString()]) / tuple.Item3;
                    var paymentDocEntry = tuple.Item1.ToString();

                    rateByDocuments.Add(paymentDocEntry, rate);

                    DocsWithRateAndValue.Add(new XContainer()
                    {
                        CurrRate = rate,
                        OrctDocEntry = paymentDocEntry
                    });
                }
                // აქ გადახდის სრულ თანხას ვაკლებ ლოკალურ ვალუტაში გადახდილი დოკუმენტის ჯამს და ვყოფ სრულ თანხაზე უცხოურ ვალუტაში


                //while (!recSet2.EoF)
                //{
                //    decimal rate = (decimal.Parse(recSet2.Fields.Item("TrsfrSum").Value.ToString()) - DocumentLcPriceSums[recSet2.Fields.Item("DocEntry").Value.ToString()]) /
                //            decimal.Parse(recSet2.Fields.Item("AppliedFC").Value.ToString());
                //    string paymentDocEntry = recSet2.Fields.Item("DocEntry").Value.ToString();
                //    rateByDocuments.Add(paymentDocEntry, rate);
                //    recSet2.MoveNext();

                //}

                Recordset recSet3 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                recSet3.DoQuery(
                    "select ORCT.DocEntry, RCT2.DocEntry as 'DpDocEntry',    RCT2.AppliedFC from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry  in ( " +
                    ORCTDocEntrys + ") and RCT2.DocEntry = '" + downPaymentDocEntry + "' and InvType = 203");


                Dictionary<string, decimal> dPIncomingPaymentShareAmountFc = new Dictionary<string, decimal>();
                while (!recSet3.EoF)
                {
                    decimal AppliedFcbyDp = decimal.Parse(recSet3.Fields.Item("AppliedFC").Value.ToString());
                    string PaymentDocEntry = recSet3.Fields.Item("DocEntry").Value.ToString();
                    dPIncomingPaymentShareAmountFc.Add(PaymentDocEntry, AppliedFcbyDp);

                    DocsWithRateAndValue.Where(z => z.OrctDocEntry == PaymentDocEntry).ToList()
                        .ForEach(s => s.AmountFC = AppliedFcbyDp);


                    recSet3.MoveNext();
                }

                //var rata = DocsWithRateAndValue.

                decimal LCSum = DocsWithRateAndValue.Sum(doc => doc.AmountFC * doc.CurrRate);
                decimal FCSum = DocsWithRateAndValue.Sum(doc => doc.AmountFC);
                sendRate = LCSum / FCSum;



            }



            paidAmountDpLc += sendRate * netAmountToDrow;
            paidAmountDpFc += netAmountToDrow;

            string globalRate = string.Empty;
            CalculateWaightedRate(totalInv, /*ref isLoss,*/ ref globalRate, ratInv, ref paidAmountDpLc, paidAmountDpFc);

            return decimal.Parse(globalRate);

        }

        public static void AddJournalEntryCredit(SAPbobsCOM.Company _comp, string creditCode, string debitCode,
            double amount, int series, string reference, string code, DateTime DocDate, int BPLID = 235)
        {

            SAPbobsCOM.JournalEntries vJE =
                (SAPbobsCOM.JournalEntries)_comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

            vJE.ReferenceDate = DocDate;
            vJE.DueDate = DocDate;
            vJE.TaxDate = DocDate;

            vJE.Memo = "Income Correction -   Invoice " + reference;
            //vJE.TransactionCode = "13";
            vJE.Reference = reference;
            vJE.TransactionCode = "1";
            vJE.Series = series;
            vJE.Lines.BPLID = BPLID; //branch
            vJE.Lines.Debit = amount;
            vJE.Lines.Credit = 0;
            vJE.Lines.AccountCode = debitCode;
            vJE.Lines.FCCredit = 0;
            vJE.Lines.FCDebit = 0;
            vJE.Lines.Add();

            vJE.Lines.BPLID = BPLID;
            vJE.Lines.ControlAccount = creditCode;
            vJE.Lines.ShortName = code;
            vJE.Lines.Debit = 0;
            vJE.Lines.Credit = amount;
            vJE.Lines.FCCredit = 0;
            vJE.Lines.FCDebit = 0;

            vJE.Lines.Add();
            int i = vJE.Add();
            if (i == 0)
            {
                return;
            }
            else
            {
                string des = _comp.GetLastErrorDescription();
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(des + " Invoice " + reference, BoMessageTime.bmt_Short);
            }
        }

        public static void AddJournalEntry(SAPbobsCOM.Company _comp, string creditCode, string debitCode,
            double amount, int series, string reference, string code, DateTime DocDate, int BPLID = 235)
        {
            SAPbobsCOM.JournalEntries vJE =
                (SAPbobsCOM.JournalEntries)_comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
            vJE.ReferenceDate = DocDate;
            vJE.DueDate = DocDate;
            vJE.TaxDate = DocDate;
            vJE.Memo = "Income Correction -   Invoice " + reference;
            //vJE.TransactionCode = "13";
            vJE.Reference = reference;
            vJE.Series = series;
            vJE.TransactionCode = "1";

            vJE.Lines.BPLID = BPLID; //branch
            vJE.Lines.Credit = amount;
            vJE.Lines.Debit = 0;
            vJE.Lines.AccountCode = creditCode;
            vJE.Lines.FCCredit = 0;
            vJE.Lines.FCDebit = 0;
            vJE.Lines.Add();

            vJE.Lines.BPLID = BPLID;
            vJE.Lines.AccountCode = debitCode;
            vJE.Lines.ShortName = code;
            vJE.Lines.Credit = 0;
            vJE.Lines.Debit = amount;
            vJE.Lines.FCCredit = 0;
            vJE.Lines.FCDebit = 0;
            // vJE.Series = 17;
            vJE.Lines.Add();
            int i = vJE.Add();
            if (i == 0)
            {
                return;
            }
            else
            {
                string des = _comp.GetLastErrorDescription();
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(des + " Invoice " + reference, BoMessageTime.bmt_Short);
            }
        }


        public static void AddJournalEntryNegative(SAPbobsCOM.Company _comp, string creditCode, string debitCode,
            double amount, int series, string reference, string code, DateTime DocDate, bool isLoss, int BPLID = 235)
        {
            SAPbobsCOM.JournalEntries vJE =
                (SAPbobsCOM.JournalEntries)_comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
            vJE.ReferenceDate = DocDate;
            vJE.DueDate = DocDate;
            vJE.TaxDate = DocDate;
            vJE.Memo = "Income Correction - Invoice " + reference;
            vJE.Reference = reference;
            vJE.Series = series;
            vJE.TransactionCode = "1";

            vJE.Lines.BPLID = BPLID; //branch
            if (isLoss)
            {
                vJE.Lines.Credit = 0;
                vJE.Lines.Debit = -amount;
            }
            else
            {
                vJE.Lines.Credit = 0;
                vJE.Lines.Debit = -amount;
                vJE.Lines.ShortName = code;
            }
            vJE.Lines.AccountCode = creditCode;
            vJE.Lines.FCCredit = 0;
            vJE.Lines.FCDebit = 0;
            vJE.Lines.Add();
            ////////////////////////////
            vJE.Lines.BPLID = BPLID;
            vJE.Lines.AccountCode = debitCode;
            if (isLoss)
            {
                vJE.Lines.Credit = -amount;
                vJE.Lines.Debit = 0;
                vJE.Lines.ShortName = code;
            }
            else
            {
                vJE.Lines.Credit = -amount;
                vJE.Lines.Debit = 0;
            }
            vJE.Lines.FCCredit = 0;
            vJE.Lines.FCDebit = 0;
            // vJE.Series = 17;
            vJE.Lines.Add();
            int i = vJE.Add();
            if (i == 0)
            {
                return;
            }
            else
            {
                string des = _comp.GetLastErrorDescription();
                _Logger.Error($"Journal Entry Not Added Coz : {des}");
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(des + " Invoice " + reference, BoMessageTime.bmt_Short);
            }
        }

        public static void AddJournalEntryDebit(SAPbobsCOM.Company _comp, string creditCode, string debitCode,
            double amount, int series, string reference, string code, DateTime DocDate, int BPLID = 235)
        {

            

            SAPbobsCOM.JournalEntries vJE =
                (SAPbobsCOM.JournalEntries)_comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
            vJE.ReferenceDate = DocDate;
            vJE.DueDate = DocDate;
            vJE.TaxDate = DocDate;
            vJE.Memo = "Income Correction -   Invoice " + reference;
            //vJE.TransactionCode = "13";
            vJE.Reference = reference;
            vJE.Series = series;
            vJE.TransactionCode = "1";

            vJE.Lines.BPLID = BPLID; //branch
            vJE.Lines.Credit = amount;
            vJE.Lines.Debit = 0;
            vJE.Lines.AccountCode = creditCode;
            vJE.Lines.FCCredit = 0;
            vJE.Lines.FCDebit = 0;
            vJE.Lines.Add();

            vJE.Lines.BPLID = BPLID;
            vJE.Lines.AccountCode = debitCode;
            vJE.Lines.ShortName = code;
            vJE.Lines.Credit = 0;
            vJE.Lines.Debit = amount;
            vJE.Lines.FCCredit = 0;
            vJE.Lines.FCDebit = 0;
            // vJE.Series = 17;
            vJE.Lines.Add();
            int i = vJE.Add();
            if (i == 0)
            {
                return;
            }
            else
            {
                string des = _comp.GetLastErrorDescription();
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(des + " Invoice " + reference, BoMessageTime.bmt_Short);
            }
        }

        public static Dictionary<int, string> FormToTransId = new Dictionary<int, string>()
        {
            {133, "13"},
            {141, "18"}
        };

        public static void CorrectionJournalEntryUI(SAPbobsCOM.Company _comp, int FormType, string businesPartnerCardCode, string apllaidDp, string docNumber, string bplName, string ExchangeGain, string ExchangeLoss, DateTime docDate)
        {
            Recordset recSet = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            recSet.DoQuery("select DebPayAcct from OCRD where CardCode = '" + businesPartnerCardCode + "'");
            string BpControlAcc = recSet.Fields.Item("DebPayAcct").Value.ToString();
            if (!string.IsNullOrWhiteSpace(apllaidDp))
            {
                var objRS = (SAPbobsCOM.Recordset)(_comp).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objRS.DoQuery(@"select * from OJDT where baseRef = " + docNumber + " and TransType = " +
                              FormToTransId[FormType] + "");
                objRS.MoveFirst();
                var x = objRS.Fields.Item("TransType").Value.ToString();
                if (objRS.Fields.Item("TransType").Value.ToString() != "13" &&
                    objRS.Fields.Item("TransType").Value.ToString() != "18")
                {
                    objRS.MoveNext();
                }
                var transID = objRS.Fields.Item("TransId").Value.ToString();
                objRS.DoQuery(@"select * from JDT1 where TransId = " + transID);
                var objRS234 = (SAPbobsCOM.Recordset)(_comp).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (objRS.Fields.Item("TransType").Value.ToString() != "13")
                {
                    objRS234.DoQuery("select  BPLId from OPCH where BPLName = '" + bplName + "'");
                }
                else if (objRS.Fields.Item("TransType").Value.ToString() != "18")
                {
                    objRS234.DoQuery("select  BPLId from OINV where BPLName = '" + bplName + "'");
                }

                int bplID = Convert.ToInt32(objRS234.Fields.Item("BPLId").Value);

                Recordset recSet12 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                recSet12.DoQuery("select Series from NNM1 where ObjectCode = 30 and Locked = 'N' and BPLId is  null");
                int series = int.Parse(recSet12.Fields.Item("Series").Value.ToString());

                if (bplID == 0)
                {
                    bplID = 235;
                }

                while (!objRS.EoF)
                {
                    var account = objRS.Fields.Item("Account").Value.ToString();

                    if (FormType.ToString() == "133")
                    {
                        if (account == ExchangeGain)
                        {
                            AddJournalEntryNegative(_comp, BpControlAcc, ExchangeGain,
                                Convert.ToDouble(objRS.Fields.Item("Credit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate, false,
                                bplID);

                        }
                        else if (account == ExchangeLoss)
                        {
                            AddJournalEntryNegative(_comp, ExchangeLoss, BpControlAcc,
                                Convert.ToDouble(objRS.Fields.Item("Debit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate, true, bplID);
                        }
                    }
                    else if (FormType.ToString() == "141")
                    {
                        if (account == ExchangeGain)
                        {
                            AddJournalEntryNegative(_comp, BpControlAcc,  ExchangeGain,
                                Convert.ToDouble(objRS.Fields.Item("Credit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate, false,
                                bplID);
                        }
                        else if (account == ExchangeLoss)
                        {
                            AddJournalEntryNegative(_comp, ExchangeLoss, BpControlAcc,
                                Convert.ToDouble(objRS.Fields.Item("Debit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate,true,  bplID);
                        }
                    }

                    objRS.MoveNext();
                }
            }

        }

        public static void CorrectionJournalEntryDI(SAPbobsCOM.Company _comp, int FormType, string businesPartnerCardCode, string docNumber, string bplName, string ExchangeGain, string ExchangeLoss, DateTime docDate)
        {
            Recordset recSet = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            recSet.DoQuery("select DebPayAcct from OCRD where CardCode = '" + businesPartnerCardCode + "'");
            string BpControlAcc = recSet.Fields.Item("DebPayAcct").Value.ToString();

            var objRS = (SAPbobsCOM.Recordset)(_comp).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            objRS.DoQuery(@"select * from OJDT where baseRef = " + docNumber + " and TransType = " +
                          FormToTransId[FormType] + "");
            objRS.MoveFirst();
            var x = objRS.Fields.Item("TransType").Value.ToString();

            if (objRS.Fields.Item("TransType").Value.ToString() != "13" &&
                objRS.Fields.Item("TransType").Value.ToString() != "18")
            {
                objRS.MoveNext();
            }
            var transID = objRS.Fields.Item("TransId").Value.ToString();
            objRS.DoQuery(@"select * from JDT1 where TransId = " + transID);
            var objRS234 = (SAPbobsCOM.Recordset)(_comp).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            if (objRS.Fields.Item("TransType").Value.ToString() != "13")
            {
                objRS234.DoQuery("select  BPLId from OPCH where BPLName = '" + bplName + "'");
            }
            else if (objRS.Fields.Item("TransType").Value.ToString() != "18")
            {
                objRS234.DoQuery("select  BPLId from OINV where BPLName = '" + bplName + "'");
            }

            int bplID = Convert.ToInt32(objRS234.Fields.Item("BPLId").Value);

            Recordset recSet12 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            recSet12.DoQuery("select Series from NNM1 where ObjectCode = 30 and Locked = 'N' and BPLId is  null");
            int series = int.Parse(recSet12.Fields.Item("Series").Value.ToString());

            if (bplID == 0)
            {
                bplID = 235;
            }

            while (!objRS.EoF)
            {
                var account = objRS.Fields.Item("Account").Value.ToString();

                if (FormType.ToString() == "133")
                {
                    if (account == ExchangeGain)
                    {
                        AddJournalEntryCredit(_comp,  BpControlAcc, ExchangeGain,
                            Convert.ToDouble(objRS.Fields.Item("Credit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate,
                            bplID);

                    }
                    else if (account == ExchangeLoss)
                    {
                        AddJournalEntryDebit(_comp, ExchangeLoss, BpControlAcc, 
                            Convert.ToDouble(objRS.Fields.Item("Debit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate, bplID);
                    }
                }
                else if (FormType.ToString() == "141")
                {
                    if (account == ExchangeGain)
                    {
                        AddJournalEntryCredit(_comp,  BpControlAcc, ExchangeGain,
                            Convert.ToDouble(objRS.Fields.Item("Credit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate,
                            bplID);
                    }
                    else if (account == ExchangeLoss)
                    {
                        AddJournalEntryDebit(_comp, ExchangeLoss, BpControlAcc,
                            Convert.ToDouble(objRS.Fields.Item("Debit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate, bplID);
                    }
                }

                objRS.MoveNext();
            }


        }


    }
}

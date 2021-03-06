
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DownPaymentLogic.Classes;
using SAPApi;
using SAPbobsCOM;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using Application = SAPbouiCOM.Framework.Application;
using Company = SAPbouiCOM.Company;

namespace jo0urnaltest
{

    [FormAttribute("133", "A_R Invoice.b1f")]
    class A_R_Invoice : SystemFormBase
    {

        public A_R_Invoice()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.PressedAfter +=
                new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button0.PressedBefore +=
                new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.Button0_PressedBefore);
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_0000").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1000").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("4").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("213").Specific));
            this.Button1.PressedAfter +=
                new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.Button1.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button1_ClickBefore);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.CloseAfter += new SAPbouiCOM.Framework.FormBase.CloseAfterHandler(this.Form_CloseAfter);
            this.ActivateAfter += new SAPbouiCOM.Framework.FormBase.ActivateAfterHandler(this.Form_ActivateAfter);
            this.DataAddAfter += new DataAddAfterHandler(this.Form_DataAddAfter);

        }

        private SAPbouiCOM.Button Button0;
        private static string id;
        private static string down;


        private void OnCustomInitialize()
        {
            Program.BusinesPartnerName = string.Empty;
            Program.bplName = string.Empty;
            DataForCalculationRate = new DataForCalculationRate();
            DownPaymentsForInvFormIds = new List<Dictionary<string, string>>();
            _logger = new SimpleLogger();
        }

        private SimpleLogger _logger;

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText1;
        private Button Button1;



        public static List<Dictionary<string, string>> DownPaymentsForInvFormIds { get; set; }


        public static Dictionary<string, string> DownPaymentsForInvFormId = new Dictionary<string, string>();

        private string _formUIdInv;
        private string _formUIdDps;
        //   private string globalRate = string.Empty;

        private void Button1_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            ((Matrix)(Application.SBO_Application.Forms.ActiveForm.Items.Item("38").Specific)).Columns
                .Item("15").Cells.Item(1).Click();

            _formUIdInv = Application.SBO_Application.Forms.ActiveForm.UDFFormUID;
            DataForCalculationRate.FormUIdInv = _formUIdInv;
            BubbleEvent = true;

            Form arInoviceForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;

            try
            {
                DataForCalculationRate.BusinesPartnerName = ((SAPbouiCOM.EditText)(arInoviceForm.Items
                    .Item("4").Specific)).Value;
                try
                {

                    DataForCalculationRate.BplName = ((SAPbouiCOM.ComboBox)(arInoviceForm.Items //es iwvevs catches
                        .Item("2001").Specific)).Selected.Description;
                }
                catch (Exception e)
                {

                    DataForCalculationRate.BplName = "235";
                    _logger.Error($"{e.Message} No Branch");
                }

                var docNumInvItm = arInoviceForm.Items.Item("8");
                var docNumInvEditText = (SAPbouiCOM.EditText)docNumInvItm.Specific;
                DataForCalculationRate.DocNum = docNumInvEditText.Value;

                var cardCodeItem = arInoviceForm.Items.Item("4");
                var cardCodeEditText = (SAPbouiCOM.EditText)cardCodeItem.Specific;
                DataForCalculationRate.CardCode = cardCodeEditText.Value;

                var totalDownPaymentItem = arInoviceForm.Items.Item("204");
                var totalDownPaymentEditText = (SAPbouiCOM.EditText)totalDownPaymentItem.Specific;
                DataForCalculationRate.DownPaymentAmount = totalDownPaymentEditText.Value;
                DataForCalculationRate.FormTypex = "133";

                var txtTotalWithCurr =
                    (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("22").Specific); //totalInv before discount
                var vatWithCurr = (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("27").Specific);
                var discountWithCurr = (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("42").Specific);

                string txtTotal = txtTotalWithCurr.Value.Split(' ')[0] == string.Empty ? "0" : txtTotalWithCurr.Value.Split(' ')[0];
                string vat = vatWithCurr.Value.Split(' ')[0] == string.Empty ? "0" : vatWithCurr.Value.Split(' ')[0];
                string discount = discountWithCurr.Value.Split(' ')[0] == string.Empty
                    ? "0"
                    : discountWithCurr.Value.Split(' ')[0];


                SBObob sbObob = (SBObob)Program._comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                string currency = ((ComboBox)SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Items
                    .Item("63").Specific).Value;
                DataForCalculationRate.DocCurrency = currency;
                string postingDateString = ((EditText)arInoviceForm.Items.Item("10").Specific).Value;
                DateTime postingDate = DateTime.ParseExact(postingDateString, "yyyyMMdd", null);
                DataForCalculationRate.PostingDate = postingDate;
                decimal currencyValue;
                if (currency != sbObob.GetLocalCurrency().Fields.Item(0).Value.ToString())
                {
                    currencyValue =
                        Math.Round(
                            decimal.Parse(sbObob.GetCurrencyRate(currency, postingDate).Fields.Item(0).Value
                                .ToString()), 6);
                    ((SAPbouiCOM.EditText)(arInoviceForm.Items.Item("Item_0000").Specific)).Value =
                        currencyValue.ToString();
                }
                else
                {
                    currencyValue = 1;
                    ((SAPbouiCOM.EditText)(arInoviceForm.Items.Item("Item_0000").Specific)).Value =
                        currencyValue.ToString();
                }

                try
                {
                    DataForCalculationRate.TotalInv = decimal.Parse(txtTotal) - decimal.Parse(discount) +
                                                      decimal.Parse(String.IsNullOrEmpty(vat) ? "0" : vat);

                    var docDate = DateTime.ParseExact(
                        ((EditText)(arInoviceForm.Items.Item("10").Specific)).Value.ToString(),
                        "yyyyMMdd", CultureInfo.InvariantCulture);
                    DataForCalculationRate.RateInv = decimal.Parse(UiManager.GetCurrencyRate(DataForCalculationRate.DocCurrency, docDate, Program._comp).ToString());

                    var txtRate = (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("Item_0000").Specific);
                    DataForCalculationRate.RateInv = decimal.Parse(txtRate.Value.Replace(",", "."));
                }
                catch (Exception e)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(e.Message,
                        BoMessageTime.bmt_Short, true);
                    SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Close();
                    _logger.Error(e.Message);
                }
            }
            catch (Exception e)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(e.Message, BoMessageTime.bmt_Short,
                    true);
                _logger.Error(e.Message);
                //branch araa
            }
        }



        private void Form_CloseAfter(SBOItemEventArg pVal)
        {
            Program.BusinesPartnerName = string.Empty;
            Program.bplName = string.Empty;
            SharedClass.ListOfDataForCalculationRates.Remove(DataForCalculationRate);
        }

        public DataForCalculationRate DataForCalculationRate { get; set; }
        private void Button0_PressedBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {

            BubbleEvent = true;
            if (pVal.FormMode != 3)
            {
                return;
            }
            Form arInoviceForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;

            try
            {


                DataForCalculationRate.BusinesPartnerName = ((SAPbouiCOM.EditText)(arInoviceForm.Items
                    .Item("4").Specific)).Value;
                try
                {

                    DataForCalculationRate.BplName = ((SAPbouiCOM.ComboBox)(arInoviceForm.Items //es iwvevs catches
                        .Item("2001").Specific)).Selected.Description;
                }
                catch (Exception)
                {

                    DataForCalculationRate.BplName = "235";
                }

                //PostingDate = DateTime.ParseExact(
                //    ((EditText)(arInoviceForm.Items.Item("10").Specific)).Value,
                //    "yyyyMMdd", CultureInfo.InvariantCulture);
                //RateInv = decimal.Parse(UiManager.GetCurrencyRate(DocCurrency, PostingDate, Program._comp).ToString());


                var docNumInvItm = arInoviceForm.Items.Item("8");
                var docNumInvEditText = (SAPbouiCOM.EditText)docNumInvItm.Specific;
                DataForCalculationRate.DocNum = docNumInvEditText.Value;

                var cardCodeItem = arInoviceForm.Items.Item("4");
                var cardCodeEditText = (SAPbouiCOM.EditText)cardCodeItem.Specific;
                DataForCalculationRate.CardCode = cardCodeEditText.Value;

                var totalDownPaymentItem = arInoviceForm.Items.Item("204");
                var totalDownPaymentEditText = (SAPbouiCOM.EditText)totalDownPaymentItem.Specific;
                DataForCalculationRate.DownPaymentAmount = totalDownPaymentEditText.Value;
                DataForCalculationRate.FormTypex = "133";

                var txtTotalWithCurr =
                    (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("22").Specific); //totalInv before discount
                var vatWithCurr = (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("27").Specific);
                var discountWithCurr = (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("42").Specific);

                string txtTotal = txtTotalWithCurr.Value.Split(' ')[0];
                string vat = vatWithCurr.Value.Split(' ')[0];
                string discount = discountWithCurr.Value.Split(' ')[0] == string.Empty
                    ? "0"
                    : discountWithCurr.Value.Split(' ')[0];


                SBObob sbObob = (SBObob)Program._comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                string currency = ((ComboBox)SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Items
                    .Item("63").Specific).Value;
                DataForCalculationRate.DocCurrency = currency;
                string postingDateString = ((EditText)arInoviceForm.Items.Item("10").Specific).Value;
                DateTime postingDate = DateTime.ParseExact(postingDateString, "yyyyMMdd", null);
                DataForCalculationRate.PostingDate = postingDate;

                decimal currencyValue;
                if (currency != sbObob.GetLocalCurrency().Fields.Item(0).Value.ToString())
                {
                    currencyValue =
                        Math.Round(
                            decimal.Parse(sbObob.GetCurrencyRate(currency, postingDate).Fields.Item(0).Value
                                .ToString()), 6);
                    ((SAPbouiCOM.EditText)(arInoviceForm.Items.Item("Item_0000").Specific)).Value =
                        currencyValue.ToString();
                }
                else
                {
                    currencyValue = 1;
                    ((SAPbouiCOM.EditText)(arInoviceForm.Items.Item("Item_0000").Specific)).Value =
                        currencyValue.ToString();
                }


                try
                {
                    if (string.IsNullOrWhiteSpace(txtTotal))
                    {
                        return;
                    }
                    DataForCalculationRate.TotalInv = decimal.Parse(txtTotal) - decimal.Parse(discount) +
                               decimal.Parse(String.IsNullOrEmpty(vat) ? "0" : vat);

                    var docDate = DateTime.ParseExact(
                        ((EditText)(arInoviceForm.Items.Item("10").Specific)).Value.ToString(),
                        "yyyyMMdd", CultureInfo.InvariantCulture);
                    DataForCalculationRate.RateInv = decimal.Parse(UiManager.GetCurrencyRate(DataForCalculationRate.DocCurrency, docDate, Program._comp).ToString());

                    var txtRate = (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("Item_0000").Specific);
                    DataForCalculationRate.RateInv = decimal.Parse(txtRate.Value.Replace(",", "."));
                }
                catch (Exception e)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(e.Message,
                        BoMessageTime.bmt_Short, true);
                    SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Close();
                }
            }
            catch (Exception e)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(e.Message, BoMessageTime.bmt_Short,
                    true);
                //branch araa
            }


            try
            {
                DownPaymentLogic.DownPaymentLogic.ExchangeRateCorrectionUi(DataForCalculationRate, Program._comp);
                if (DataForCalculationRate.IsCalculated)
                {
                    Form invoiceForm = Application.SBO_Application.Forms.ActiveForm;
                    var txtRate = (SAPbouiCOM.EditText)(invoiceForm.Items.Item("64").Specific);
                    txtRate.Value = Math.Round(decimal.Parse(DataForCalculationRate.GlobalRate), 6).ToString();
                    DataForCalculationRate.IsCalculated = false;
                }

            }
            catch (Exception e)
            {

            }

        }





        private void Button0_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (pVal.ActionSuccess && pVal.FormMode == 3)
            {
                Documents invoice =
                    (SAPbobsCOM.Documents)Program._comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                invoice.GetByKey(DataForCalculationRate.DocEntry);
                DataForCalculationRate.DocNum = invoice.DocNum.ToString();
                DownPaymentLogic.DownPaymentLogic.CorrectionJournalEntryUI(Program._comp, 133, DataForCalculationRate.CardCode,
                    DataForCalculationRate.DownPaymentAmount, DataForCalculationRate.DocNum, DataForCalculationRate.BplName, Program.ExchangeGain, Program.ExchangeLoss, DataForCalculationRate.PostingDate);
            }

        }

        private void Button1_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {

            Form downPaymentToDrowForm = Application.SBO_Application.Forms.ActiveForm;
            

            _formUIdDps = downPaymentToDrowForm.UDFFormUID;
            DataForCalculationRate.FormUIdDps = _formUIdDps;

            if (!SharedClass.ListOfDataForCalculationRates.Any())
            {
                SharedClass.ListOfDataForCalculationRates.Add(DataForCalculationRate);
            }
            else
            {

                var x1 = SharedClass.ListOfDataForCalculationRates.First();
                x1 = DataForCalculationRate;
                SharedClass.ListOfDataForCalculationRates.Remove(x1);
                SharedClass.ListOfDataForCalculationRates.Add(DataForCalculationRate);
            }



        }


        private void Form_ActivateAfter(SBOItemEventArg pVal)
        {

            if (DataForCalculationRate.IsCalculated)
            {
                Form invoiceForm = Application.SBO_Application.Forms.ActiveForm;
                var txtRate = (SAPbouiCOM.EditText)(invoiceForm.Items.Item("64").Specific);
                txtRate.Value = Math.Round(decimal.Parse(DataForCalculationRate.GlobalRate), 6).ToString();
                DataForCalculationRate.IsCalculated = false;
                _logger.Info($"Activation After A/R Invoice IsCalculated = {DataForCalculationRate.IsCalculated} Rate = {DataForCalculationRate.GlobalRate}  ");
            }
        }

        private void Form_DataAddAfter(ref BusinessObjectInfo pVal)
        {
            try
            {
                string xmlObjectKey = pVal.ObjectKey;
                XElement xmlnew = XElement.Parse(xmlObjectKey);
                int docEntry = int.Parse(xmlnew.Element("DocEntry").Value);
                DataForCalculationRate.DocEntry = docEntry;
            }
            catch (Exception e)
            {

            }

        }
    }
}


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

namespace jo0urnaltest
{

    [FormAttribute("141", "A_P Invoice.b1f")]
    class A_P_Invoice : SystemFormBase
    {
        public A_P_Invoice()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_0000").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1000").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("70").Specific));
            //     this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            //     this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("4").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("213").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.Button1.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button1_ClickBefore);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);
            this.Button2.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.Button2_PressedBefore);
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


        private SAPbouiCOM.EditText EditText0;
        private SimpleLogger _logger;
        private void OnCustomInitialize()
        {
            Program.BusinesPartnerName = string.Empty;
            Program.bplName = string.Empty;
            DataForCalculationRate = new DataForCalculationRate();
            _logger = new SimpleLogger();
        }

        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.Button Button0;

        private void ComboBox0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {

        }

        private EditText EditText1;
        private Button Button1;

        public DataForCalculationRate DataForCalculationRate { get; set; }
        private string _formUIdInv;
        private string _formUIdDps;

        private void Button1_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            {
                ((Matrix)(Application.SBO_Application.Forms.ActiveForm.Items.Item("38").Specific)).Columns
                    .Item("15").Cells.Item(1).Click();
                _formUIdInv = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.UDFFormUID;
                DataForCalculationRate.FormUIdInv = _formUIdInv;
                BubbleEvent = true;

                Form arInoviceForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;

                try
                {


                    DataForCalculationRate.BusinesPartnerName = ((SAPbouiCOM.EditText)(arInoviceForm.Items
                        .Item("4").Specific)).Value;
                    DataForCalculationRate.BplName = ((SAPbouiCOM.ComboBox)(arInoviceForm.Items //es iwvevs catches
                        .Item("2001").Specific)).Selected.Description;


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
                    DataForCalculationRate.FormTypex = "141";

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
                    if (string.IsNullOrWhiteSpace(postingDateString))
                    {
                        Application.SBO_Application.SetStatusBarMessage("მიუთითეთ თარიღი",
                            BoMessageTime.bmt_Short, true);
                        return;
                    }
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
                    //branch araa
                    _logger.Error(e.Message);
                }




            }

        }

        private void Form_CloseAfter(SBOItemEventArg pVal)
        {
            Program.BusinesPartnerName = string.Empty;
            Program.bplName = string.Empty;
            SharedClass.ListOfDataForCalculationRates.Remove(DataForCalculationRate);
        }

        private void Button1_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            {

                Form downPaymentToDrowForm = Application.SBO_Application.Forms.ActiveForm;

                //if (downPaymentToDrowForm.TypeEx != "60511")
                //{
                //    return;
                //}
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

        }

        private Button Button2;

        private void Button2_PressedBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
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
                    DataForCalculationRate.BplName = ((SAPbouiCOM.ComboBox)(arInoviceForm.Items //es iwvevs catches
                        .Item("2001").Specific)).Selected.Description;


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
                    DataForCalculationRate.FormTypex = "141";

                    var txtTotalWithCurr =
                        (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("22").Specific); //totalInv before discount
                    var vatWithCurr = (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("27").Specific);
                    var discountWithCurr = (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("42").Specific);

                    string txtTotal = txtTotalWithCurr.Value.Split(' ')[0];
                    if (string.IsNullOrWhiteSpace(txtTotal))
                    {
                        return;
                    }
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
                    //EditText txtRate = (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("64").Specific); // BP Currency A/R Invoice  Exchange Rate
                    //txtRate.Value = DataForCalculationRate.GlobalRate;
                }
                catch (Exception e)
                {
                    _logger.Error(e.Message);
                }

            }

        }

        private void Button2_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            _logger.Info($"Before A/P Invoice Add   ActionSuccess  = {pVal.ActionSuccess} \n  FormMode = {pVal.FormMode} (3 = ADD)");
            if (pVal.ActionSuccess && pVal.FormMode == 3)
            {
                Documents invoice =
                    (SAPbobsCOM.Documents)Program._comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                invoice.GetByKey(DataForCalculationRate.DocEntry);
                DataForCalculationRate.DocNum = invoice.DocNum.ToString();
                DownPaymentLogic.DownPaymentLogic.CorrectionJournalEntryUI(Program._comp, 141, DataForCalculationRate.CardCode,
                    DataForCalculationRate.DownPaymentAmount, DataForCalculationRate.DocNum, DataForCalculationRate.BplName, Program.ExchangeGain, Program.ExchangeLoss, DataForCalculationRate.PostingDate);
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
                _logger.Info($"Activation After A/P Invoice IsCalculated = {DataForCalculationRate.IsCalculated} Rate = {DataForCalculationRate.GlobalRate}  ");
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
            catch (Exception)
            {


            }
        }
    }
}

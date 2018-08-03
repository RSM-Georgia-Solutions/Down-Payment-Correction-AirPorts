using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using DownPaymentLogic.Classes;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using Application = SAPbouiCOM.Framework.Application;
using Form = SAPbouiCOM.Form;

namespace jo0urnaltest
{
    [FormAttribute("60511", "DownPaymentsToDrow.b1f")]
    class DownPaymentsToDrow : SystemFormBase
    {
        public DownPaymentsToDrow()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);

            this.OnCustomInitialize();
        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Button Button0;

        private SimpleLogger _logger;

        private void OnCustomInitialize()
        {
            _logger = new SimpleLogger();
        }

        //public static Action action;

        private static PropertyInfo[] GetProperties(object obj)
        {
            return obj.GetType().GetProperties();
        }




        private void Button0_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
             

            Form downPaymentToDrowForm = Application.SBO_Application.Forms.ActiveForm;
            if (pVal.ActionSuccess)
            {
                Item downPaymentFormMatrix = downPaymentToDrowForm.Items.Item("6"); //Down Payment to drow
                Matrix matrix = (SAPbouiCOM.Matrix)downPaymentFormMatrix.Specific;

                List<Dictionary<string, string>> downPaymentDocEntryNetAmount = new List<Dictionary<string, string>>();

                for (int i = 1; i <= matrix.RowCount; i++)
                {
                    var checkbox = (SAPbouiCOM.CheckBox)matrix.Columns.Item("380000138").Cells.Item(i).Specific;
                    if (checkbox.Checked)
                    {
                        EditText txtMoney = (SAPbouiCOM.EditText)matrix.Columns.Item("29").Cells.Item(i).Specific; //net amount to drow//
                        EditText txtID = (SAPbouiCOM.EditText)matrix.Columns.Item("68").Cells.Item(i).Specific; //docNumber

                        downPaymentDocEntryNetAmount.Add(new Dictionary<string, string>
                        {
                            {txtID.Value, txtMoney.Value.Split(' ')[0]}
                        });
                        _logger.Info($"Net Amount To Drow = {txtMoney.Value.Split(' ')[0]}");
                    }
                }


                try
                {
                    var formCouples = SharedClass.ListOfDataForCalculationRates.First(/*x => x.FormUIdDps == downPaymentToDrowForm.UDFFormUID*/);
                    formCouples.NetAmountsForDownPayment = downPaymentDocEntryNetAmount;
                }
                catch (Exception e)
                {
                    _logger.Error($"Lambda Expression For Couple Forms {e.Message}");
                }


                try
                {
                    var x1 = SharedClass.ListOfDataForCalculationRates.First(/*x => x.FormUIdDps == downPaymentToDrowForm.UDFFormUID*/);
                    var properties = GetProperties(x1);
                    foreach (var p in properties)
                    {
                        string name = p.Name;
                        var value = p.GetValue(x1, null);
                        _logger.Info($"DataForCalculation Values {name}  = {value}");

                    }
                    DownPaymentLogic.DownPaymentLogic.ExchangeRateCorrectionUi(x1, Program._comp);
                }
                catch (Exception e)
                {
                    _logger.Fatal($"Calculation Exchange rate Filed {e.Message}" );
                }
            }
        }
    }
}

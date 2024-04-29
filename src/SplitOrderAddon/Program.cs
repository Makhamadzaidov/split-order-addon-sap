using SAPbouiCOM.Framework;
using SplitOrderAddon.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace SplitOrderAddon
{
    class Program
    {
        static SAPbobsCOM.Company oCom;
        static SAPbobsCOM.Recordset oRS;

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new Application();
                }
                else
                {
                    oApp = new Application(args[0]);
                }
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;

                oCom = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                oRS = (SAPbobsCOM.Recordset)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.FormTypeEx == "139" && pVal.BeforeAction == false && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                SAPbouiCOM.Item oButtonPurchase = oForm.Items.Add("Click", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                SAPbouiCOM.Item oTempItem = oForm.Items.Item("2");
                SAPbouiCOM.Button oPostButton = (SAPbouiCOM.Button)oButtonPurchase.Specific;

                oPostButton.Caption = "Заказ болиш";
                oButtonPurchase.Left = oTempItem.Left + oTempItem.Width + 5;
                oButtonPurchase.Top = oTempItem.Top;
                oButtonPurchase.Width = 130;
                oButtonPurchase.Height = oTempItem.Height;
                oButtonPurchase.AffectsFormMode = false;
            }

            if (pVal.FormTypeEx == "139" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "Click" && pVal.BeforeAction == false)
            {
                try
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                    SAPbouiCOM.DBDataSource matrixDT = oForm.DataSources.DBDataSources.Item("RDR1");
                    SAPbouiCOM.DBDataSource dbDataSource = Application.SBO_Application.Forms.Item((object)pVal.FormUID).DataSources.DBDataSources.Item((object)"ORDR");

                    Application.SBO_Application.StatusBar.SetSystemMessage("Начался процесс генерации ...",
                        SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                    IList<Item> items = new List<Item>();

                    Partner partner = new Partner()
                    {
                        CardCode = ((SAPbouiCOM.EditText)oForm.Items.Item("4").Specific).Value,
                        CardName = ((SAPbouiCOM.EditText)oForm.Items.Item("54").Specific).Value,
                        DocDate = DateTime.ParseExact(((SAPbouiCOM.EditText)oForm.Items.Item("10").Specific).Value, "yyyyMMdd", CultureInfo.InvariantCulture),
                        DocDueDate = DateTime.ParseExact(((SAPbouiCOM.EditText)oForm.Items.Item("12").Specific).Value, "yyyyMMdd", CultureInfo.InvariantCulture),
                        TaxDate = DateTime.ParseExact(((SAPbouiCOM.EditText)oForm.Items.Item("46").Specific).Value, "yyyyMMdd", CultureInfo.InvariantCulture)
                    };

                    SAPbouiCOM.ComboBox contactPerson = (SAPbouiCOM.ComboBox)oForm.Items.Item("85").Specific;
                    if (contactPerson.Selected != null)
                    {
                        string descrition = contactPerson.Selected.Description;
                        string value = contactPerson.Selected.Value;

                        partner.ContactPerson = descrition;
                    }

                    for (int i = 1; i <= oMatrix.RowCount; i++)
                    {
                        oMatrix.GetLineData(i);

                        if (!string.IsNullOrEmpty(matrixDT.GetValue("ItemCode", i - 1)))
                        {
                            Item item = new Item()
                            {
                                ItemCode = matrixDT.GetValue("ItemCode", i - 1),
                                ItemName = matrixDT.GetValue("Dscription", i - 1),
                                Quantity = double.Parse(matrixDT.GetValue("Quantity", i - 1).Split('.')[0] + "," + matrixDT.GetValue("Quantity", i - 1).Split('.')[1]),
                                Price = double.Parse(((SAPbouiCOM.EditText)oMatrix.Columns.Item("14").Cells.Item(i).Specific).Value.Split()[0]),
                                WarehouseCode = matrixDT.GetValue("WhsCode", i - 1),
                                DiscountPercent = double.Parse(matrixDT.GetValue("DiscPrcnt", i - 1).Split('.')[0] + "," + matrixDT.GetValue("DiscPrcnt", i - 1).Split('.')[1]),
                            };

                            oRS.DoQuery($"SELECT T0.\"U_ItmGrp\" FROM OITM T0 WHERE T0.\"ItemCode\" = '{item.ItemCode}'");

                            if (!oRS.EoF)
                            {
                                item.ItmGrp = oRS.Fields.Item("U_ItmGrp").Value.ToString();
                            }

                            items.Add(item);
                        }
                    }

                    var groupedItems = items.GroupBy(item => item.ItmGrp);
                    bool flag1 = false;

                    string docNum = dbDataSource.GetValue("DocNum", 0).ToString();
                    Console.WriteLine("Document Number: " + docNum);

                    oRS.DoQuery($"SELECT T0.\"DocEntry\" FROM ORDR T0 WHERE T0.\"DocNum\" = {docNum}");
                    var docEntry = int.Parse(oRS.Fields.Item("DocEntry").Value.ToString());

                    foreach (var product in groupedItems)
                    {
                        SAPbobsCOM.Documents salesOrder = (SAPbobsCOM.Documents)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

                        salesOrder.CardCode = partner.CardCode;
                        salesOrder.CardName = partner.CardName;
                        salesOrder.DocDate = partner.DocDate;
                        salesOrder.DocDueDate = partner.DocDueDate;
                        salesOrder.TaxDate = partner.TaxDate;
                        salesOrder.NumAtCard = docNum;
                        salesOrder.ImportFileNum = docEntry;

                        foreach (var item in product)
                        {
                            salesOrder.Lines.ItemCode = item.ItemCode;
                            salesOrder.Lines.ItemDescription = item.ItemName;
                            salesOrder.Lines.Quantity = item.Quantity;
                            salesOrder.Lines.UnitPrice = item.Price;
                            salesOrder.Lines.WarehouseCode = item.WarehouseCode;
                            salesOrder.Lines.DiscountPercent = item.DiscountPercent;
                            salesOrder.Lines.Add();
                        }

                        int status = salesOrder.Add();

                        if (status != 0)
                        {
                            flag1 = true;
                            int errorCode = oCom.GetLastErrorCode();
                            string error = oCom.GetLastErrorDescription();

                            Application.SBO_Application.StatusBar.SetSystemMessage("Ошибка при генерации", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                        else
                        {
                            Application.SBO_Application.StatusBar.SetSystemMessage($"Генерации завершена {product.Key}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        }
                    }

                    if (flag1)
                    {
                        Application.SBO_Application.StatusBar.SetSystemMessage("Произошли ошибки при генерации. Просмотрите логи чтобы их увидеть.",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                    else
                    {
                        SAPbobsCOM.Documents sboPrClose = (SAPbobsCOM.Documents)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                        sboPrClose.NumAtCard = docNum;

                        sboPrClose.GetByKey(docEntry);

                        int res = sboPrClose.Close();
                        int resultClose = sboPrClose.Update();

                        Application.SBO_Application.StatusBar.SetSystemMessage("Генерация успешно завершена",
                            SAPbouiCOM.BoMessageTime.bmt_Short, Type: SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                }
                catch (Exception exception)
                {
                    Application.SBO_Application.MessageBox($"{exception.Message}", 1, "Ok");
                }
            }
        }

        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }
    }
}


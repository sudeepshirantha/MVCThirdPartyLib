using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ThirdParty.lib
{
    public class Utility
    {
        public static void UpdateExPrecentageColor(ExcelWorksheet ws, string colName, int skip, int end){
            //Color
            for (int i = skip; i < (skip + end); i++)
            {
                object val = ws.Cells[colName + i + ":" + colName + i].Value;
                if (val!=null)
                {
                    try
                    {
                        decimal dVal = Decimal.Parse(val.ToString());
                        using (ExcelRange rng = ws.Cells[colName + i + ":" + colName + i])
                        {
                            dVal = dVal * 100;
                            if (dVal >=80 && dVal<=100)
                            {
                                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                rng.Style.Fill.BackgroundColor.SetColor(Color.Orange);
                            }
                            if (dVal > 100)
                            {
                                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                rng.Style.Fill.BackgroundColor.SetColor(Color.Red);
                            }
                        }

                    }
                    catch { }
                }
            }
        }


        //If Limit is 0 and Target >0 then val is RED
        public static void UpdateExPrecentageColorWithCondition(ExcelWorksheet ws, string colNameLimit, string colNameTarget, string colNameVal, int skip, int end)
        {
            //Color
            for (int i = skip; i < (skip + end); i++)
            {
                object val = ws.Cells[colNameVal + i + ":" + colNameVal + i].Value;
                object limit = ws.Cells[colNameLimit + i + ":" + colNameLimit + i].Value;
                object target = ws.Cells[colNameTarget + i + ":" + colNameTarget + i].Value;
                if (val != null && limit!=null && target !=null)
                {
                    try
                    {
                        decimal dVal = Decimal.Parse(val.ToString());
                        decimal dLimit = Decimal.Parse(limit.ToString());
                        decimal dTarget = Decimal.Parse(target.ToString());
                        using (ExcelRange rng = ws.Cells[colNameVal + i + ":" + colNameVal + i])
                        {
                            if (dLimit <= 0 && dTarget>0)
                            {
                                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                rng.Style.Fill.BackgroundColor.SetColor(Color.Red);
                            }
                        }

                    }
                    catch { }
                }
            }
        }


        public static DataTable ListToDataTable<T>(IList<T> data)
        {
            DataTable table = new DataTable();

            //special handling for value types and string
            if (typeof(T).IsValueType || typeof(T).Equals(typeof(string)))
            {

                DataColumn dc = new DataColumn("Value");
                table.Columns.Add(dc);
                foreach (T item in data)
                {
                    DataRow dr = table.NewRow();
                    dr[0] = item;
                    table.Rows.Add(dr);
                }
            }
            else
            {
                PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
                foreach (PropertyDescriptor prop in properties)
                {
                    table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
                }

                foreach (T item in data)
                {
                    DataRow row = table.NewRow();
                    foreach (PropertyDescriptor prop in properties)
                    {
                        try
                        {
                            row[prop.DisplayName] = prop.GetValue(item) ?? DBNull.Value;
                        }
                        catch (Exception ex)
                        {
                            row[prop.Name] = DBNull.Value;
                        }
                    }
                    table.Rows.Add(row);
                }
            }
            return table;
        }


        public static SortedSet<Currency> GetCurrenciesList()
        {
            SortedSet<Currency> list = new SortedSet<Currency>(new CurrencyComparer());
            foreach (CultureInfo cultureInfo in CultureInfo.GetCultures(CultureTypes.SpecificCultures))
            {
                RegionInfo regionInfo = new RegionInfo(cultureInfo.LCID);
                bool found = false;
                foreach (Currency item in list)
                {
                    if (item.CurrenctSymbol == regionInfo.ISOCurrencySymbol)
                    {
                        found = true;
                    }
                }
                if (!found)
                {
                    Currency li = new Currency();
                    li.CurrencyName = regionInfo.CurrencyEnglishName;
                    li.CurrenctSymbol = regionInfo.ISOCurrencySymbol;
                    list.Add(li);
                }
            }
            return list; 
        }
    }


    public class Currency : IComparer<Currency>
    {
        public string CurrencyName { set; get; }
        public string CurrenctSymbol { set; get; }

        public int Compare(Currency x, Currency y)
        {
            // TODO: Handle x or y being null, or them not having names
            return x.CurrencyName.CompareTo(y.CurrencyName);
        }
    }

    public class CurrencyComparer : IComparer<Currency>
    {
        public int Compare(Currency x, Currency y)
        {
            // TODO: Handle x or y being null, or them not having names
            return x.CurrencyName.CompareTo(y.CurrencyName);
        }
    }


    public class DecimalModelBinder : IModelBinder
    {
        public object BindModel(ControllerContext controllerContext,
            ModelBindingContext bindingContext)
        {
            ValueProviderResult valueResult = bindingContext.ValueProvider
                .GetValue(bindingContext.ModelName);
            ModelState modelState = new ModelState { Value = valueResult };
            object actualValue = null;
            try
            {
                actualValue = Convert.ToDecimal(valueResult.AttemptedValue,
                    CultureInfo.CurrentCulture);
            }
            catch (FormatException e)
            {
                modelState.Errors.Add(e);
            }

            bindingContext.ModelState.Add(bindingContext.ModelName, modelState);
            return actualValue;
        }
    }

}
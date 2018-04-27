using CucumberConverter.Common.model;
using System;
using System.Collections.Generic;

namespace CucumberConverter.Common
{
    public class ExcelConverter
    {
        #region Mapping method
        public void MapTestcase(TCExcelDto mapFrom, List<ExcelModel> mapTo)
        {
            foreach (var excelDto in mapFrom.TCList)
            {
                var tcModel = new ExcelModel();
                tcModel.TestcaseNumber = Convert.ToString(excelDto[2]);
                tcModel.HotelID = Convert.ToString(excelDto[3]);
                tcModel.CheckIn = Convert.ToString(excelDto[6]);
                tcModel.Los = Convert.ToString(excelDto[7]);
                tcModel.RatePlan = Convert.ToString(excelDto[8]);
                tcModel.Adult = Convert.ToString(excelDto[9]);
                tcModel.Children = Convert.ToString(excelDto[10]);
                tcModel.ChildAge = Convert.ToString(excelDto[11]);
                tcModel.Rooms = Convert.ToString(excelDto[12]);
                tcModel.IsAllOcc = Convert.ToString(excelDto[13]);
                tcModel.AllowOverideOcc = Convert.ToString(excelDto[14]);
                tcModel.HotelID2 = Convert.ToString(excelDto[15]);
                tcModel.RoomID = Convert.ToString(excelDto[16]);
                tcModel.Channel = Convert.ToString(excelDto[17]);
                tcModel.Currency = Convert.ToString(excelDto[18]);
                tcModel.RatePlan2 = Convert.ToString(excelDto[19]);
                tcModel.Occupancy = Convert.ToString(excelDto[20]);
                tcModel.MaxExtraBed = Convert.ToString(excelDto[22]);
                tcModel.IsFit = Convert.ToString(excelDto[23]);
                tcModel.ExtraBad = Convert.ToString(excelDto[25]);
                tcModel.SellEx = Convert.ToString(excelDto[26]);
                
                //breakdown task
                tcModel.Date = Convert.ToString(excelDto[27]);
                tcModel.Type = Convert.ToString(excelDto[28]);
                tcModel.Option = Convert.ToString(excelDto[29]);
                tcModel.Quantity = Convert.ToString(excelDto[30]);
                tcModel.SellEx2 = Convert.ToString(excelDto[31]);
                tcModel.Type2 = Convert.ToString(excelDto[33]);
                tcModel.Option2 = Convert.ToString(excelDto[34]);
                tcModel.Quantity2 = Convert.ToString(excelDto[35]);
                tcModel.SellEx3 = Convert.ToString(excelDto[36]);

                mapTo.Add(tcModel);
            }
        }
        #endregion
    }
}

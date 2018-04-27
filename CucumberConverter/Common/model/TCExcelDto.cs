using System;
using System.Collections.Generic;

namespace CucumberConverter.Common.model
{
    public class TCExcelDto
    {
        #region Propertyies
        public List<Dictionary<int, Object>> TCList { get; set; }
        #endregion

        public TCExcelDto()
        {
            TCList = new List<Dictionary<int, Object>>();
        }
    }
}

using CucumberConverter.Common.model;

namespace CucumberConverter.Common
{
    public enum ExcelHelperRule
    {
        CUCUMBER = 1
    }

    public class ExcelHelper
    {
        private const int TESTCASE_COL = 2;
        private const int CHECKIN_COL = 6;
        private const int HOTELID_COL = 15;
        private const int ROOMID_COL = 16;
        private const int CHANNEL_COL = 17;
        private const int CURRENCY_COL = 18;
        private const int RATEPLAN_COL = 19;

        public void HandleEmptyAndNullData(TCExcelDto arg, ExcelHelperRule rule)
        {
            switch(rule)
            {
                case ExcelHelperRule.CUCUMBER:
                    HandleByCucumberRule(arg);
                    break;
                default: return;
            }
        }

        public void HandleByCucumberRule(TCExcelDto arg)
        {
            object previousTestcaseId = null;
            object previousCheckinId = null;
            object previousPropertyId = null;
            object previousRoomId = null;
            object previousChannelId = null;
            object previousCurrencyId = null;
            object previousRatePlanId = null;
                
            foreach (var data in arg.TCList)
            {
                if (data[TESTCASE_COL] == null) { data[TESTCASE_COL] = previousTestcaseId; }
                if(data[CHECKIN_COL] == null) { data[CHECKIN_COL] = previousCheckinId; }
                if (data[HOTELID_COL] == null) { data[HOTELID_COL] = previousPropertyId; }
                if (data[ROOMID_COL] == null) { data[ROOMID_COL] = previousRoomId; }
                if (data[CHANNEL_COL] == null) { data[CHANNEL_COL] = previousChannelId; }
                if (data[CURRENCY_COL] == null) { data[CURRENCY_COL] = previousCurrencyId; }
                if (data[RATEPLAN_COL] == null) { data[RATEPLAN_COL] = previousRatePlanId; }

                previousTestcaseId = data[TESTCASE_COL];
                previousCheckinId = data[CHECKIN_COL];
                previousPropertyId = data[HOTELID_COL];
                previousRoomId = data[ROOMID_COL];
                previousChannelId = data[CHANNEL_COL];
                previousCurrencyId = data[CURRENCY_COL];
                previousRatePlanId = data[RATEPLAN_COL];
            }
        }
    }
}

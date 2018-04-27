using CucumberConverter.Common;
using CucumberConverter.Common.model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace CucumberConverter
{
    public partial class Window : Form
    {
        private readonly ExcelEPPluse _excelManager; // For read&write IO to excel
        private readonly ExcelHelper _excelHelper; // For read&write IO to excel
        private readonly ExcelConverter _excelConverter; // For read&write IO to excel
        System.IO.FileInfo fInfo;
        string scenario;
        string _contentHeader;
        string _contentBody;
        string fullmainfile;
        //breakdown
        string scenarioBreakdown;
        string _contentHeaderBreakdown;
        string _contentBodyBreakdown;
        string fullbreakdownfile;

        //decalr middle temp
        string day;
        string month;
        string year;
        string checkin;
        string day2;
        string month2;
        string year2;
        string checkin2;
        string occ;
        string adult;
        string child;
        string occ2;
        string adult2;
        string child2;
        string isfit;
        string option;
        string option2;
        string type;
        string type2;
        string sellex3;

        List<string> _testcase = new List<string>();
        List<string> _hotelid = new List<string>();
        List<string> _checkin = new List<string>();
        List<string> _los = new List<string>();
        List<string> _rateplan = new List<string>();
        List<string> _adult = new List<string>();
        List<string> _children = new List<string>();
        List<string> _childage = new List<string>();
        List<string> _room = new List<string>();
        List<string> _isallocc = new List<string>();
        List<string> _allownocc = new List<string>();
        List<string> _hotelid2 = new List<string>();
        List<string> _roomid = new List<string>();
        List<string> _channel = new List<string>();
        List<string> _currency = new List<string>();
        List<string> _rateplan2 = new List<string>();
        List<string> _occupancy = new List<string>();
        List<string> _maxextrabed = new List<string>();
        List<string> _isfit = new List<string>();
        List<string> _extrabed = new List<string>();
        List<string> _sellex = new List<string>();

        //breakdown
        List<string> _date = new List<string>();
        List<string> _type = new List<string>();
        List<string> _option = new List<string>();
        List<string> _quantity = new List<string>();
        List<string> _sellex2 = new List<string>();
        List<string> _type2 = new List<string>();
        List<string> _option2 = new List<string>();
        List<string> _quantity2 = new List<string>();
        List<string> _sellex3 = new List<string>();

        public Window()
        {
            InitializeComponent();
            _excelManager = new ExcelEPPluse();
            _excelHelper = new ExcelHelper();
            _excelConverter = new ExcelConverter();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Browse Text Files";

            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;

            openFileDialog1.DefaultExt = "feature";
            openFileDialog1.Filter = "Text files (*.feature)|*.feature|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            openFileDialog1.ReadOnlyChecked = true;
            openFileDialog1.ShowReadOnly = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fInfo = new System.IO.FileInfo(openFileDialog1.FileName);
                textBox1.Text = openFileDialog1.FileName;
                ScreenData(openFileDialog1.FileName);
            }
        }

        public void ScreenData(string filename)
        {
            string ExcelLocation = filename;
            int ExcelPageNumber = 1;

            var tcList = new List<ExcelModel>();  // Data which we really want to use in this test
            var excelData = _excelManager.Read(ExcelLocation, ExcelPageNumber);
            _excelHelper.HandleEmptyAndNullData(excelData, ExcelHelperRule.CUCUMBER);
            _excelConverter.MapTestcase(excelData, tcList);

            for (var i = 0; i < tcList.Count; i++)
            {
                _testcase.Add(tcList[i].TestcaseNumber);
                _hotelid.Add(tcList[i].HotelID);
                _checkin.Add(tcList[i].CheckIn);
                _los.Add(tcList[i].Los);
                _rateplan.Add(tcList[i].RatePlan);
                _adult.Add(tcList[i].Adult);
                _children.Add(tcList[i].Children);
                _childage.Add(tcList[i].ChildAge);
                _room.Add(tcList[i].Rooms);
                _isallocc.Add(tcList[i].IsAllOcc);
                _allownocc.Add(tcList[i].AllowOverideOcc);
                _hotelid2.Add(tcList[i].HotelID2);
                _roomid.Add(tcList[i].RoomID);
                _channel.Add(tcList[i].Channel);
                _currency.Add(tcList[i].Currency);
                _rateplan2.Add(tcList[i].RatePlan2);
                _occupancy.Add(tcList[i].Occupancy);
                _maxextrabed.Add(tcList[i].MaxExtraBed);
                _isfit.Add(tcList[i].IsFit);
                _extrabed.Add(tcList[i].ExtraBad);
                _sellex.Add(tcList[i].SellEx);

                //breakdown task
                _date.Add(tcList[i].Date);
                _type.Add(tcList[i].Type);
                _option.Add(tcList[i].Option);
                _quantity.Add(tcList[i].Quantity);
                _sellex2.Add(tcList[i].SellEx2);
                _type2.Add(tcList[i].Type2);
                _option2.Add(tcList[i].Option2);
                _quantity2.Add(tcList[i].Quantity2);
                _sellex3.Add(tcList[i].SellEx3);
            }
        }
    
        public void ConvertToCucumber()
        {
            scenario = textBox2.Text;
            scenarioBreakdown = textBox5.Text;

            #region main file start
            for (var i = 0; i < _testcase.Count; i++)
            {
                if (_checkin[i] != "" && _occupancy[i] != "")
                {
                    string[] checkinDate = _checkin[i].Split('/');
                    //day = _checkin[i].Substring(0, 2).Trim();
                    //month = _checkin[i].Substring(3, 2).Trim();
                    //year = _checkin[i].Substring(6, 4).Trim();
                    //checkin = year + "-" + month + "-" + day;
                    checkin = checkinDate[0] + "-" + checkinDate[1] + "-" + checkinDate[2];

                    occ = _occupancy[i].Substring(0, 1);
                    adult = _occupancy[i].Substring(2, 1);
                    child = _occupancy[i].Substring(5, 1);

                    if (i == 0)
                    {
                        _contentHeader = "  @OccupancySearch" + System.Environment.NewLine + "  Scenario: " + scenario + " - TC" + _testcase[i] + System.Environment.NewLine + "    Given Hotels " + _hotelid[i] +
                                          " checkin " + checkin + " los " + _los[i] + System.Environment.NewLine + "    And RatePlans " + _rateplan[i] + System.Environment.NewLine + "    And Adults " + _adult[i] + System.Environment.NewLine +
                                          "    And Children " + _children[i] + System.Environment.NewLine + "    And ChildAges " + _childage[i] + System.Environment.NewLine + "    And Rooms " + _room[i] + System.Environment.NewLine + "    And SuggestedPrice " + _isallocc[i].ToLower() + System.Environment.NewLine +
                                          "    And OverrideOcc " + _allownocc[i].ToLower() + System.Environment.NewLine + "    And Currency " + _currency[i] + System.Environment.NewLine + "    When The user search" + System.Environment.NewLine + "    Then OccupancySearch in currency " + _currency[i] +
                                          " should match result" + System.Environment.NewLine + "      | hotelid | roomid   | channel | rateplan | occupancy | adults | children | isFit | extrabed | maxExtrabed | sellEx  |" + System.Environment.NewLine;
                    }
                    if (i != 0 && (_testcase[i] != _testcase[i - 1]))
                    {
                        _contentBody = _contentBody + System.Environment.NewLine + "  @OccupancySearch" + System.Environment.NewLine + "  Scenario: " + scenario + " - TC" + _testcase[i] + System.Environment.NewLine + "    Given Hotels " + _hotelid[i] +
                                          " checkin " + checkin + " los " + _los[i] + System.Environment.NewLine + "    And RatePlans " + _rateplan[i] + System.Environment.NewLine + "    And Adults " + _adult[i] + System.Environment.NewLine +
                                          "    And Children " + _children[i] + System.Environment.NewLine + "    And ChildAges " + _childage[i] + System.Environment.NewLine + "    And Rooms " + _room[i] + System.Environment.NewLine + "    And SuggestedPrice " + _isallocc[i].ToLower() + System.Environment.NewLine +
                                          "    And OverrideOcc " + _allownocc[i].ToLower() + System.Environment.NewLine + "    And Currency " + _currency[i] + System.Environment.NewLine + "    When The user search" + System.Environment.NewLine + "    Then OccupancySearch in currency " + _currency[i] +
                                          " should match result" + System.Environment.NewLine + "      | hotelid | roomid   | channel | rateplan | occupancy | adults | children | isFit | extrabed | maxExtrabed | sellEx  |" + System.Environment.NewLine;
                    }

                    if (_isfit[i].ToLower() == "false")
                    {
                        isfit = _isfit[i].ToLower() + " | ";
                    }
                    else
                    {
                        isfit = _isfit[i].ToLower() + "  | ";
                    }

                    _contentBody = _contentBody + "      | " + _hotelid2[i] + "  | " + _roomid[i] + " | " + _channel[i] + "       | " + _rateplan2[i] + "        | " + occ + "         | " + adult + "      | " + child +
                                          "        | " + isfit + _extrabed[i] + "        | " + _maxextrabed[i] + "           |  " + Convert.ToDecimal(_sellex[i]).ToString("0.00") + " |" + System.Environment.NewLine;
                }
                else
                {
                    _contentBody = _contentBody + System.Environment.NewLine + "  @OccupancySearch" + System.Environment.NewLine + "  Scenario: " + scenario + " - TC" + _testcase[i] + System.Environment.NewLine + "    Given Hotels " + _hotelid[i] +
                                          " checkin " + checkin + " los " + _los[i] + System.Environment.NewLine + "    And RatePlans " + _rateplan[i] + System.Environment.NewLine + "    And Adults " + _adult[i] + System.Environment.NewLine +
                                          "    And Children " + _children[i] + System.Environment.NewLine + "    And ChildAges " + _childage[i] + System.Environment.NewLine + "    And Rooms " + _room[i] + System.Environment.NewLine + "    And SuggestedPrice " + _isallocc[i].ToLower() + System.Environment.NewLine +
                                          "    And OverrideOcc " + _allownocc[i].ToLower() + System.Environment.NewLine + "    And Currency " + _currency[i] + System.Environment.NewLine + "    When The user search" + System.Environment.NewLine + "    Then OccupancySearch in currency " + _currency[i] +
                                          " should match result" + System.Environment.NewLine + "      | hotelid | roomid   | channel | rateplan | occupancy | adults | children | isFit | extrabed | maxExtrabed | sellEx  |" + System.Environment.NewLine;
                }
            }
            fullmainfile = _contentHeader + _contentBody;
            #endregion main file end

            #region breakdown file start
            var firstIndexInPreviousSet = 0;

            for (var i = 0; i < _testcase.Count; i++)
            {
                if (_checkin[i] != "" && _occupancy[i] != "")
                {
                    string[] checkinDate = _checkin[i].Split('/');
                    //day = _checkin[i].Substring(0, 2).Trim();
                    //month = _checkin[i].Substring(3, 2).Trim();
                    //year = _checkin[i].Substring(6, 4).Trim();
                    //checkin = year + "-" + month + "-" + day;
                    checkin = checkinDate[0] + "-" + checkinDate[1] + "-" + checkinDate[2];

                    occ = _occupancy[i].Substring(0, 1);
                    adult = _occupancy[i].Substring(2, 1);
                    child = _occupancy[i].Substring(5, 1);

                    if (i == 0)
                    {
                        _contentHeaderBreakdown = "  @OccupancySearchPriceBreakdowns" + System.Environment.NewLine + "  Scenario: " + scenarioBreakdown + " - TC" + _testcase[i] + System.Environment.NewLine + "    Given Hotels " + _hotelid[i] +
                                          " checkin " + checkin + " los " + _los[i] + System.Environment.NewLine + "    And RatePlans " + _rateplan[i] + System.Environment.NewLine + "    And Adults " + _adult[i] + System.Environment.NewLine +
                                          "    And Children " + _children[i] + System.Environment.NewLine + "    And ChildAges " + _childage[i] + System.Environment.NewLine + "    And Rooms " + _room[i] + System.Environment.NewLine + "    And SuggestedPrice " + _isallocc[i].ToLower() + System.Environment.NewLine +
                                          "    And OverrideOcc " + _allownocc[i].ToLower() + System.Environment.NewLine + "    And Currency " + _currency[i] + System.Environment.NewLine + "    When The user search" + System.Environment.NewLine + "    Then OccupancySearch(PriceBreakdown - Room) in currency " + _currency[i] +
                                          " should match result" + System.Environment.NewLine + "      | hotelid | roomid   | channel | rateplan | occupancy | adults | children | extrabed | date       | taxType  | quantity | sellEx  | isMandatory |" + System.Environment.NewLine;
                    }

                    //re-line
                    if (_option[i] == "Mandatory")
                    {
                        option = "true" + "        |";
                    }
                    else
                    {
                        option = "false" + "       |";
                    }

                    if (_type[i] == "Room")
                    {
                        type = "Room" + "     | ";
                    }
                    else
                    {
                        type = "Extrabed" + " | ";
                    }
                    //

                    if (i != 0 && ((i == _testcase.Count - 1) || (_roomid[i] != _roomid[i - 1])))
                    {
                        var lastIndexToPrint = (i == _testcase.Count - 1) ? i : i - 1;

                        //breakdown2
                        for (var j = firstIndexInPreviousSet; j <= lastIndexToPrint; j++)
                        {
                            if (_checkin[j] != "" && _occupancy[j] != "")
                            {
                                string[] checkinDate2 = _checkin[j].Split('/');
                                checkin2 = checkinDate[0] + "-" + checkinDate[1] + "-" + checkinDate[2];
                                //day2 = _checkin[j].Substring(0, 2).Trim();
                                //month2 = _checkin[j].Substring(3, 2).Trim();
                                //year2 = _checkin[j].Substring(6, 4).Trim();
                                //checkin2 = year + "-" + month + "-" + day;

                                occ2 = _occupancy[j].Substring(0, 1);
                                adult2 = _occupancy[j].Substring(2, 1);
                                child2 = _occupancy[j].Substring(5, 1);

                                //re-line
                                if (_option2[j] == "Mandatory")
                                {
                                    option2 = "true" + "        |";
                                }
                                else
                                {
                                    option2 = "false" + "       |";
                                }

                                if (_type2[j] == "Room")
                                {
                                    type2 = "Room" + "     | ";
                                }
                                else
                                {
                                    type2 = "Extrabed" + " | ";
                                }
                                //

                                if (_type2[j] != "")
                                {
                                    sellex3 = Convert.ToDecimal(_sellex3[j]).ToString("0.00");
                                    if (sellex3.Length < 6)
                                    {
                                        sellex3 = " " + sellex3;
                                    }
                                    _contentBodyBreakdown = _contentBodyBreakdown + "      | " + _hotelid2[j] + "  | " + _roomid[j] + " | " + _channel[j] + "       | " + _rateplan2[j] + "        | " + occ2 + "         | " + adult2 + "      | " + child2 +
                                                          "        | " + _extrabed[j] + "        | " + checkin2 + " | " + type2 + _quantity2[j] + "        | " + sellex3 + "  | " + option2 + System.Environment.NewLine;
                                }
                            }

                            firstIndexInPreviousSet = i;
                        }
                    }
                    if (_checkin[i] != "" && _occupancy[i] != "")
                    {
                        if (i != 0 && (_testcase[i] != _testcase[i - 1]))
                        {
                            _contentBodyBreakdown = _contentBodyBreakdown + System.Environment.NewLine + "  @OccupancySearchPriceBreakdowns" + System.Environment.NewLine + "  Scenario: " + scenarioBreakdown + " - TC" + _testcase[i] + System.Environment.NewLine + "    Given Hotels " + _hotelid[i] +
                                              " checkin " + checkin + " los " + _los[i] + System.Environment.NewLine + "    And RatePlans " + _rateplan[i] + System.Environment.NewLine + "    And Adults " + _adult[i] + System.Environment.NewLine +
                                              "    And Children " + _children[i] + System.Environment.NewLine + "    And ChildAges " + _childage[i] + System.Environment.NewLine + "    And Rooms " + _room[i] + System.Environment.NewLine + "    And SuggestedPrice " + _isallocc[i].ToLower() + System.Environment.NewLine +
                                              "    And OverrideOcc " + _allownocc[i].ToLower() + System.Environment.NewLine + "    And Currency " + _currency[i] + System.Environment.NewLine + "    When The user search" + System.Environment.NewLine + "    Then OccupancySearch(PriceBreakdown - Room) in currency " + _currency[i] +
                                              " should match result" + System.Environment.NewLine + "      | hotelid | roomid   | channel | rateplan | occupancy | adults | children | extrabed | date       | taxType  | quantity | sellEx  | isMandatory |" + System.Environment.NewLine;
                        }

                        if (!(i == _testcase.Count - 1))
                        {
                            _contentBodyBreakdown = _contentBodyBreakdown + "      | " + _hotelid2[i] + "  | " + _roomid[i] + " | " + _channel[i] + "       | " + _rateplan2[i] + "        | " + occ + "         | " + adult + "      | " + child +
                                                           "        | " + _extrabed[i] + "        | " + checkin + " | " + type + _quantity[i] + "        | " + Convert.ToDecimal(_sellex2[i]).ToString("0.00") + "  | " + option + System.Environment.NewLine;
                        }
                    }
                }
                else
                {
                    _contentBodyBreakdown = _contentBodyBreakdown + System.Environment.NewLine + "  @OccupancySearchPriceBreakdowns" + System.Environment.NewLine + "  Scenario: " + scenarioBreakdown + " - TC" + _testcase[i] + System.Environment.NewLine + "    Given Hotels " + _hotelid[i] +
                                             " checkin " + checkin + " los " + _los[i] + System.Environment.NewLine + "    And RatePlans " + _rateplan[i] + System.Environment.NewLine + "    And Adults " + _adult[i] + System.Environment.NewLine +
                                             "    And Children " + _children[i] + System.Environment.NewLine + "    And ChildAges " + _childage[i] + System.Environment.NewLine + "    And Rooms " + _room[i] + System.Environment.NewLine + "    And SuggestedPrice " + _isallocc[i].ToLower() + System.Environment.NewLine +
                                             "    And OverrideOcc " + _allownocc[i].ToLower() + System.Environment.NewLine + "    And Currency " + _currency[i] + System.Environment.NewLine + "    When The user search" + System.Environment.NewLine + "    Then OccupancySearch(PriceBreakdown - Room) in currency " + _currency[i] +
                                             " should match result" + System.Environment.NewLine + "      | hotelid | roomid   | channel | rateplan | occupancy | adults | children | extrabed | date       | taxType  | quantity | sellEx  | isMandatory |" + System.Environment.NewLine;
                }
            }
            fullbreakdownfile = _contentHeaderBreakdown + _contentBodyBreakdown;
            #endregion breakdown file end

            WriteFile(fullmainfile, fullbreakdownfile);
        }

        public void WriteFile(string mainbody, string breakdownbody)
        {
            scenario = textBox2.Text;
            scenarioBreakdown = textBox5.Text;

            string strFilePathMain = fInfo.DirectoryName+"\\" + scenario + ".feature";
            string strFilePathbreakdown = fInfo.DirectoryName + "\\" + scenarioBreakdown + ".feature";

            if (File.Exists(strFilePathMain))
            {
                File.Delete(strFilePathMain);
            }
            if (File.Exists(strFilePathbreakdown))
            {
                File.Delete(strFilePathbreakdown);
            }

            { // Consider File Operation 1
                FileStream fs1 = new FileStream(strFilePathMain, FileMode.OpenOrCreate);
                StreamWriter str1 = new StreamWriter(fs1);
                str1.BaseStream.Seek(0, SeekOrigin.End);
                str1.Write(mainbody);
                str1.Flush();
                str1.Close();
                fs1.Close();
                // Close the Stream then Individually you can access the file.
            }
            { // Consider File Operation 2
                FileStream fs2 = new FileStream(strFilePathbreakdown, FileMode.OpenOrCreate);
                StreamWriter str2 = new StreamWriter(fs2);
                str2.BaseStream.Seek(0, SeekOrigin.End);
                str2.Write(breakdownbody);
                str2.Flush();
                str2.Close();
                fs2.Close();
                // Close the Stream then Individually you can access the file.
            }
            this.textBox3.Text = strFilePathMain;
            this.textBox4.Text = strFilePathbreakdown;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ConvertToCucumber();
        }

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using DocumentFormat.OpenXml;
using SpreadsheetLight;
using System.IO;
using System.Globalization;

namespace Holistic
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
            CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.InvariantCulture;
        }
        public class Person
        {
            public int ID { get; set; }
            public string Name { get; set; }
            public string ProName { get; set; }
            public int OtherCount { get; set; }
            public int OpenCount { get; set; }
            public int RecordCount { get; set; }
            /// <summary>
            /// 開案記錄
            /// 1 : 開案
            /// 2 : 主記錄
            /// </summary>
            public int OpenCase { get; set; }
            public bool YesRecord
            {
                get
                {
                    return RecordID.Count > 0;
                }
            }
            public List<string> RecordID { get; set; }
            public Person()
            {
                RecordID = new List<string>();
            }
        }
        public Dictionary<string, int> StationDatas = new Dictionary<string, int>();
        public List<Person> PersonDatas = new List<Person>();
        public int PI_Count;
        public int TotalCount;
        private void Btn_Cal_Click(object sender, RoutedEventArgs e)
        {
            string fpath = Environment.CurrentDirectory + @"\Data";
            if (!Directory.Exists(fpath))
            {
                return;
            }
            string fname = fpath + @"\全人.xlsx";
            if (!System.IO.File.Exists(fname))
                return;
            using (SLDocument sl = new SLDocument(fname))
            {
                SLWorksheetStatistics wsstats = sl.GetWorksheetStatistics();
                int slrows = wsstats.EndRowIndex;
                for (int i = 0; i < slrows + 10; i++)
                {
                    if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 3, 2)))
                        break;
                    TotalCount++;
                    string pid = sl.GetCellValueAsString(i + 3, 4);
                    string station = sl.GetCellValueAsString(i + 3, 11);
                    ///護理站
                    if (StationDatas.ContainsKey(station))
                    {
                        StationDatas[station]++;
                    }
                    else
                    {
                        StationDatas.Add(station, 1);
                    }
                    if (station == "PI")
                        PI_Count++;

                    ///開案
                    PersonDatas.Add(new Person()
                    {
                        ID = sl.GetCellValueAsInt32(i + 3, 16),
                        Name = sl.GetCellValueAsString(i + 3, 17),
                        ProName = "護理師",
                        OpenCase = 1
                    });
                    ///醫護
                    PersonDatas.Add(new Person()
                    {
                        ID = sl.GetCellValueAsInt32(i + 3, 21),
                        Name = sl.GetCellValueAsString(i + 3, 22),
                        ProName = sl.GetCellValueAsString(i + 3, 24),
                        OpenCase = 2
                    });
                    for (int j = 0; j < 13; j++)
                    {
                        if (!string.IsNullOrEmpty(sl.GetCellValueAsString(i + 3, 25 + (j * 3))))
                        {
                            PersonDatas.Add(new Person()
                            {
                                ID = sl.GetCellValueAsInt32(i + 3, 25 + (j * 3)),
                                Name = sl.GetCellValueAsString(i + 3, 26 + (j * 3)),
                                ProName = sl.GetCellValueAsString(2, 26 + (j * 3)),
                                RecordID = sl.GetCellValueAsString(i + 3, 27 + (j * 3)) == "Y" ?
                                new List<string>() { pid } : new List<string>()
                            });
                        }
                    }
                }
                sl.CloseWithoutSaving();
            }
            try
            {
                fname = fpath + @"\全人(1).xlsx";
                using (SLDocument sl = new SLDocument())
                {
                    sl.RenameWorksheet("Sheet1", "病房獎勵金表");
                    for (int z = 0; z < 16; z++)
                        sl.SetColumnWidth(z + 1, 15);
                    sl.SetCellValue(1, 1, "病房獎勵金");
                    sl.SetCellValue(2, 1, "病房獎勵金");
                    sl.SetCellValue(2, 2, "件數");
                    sl.SetCellValue(2, 3, "獎勵金");
                    sl.SetCellValue(2, 4, "NP公款件數");
                    sl.SetCellValue(2, 5, "NP公款獎勵金");
                    sl.SetCellValue(2, 7, "代領人名稱");
                    sl.SetCellValue(2, 8, "代領人代號");
                    var sort = (from obj in StationDatas orderby obj.Key ascending select obj).ToDictionary(o => o.Key, o => o.Value);
                    int i = 0;
                    foreach (var x in sort)
                    {
                        sl.SetCellValue(3 + i, 1, x.Key);
                        sl.SetCellValue(3 + i, 2, x.Value);
                        sl.SetCellValue(3 + i, 3, Convert.ToInt32(x.Value) * 200);
                        if (x.Key == "PI")
                        {
                            sl.SetCellValue(3 + i, 4, PI_Count);
                            sl.SetCellValue(3 + i, 5, PI_Count * 100);
                        }
                        i++;
                    }
                    sl.AddWorksheet("記錄獎勵金表");
                    sl.SelectWorksheet("記錄獎勵金表");
                    for (int z = 0; z < 16; z++)
                        sl.SetColumnWidth(z + 1, 15);
                    sl.SetCellValue(1, 1, "職類獎勵金");
                    sl.SetCellValue(2, 1, "員工代號");
                    sl.SetCellValue(2, 2, "員工名稱");
                    sl.SetCellValue(2, 3, "職稱");
                    sl.SetCellValue(2, 4, "開案件數");
                    sl.SetCellValue(2, 5, "金額 (50)");
                    sl.SetCellValue(2, 6, "主記錄件數");
                    sl.SetCellValue(2, 7, "金額 (150)");
                    sl.SetCellValue(2, 8, "職類記錄件數");
                    sl.SetCellValue(2, 9, "金額 (50)");
                    sl.SetCellValue(2, 10, "總金額");
                    List<Person> personcount = new List<Person>();
                    foreach (var x in PersonDatas)
                    {
                        var data = personcount.FirstOrDefault(o => o.ID == x.ID);
                        if (data != null)
                        {
                            if (x.OpenCase == 1)
                                data.OpenCount++;
                            else if (x.OpenCase == 2)
                                data.RecordCount++;
                            else if (x.YesRecord)
                                data.OtherCount++;
                        }
                        else
                        {
                            Person addon = new Person() { ID = x.ID, Name = x.Name, ProName = x.ProName };
                            if (x.OpenCase == 1)
                                addon.OpenCount++;
                            else if (x.OpenCase == 2)
                                addon.RecordCount++;
                            else if (x.YesRecord)
                                addon.OtherCount++;
                            personcount.Add(addon);
                        }
                    }
                    //personcount.Sort((x, y) => { return x.ID.CompareTo(y.ID); });
                    //personcount.Sort((x, y) => { return x.ProName.CompareTo(y.ProName); });
                    var pdata = personcount.GroupBy(o => o.ProName).ToDictionary(o => o.Key, o => o.ToList<Person>());
                    i = 0;
                    int j = 8;
                    foreach (var o in pdata)
                    {
                        o.Value.Sort((x, y) => { return x.ID.CompareTo(y.ID); });
                        foreach (var x in o.Value)
                        {
                            sl.SetCellValue(3 + i, 1, x.ID);
                            sl.SetCellValue(3 + i, 2, x.Name);
                            sl.SetCellValue(3 + i, 3, x.ProName);
                            if (x.OpenCount > 0)
                            {
                                sl.SetCellValue(3 + i, 4, x.OpenCount);
                                sl.SetCellValue(3 + i, 5, x.OpenCount * 50);
                            }
                            else if (x.RecordCount > 0)
                            {
                                sl.SetCellValue(3 + i, 6, x.RecordCount);
                                sl.SetCellValue(3 + i, 7, x.RecordCount * 150);
                            }
                            else if (x.OtherCount > 0)
                            {
                                sl.SetCellValue(3 + i, 8, x.OtherCount);
                                sl.SetCellValue(3 + i, 9, x.OtherCount * 50);
                            }
                            sl.SetCellValue(3 + i, 10, x.OpenCount * 50 + x.RecordCount * 150 + x.OtherCount * 50);
                            i++;

                            if (o.Key.Contains("個管") && x.OtherCount + x.OpenCount + x.RecordCount > 0)
                            {
                                sl.SetCellValue(j, 12, x.ID);
                                sl.SetCellValue(j, 13, x.Name);
                                sl.SetCellValue(j, 14, x.ProName);
                                sl.SetCellValue(j, 15, x.OtherCount + x.OpenCount + x.RecordCount);
                                j++;
                            }
                        }
                    }
                    sl.SetCellValue(2, 12, "員工代號");
                    sl.SetCellValue(2, 13, "員工名稱");
                    sl.SetCellValue(2, 14, "職稱");
                    sl.SetCellValue(2, 15, "件數(NP)");
                    sl.SetCellValue(2, 16, "總金額 (100)");
                    sl.SetCellValue(3, 15, TotalCount - PI_Count);
                    sl.SetCellValue(3, 16, 100 * (TotalCount - PI_Count));

                    sl.SetCellValue(5, 12, "員工代號");
                    sl.SetCellValue(5, 13, "員工名稱");
                    sl.SetCellValue(5, 14, "職稱");
                    sl.SetCellValue(5, 15, "件數(護理部公款)");
                    sl.SetCellValue(5, 16, "總金額 (100)");
                    sl.SetCellValue(6, 15, TotalCount);
                    sl.SetCellValue(6, 16, 100 * TotalCount);
                    /*
                    Dictionary<string, List<Person>> pdatas = new Dictionary<string, List<Person>>();

                    foreach (var x in PersonDatas)
                    {
                        if (pdatas.ContainsKey(x.ProName))
                        {
                            foreach (var y in pdatas)
                            {
                                if (y.Key != x.ProName)
                                    continue;
                                var data = y.Value.FirstOrDefault(o => o.Name == x.Name);
                                if (data != null)
                                {
                                    y.Value.FirstOrDefault(o => o.Name == x.Name).OtherCount++;
                                }
                                else
                                {
                                    y.Value.Add(new Person() { Name = x.Name, ID = x.ID, OtherCount = 1 });
                                }
                            }
                        }
                        else
                        {
                            pdatas.Add(x.ProName, new List<Person>() { new Person() { Name = x.Name, ID = x.ID, OtherCount = 1 } });
                        }
                    }
                    i = 0;
                    foreach (var o in pdatas)
                    {
                        o.Value.Sort((x, y) => { return x.ID.CompareTo(y.ID); });
                        foreach (var x in o.Value)
                        {
                            sl.SetCellValue(3 + i, 1, x.ID);
                            sl.SetCellValue(3 + i, 2, x.Name);
                            sl.SetCellValue(3 + i, 3, o.Key);
                            if (x.OpenCase == 1)
                            {
                                sl.SetCellValue(3 + i, 4, x.OtherCount);
                                sl.SetCellValue(3 + i, 5, x.OtherCount * 50);
                            }
                            else if (x.OpenCase == 2)
                            {
                                sl.SetCellValue(3 + i, 6, x.OtherCount);
                                sl.SetCellValue(3 + i, 7, x.OtherCount * 150);
                            }
                            else
                            {
                                sl.SetCellValue(3 + i, 8, x.OtherCount);
                                sl.SetCellValue(3 + i, 9, x.OtherCount * 50);
                            }
                            i++;
                        }
                    }
                    */
                    sl.SaveAs(fname);
                    MessageBox.Show("Done");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}

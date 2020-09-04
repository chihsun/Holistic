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
            System.Threading.Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            System.Threading.Thread.CurrentThread.CurrentUICulture = CultureInfo.InvariantCulture;
        }
        public class Person
        {
            public int ID { get; set; }
            public string Name { get; set; }
            public string ProName { get; set; }
            public int OtherCount { get; set; }
            public int OpenCount { get; set; }
            public int MultiOpen { get; set; }
            public int RecordCount { get; set; }
            public int MultiRecord { get; set; }
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
            public int Multi { get; set; }
        }
        public Dictionary<string, int> StationDatas = new Dictionary<string, int>();
        public List<Person> PersonDatas = new List<Person>();
        /// <summary>
        /// PI 護理長津貼
        /// </summary>
        public int PI_Count;
        /// <summary>
        /// 案例總數
        /// </summary>
        public int TotalCount;
        /// <summary>
        /// 各案例職類數
        /// </summary>
        public List<int> MultiCount = new List<int>();
        /// <summary>
        /// 個案ID及計數(重覆)
        /// </summary>
        public Dictionary<int, int> PID = new Dictionary<int, int>();
        private void Btn_Cal_Click(object sender, RoutedEventArgs e)
        {
            string fname;
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog
            {
                InitialDirectory = Environment.CurrentDirectory,
                Title = "選取資料檔",
                Filter = "xlsx files (*.*)|*.xlsx"
            };
            if (dlg.ShowDialog() == true)
            {
                fname = dlg.FileName;
            }
            else
                return;

            if (!System.IO.File.Exists(fname))
                return;
            try
            {
                using (SLDocument sl = new SLDocument(fname))
                {
                    SLWorksheetStatistics wsstats = sl.GetWorksheetStatistics();
                    for (int i = 0; i < wsstats.EndRowIndex; i++)
                    {
                        if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, 1)))
                            break;
                        TotalCount++;
                        string pid = sl.GetCellValueAsString(i + 2, 3);
                        string station = sl.GetCellValueAsString(i + 2, 10);
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
                        ///其他職類
                        int multp = 1;
                        for (int j = 0; j < 15; j++)
                        {
                            if (!string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, 24 + (j * 3)))
                                && !sl.GetCellValueAsString(i + 2, 24 + (j * 3)).Contains("會議記錄完成")
                                && Int32.TryParse(sl.GetCellValueAsString(i + 2, 24 + (j * 3)), out int id))
                            {
                                PersonDatas.Add(new Person()
                                {
                                    ID = sl.GetCellValueAsInt32(i + 2, 24 + (j * 3)),
                                    Name = sl.GetCellValueAsString(i + 2, 25 + (j * 3)),
                                    ProName = sl.GetCellValueAsString(1, 25 + (j * 3)),
                                    RecordID = sl.GetCellValueAsString(i + 2, 26 + (j * 3)) == "Y" ?
                                    new List<string>() { pid } : new List<string>()
                                });
                                multp++;
                            }
                        }
                        ///開案
                        PersonDatas.Add(new Person()
                        {
                            ID = sl.GetCellValueAsInt32(i + 2, 15),
                            Name = sl.GetCellValueAsString(i + 2, 16),
                            ProName = "護理長",
                            OpenCase = 1,
                            Multi = multp
                        });
                        ///醫護
                        PersonDatas.Add(new Person()
                        {
                            ID = sl.GetCellValueAsInt32(i + 2, 20),
                            Name = sl.GetCellValueAsString(i + 2, 21),
                            ProName = sl.GetCellValueAsString(i + 2, 23),
                            OpenCase = 2,
                            Multi = multp
                        });
                        MultiCount.Add(multp);
                    }
                    sl.CloseWithoutSaving();
                }
            }
            catch (Exception ex )
            {
                MessageBox.Show(ex.ToString());
            }
            try
            {
                string fpath = Environment.CurrentDirectory + @"\獎勵金";
                if (!Directory.Exists(fpath))
                {
                    Directory.CreateDirectory(fpath);
                }
                fname = fpath + @"\全人" + DateTime.Now.ToString("yyyy-MM")+ ".xlsx";
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
                    sl.SetCellValue(2, 7, "代領人代號");
                    sl.SetCellValue(2, 8, "代領人名稱");
                    sl.SetCellValue(2, 9, "總金額");
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
                        sl.SetCellValue(3 + i, 9, Convert.ToInt32(x.Value) * 200 + PI_Count * 100);
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
                    sl.SetCellValue(2, 5, "金額 (50 or 70)");
                    sl.SetCellValue(2, 6, "主記錄件數");
                    sl.SetCellValue(2, 7, "金額 (150 or 180)");
                    sl.SetCellValue(2, 8, "職類記錄件數");
                    sl.SetCellValue(2, 9, "金額 (50)");
                    sl.SetCellValue(2, 10, "總金額");
                    sl.SetCellValue(2, 11, "備註");
                    List<Person> personcount = new List<Person>();
                    foreach (var x in PersonDatas)
                    {
                        var data = personcount.FirstOrDefault(o => o.ID == x.ID);
                        if (data != null)
                        {
                            if (x.OpenCase == 1)
                            {
                                data.OpenCount++;
                                if (x.Multi >= 3)
                                    data.MultiOpen++;
                            }
                            else if (x.OpenCase == 2)
                            {
                                data.RecordCount++;
                                if (x.Multi >= 3)
                                    data.MultiRecord++;
                            }
                            else if (x.YesRecord)
                                data.OtherCount++;
                        }
                        else
                        {
                            Person addon = new Person() { ID = x.ID, Name = x.Name, ProName = x.ProName };
                            if (x.OpenCase == 1)
                            {
                                addon.OpenCount++;
                                if (x.Multi >= 3)
                                    addon.MultiOpen++;
                            }
                            else if (x.OpenCase == 2)
                            {
                                addon.RecordCount++;
                                if (x.Multi >= 3)
                                    addon.MultiRecord++;
                            }
                            else if (x.YesRecord)
                                addon.OtherCount++;
                            personcount.Add(addon);
                        }
                    }
                    //personcount.Sort((x, y) => { return x.ID.CompareTo(y.ID); });
                    //personcount.Sort((x, y) => { return x.ProName.CompareTo(y.ProName); });
                    var ndata = personcount;
                    ndata.Sort((x, y) => { return x.ProName.CompareTo(y.ProName); });
                    var pdata = ndata.GroupBy(o => o.OtherCount > 0 || (o.OpenCount == 0 && o.RecordCount == 0 && o.OtherCount == 0)).ToDictionary(o => o.Key, o => o.ToList<Person>());
                    //var pdata = personcount.GroupBy(o => o.ProName).ToDictionary(o => o.Key, o => o.ToList<Person>());
                    i = 0;
                    int j = 8;
                    foreach (var o in pdata)
                    {
                        //o.Value.Sort((x, y) => { return x.ID.CompareTo(y.ID); });
                        var odata = o.Value.GroupBy(op => op.ProName).ToDictionary(op => op.Key, op => op.ToList<Person>());
                        foreach (var p in odata)
                        {
                            p.Value.Sort((x, y) => { return x.ID.CompareTo(y.ID); });
                            p.Value.ForEach(x =>
                            {
                                sl.SetCellValue(3 + i, 1, x.ID);
                                sl.SetCellValue(3 + i, 2, x.Name);
                                sl.SetCellValue(3 + i, 3, x.ProName);
                                if (x.OpenCount > 0)
                                {
                                    sl.SetCellValue(3 + i, 4, x.OpenCount);
                                    sl.SetCellValue(3 + i, 5, x.OpenCount * 50 + x.MultiOpen * 20);
                                }
                                if (x.RecordCount > 0)
                                {
                                    sl.SetCellValue(3 + i, 6, x.RecordCount);
                                    sl.SetCellValue(3 + i, 7, x.RecordCount * 150 + x.MultiRecord * 30);
                                }
                                if (x.OtherCount > 0)
                                {
                                    sl.SetCellValue(3 + i, 8, x.OtherCount);
                                    sl.SetCellValue(3 + i, 9, x.OtherCount * 50);
                                }
                                sl.SetCellValue(3 + i, 10, x.OpenCount * 50 + x.MultiOpen * 20 + x.RecordCount * 150 + x.MultiRecord * 30 + x.OtherCount * 50);
                                if ((x.MultiOpen + x.MultiRecord) > 0)
                                    sl.SetCellValue(3 + i, 11, x.MultiOpen + x.MultiRecord);
                                i++;

                                if (x.ProName.Contains("個管") && x.OtherCount + x.OpenCount + x.RecordCount > 0)
                                {
                                    sl.SetCellValue(j, 13, x.ID);
                                    sl.SetCellValue(j, 14, x.Name);
                                    sl.SetCellValue(j, 15, x.ProName);
                                    sl.SetCellValue(j, 16, x.OtherCount + x.OpenCount + x.RecordCount);
                                    j++;
                                }
                            }
                            );
                        }
                    }
                    sl.SetCellValue(2, 13, "員工代號");
                    sl.SetCellValue(2, 14, "員工名稱");
                    sl.SetCellValue(2, 15, "職稱");
                    sl.SetCellValue(2, 16, "件數(NP)");
                    sl.SetCellValue(2, 17, "總金額 (100)");
                    sl.SetCellValue(3, 16, TotalCount - PI_Count);
                    sl.SetCellValue(3, 17, 100 * (TotalCount - PI_Count));

                    sl.SetCellValue(5, 13, "員工代號");
                    sl.SetCellValue(5, 14, "員工名稱");
                    sl.SetCellValue(5, 15, "職稱");
                    sl.SetCellValue(5, 16, "件數(護理部公款)");
                    sl.SetCellValue(5, 17, "總金額 (100)");
                    sl.SetCellValue(6, 16, TotalCount);
                    sl.SetCellValue(6, 17, 100 * TotalCount);
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
                    MessageBox.Show("獎勵金計算完成!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}

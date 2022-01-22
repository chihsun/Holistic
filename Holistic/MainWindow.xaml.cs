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
            public string PublicStation { get; set; }
            public bool Hospice { get; set; }
            public int Hoscount { get; set; }
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
        /// <summary>
        /// 公基金領取人
        /// </summary>
        public List<Person> Pub_Persons = new List<Person>();
        private void Btn_Cal_Click(object sender, RoutedEventArgs e)
        {
            StationDatas.Clear();
            PersonDatas.Clear();
            MultiCount.Clear();
            PID.Clear();
            PI_Count = 0;
            TotalCount = 0;
            if (Pub_Persons.Count == 0)
            {
                MessageBox.Show("請先讀取公基金名單");
                return;
            }
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
                        for (int j = 0; j < 20; j++)
                        {
                            if (sl.GetCellValueAsString(i + 2, 24 + (j * 3)).Contains("會議記錄完成"))
                                break;
                            if (!string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, 24 + (j * 3)))
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
                            ProName = "開案者",
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
                            Hospice = sl.GetCellValueAsString(i + 2, 135) == "Y" && sl.GetCellValueAsString(i + 2, 13) == "L",
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
                fname = fpath + @"\全人" + DateTime.Now.ToString("yyyy-MM") + ".xlsx";
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
                    sl.SetCellValue(2, 7, "代領人員編");
                    sl.SetCellValue(2, 8, "代領人名稱");
                    sl.SetCellValue(2, 9, "原始總金額");
                    var sort = (from obj in StationDatas orderby obj.Key ascending select obj).ToDictionary(o => o.Key, o => o.Value);
                    int i = 0;
                    foreach (var x in sort)
                    {
                        sl.SetCellValue(3 + i, 1, x.Key);
                        sl.SetCellValue(3 + i, 2, x.Value);
                        var pub_nurse = Pub_Persons.FirstOrDefault(o => o.PublicStation == x.Key);
                        if (pub_nurse != null)
                        {
                            sl.SetCellValue(3 + i, 7, pub_nurse.ID);
                            sl.SetCellValue(3 + i, 8, pub_nurse.Name);
                        }
                        sl.SetCellValue(3 + i, 3, Convert.ToInt32(x.Value) * 200);
                        if (x.Key == "PI")
                        {
                            sl.SetCellValue(3 + i, 4, PI_Count);
                            sl.SetCellValue(3 + i, 5, PI_Count * 100);
                            sl.SetCellValue(3 + i, 9, Convert.ToInt32(x.Value) * 200 + PI_Count * 100);
                        }
                        else
                            sl.SetCellValue(3 + i, 9, Convert.ToInt32(x.Value) * 200);
                        i++;
                    }
                    ///NP公基金
                    sl.SetCellValue(3 + i + 1, 1, "專師");
                    sl.SetCellValue(3 + i + 1, 4, TotalCount - PI_Count);
                    sl.SetCellValue(3 + i + 1, 5, 100 * (TotalCount - PI_Count));
                    sl.SetCellValue(3 + i + 1, 9, 100 * (TotalCount - PI_Count));
                    var np_nurse = Pub_Persons.FirstOrDefault(o => o.PublicStation == "專師");
                    if (np_nurse != null)
                    {
                        sl.SetCellValue(3 + i + 1, 7, np_nurse.ID);
                        sl.SetCellValue(3 + i + 1, 8, np_nurse.Name);
                    }
                    i++;
                    ///護理部公基金
                    sl.SetCellValue(3 + i + 1, 1, "護理部");
                    sl.SetCellValue(3 + i + 1, 4, TotalCount);
                    sl.SetCellValue(3 + i + 1, 5, 100 * TotalCount);
                    sl.SetCellValue(3 + i + 1, 9, 100 * TotalCount);
                    var all_nurse = Pub_Persons.FirstOrDefault(o => o.PublicStation == "護理部");
                    if (all_nurse != null)
                    {
                        sl.SetCellValue(3 + i + 1, 7, all_nurse.ID);
                        sl.SetCellValue(3 + i + 1, 8, all_nurse.Name);
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
                    sl.SetCellValue(2, 10, "02020緩和醫療");
                    sl.SetCellValue(2, 11, "金額 (50)");
                    sl.SetCellValue(2, 12, "護理部公基金");
                    sl.SetCellValue(2, 13, "原始總金額");
                    sl.SetCellValue(2, 14, "罰扣件數");
                    sl.SetCellValue(2, 15, "罰扣金額");
                    sl.SetCellValue(2, 16, "發放總金額");
                    sl.SetCellValue(2, 17, "備註(多職類參與)");
                    List<Person> personcount = new List<Person>();
                    foreach (var x in PersonDatas)
                    {
                        var data = personcount.FirstOrDefault(o => o.ID == x.ID);
                        if (data == null)
                        {
                            personcount.Add(new Person() { ID = x.ID, Name = x.Name, ProName = x.ProName });
                            data = personcount.FirstOrDefault(o => o.ID == x.ID);
                        }
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
                        ///安寧
                        if (x.Hospice)
                            data.Hoscount++;
                    }
                    //修正若公基金領取人不在開案、記錄名單中
                    foreach (var x in Pub_Persons)
                    {
                        var pnurse = personcount.FirstOrDefault(o => o.ID == x.ID);
                        if (pnurse == null)
                        {
                            personcount.Add(new Person() { ID = x.ID, Name = x.Name, ProName = x.ProName });
                        }
                    }
                    /*
                    var pnurse = personcount.FirstOrDefault(o => o.ID == NurseI);
                    if (pnurse == null)
                    {
                        personcount.Add(new Person() { ID = NurseI, Name = "公基金領取人", ProName = "公基金代表", PublicNurse = 100 * TotalCount });
                        if (NurseI == NPI)
                            personcount.FirstOrDefault(o => o.ID == NurseI).PublicNP = 100 * (TotalCount - PI_Count);
                    }
                    var pnp = personcount.FirstOrDefault(o => o.ID == NPI);
                    if (pnp == null)
                    {
                        personcount.Add(new Person() { ID = NPI, Name = "公基金領取人", ProName = "公基金代表", PublicNP = 100 * (TotalCount - PI_Count) });
                    }
                    */
                    /*
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
                            if (x.ID == NurseI)
                                data.PublicNurse = 100 * TotalCount;
                            if (x.ID == NPI)
                                data.PublicNP = 100 * (TotalCount - PI_Count);
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
                    }*/
                    //personcount.Sort((x, y) => { return x.ID.CompareTo(y.ID); });
                    //personcount.Sort((x, y) => { return x.ProName.CompareTo(y.ProName); });
                    var ndata = personcount;
                    ndata.Sort((x, y) => { return x.ProName.CompareTo(y.ProName); });
                    var pdata = ndata.GroupBy(o => o.OtherCount > 0 || (o.OpenCount == 0 && o.RecordCount == 0 && o.OtherCount == 0)).ToDictionary(o => o.Key, o => o.ToList<Person>());
                    //var pdata = personcount.GroupBy(o => o.ProName).ToDictionary(o => o.Key, o => o.ToList<Person>());
                    i = 0;
                    int j = 8;
                    ///全部完成紀錄數
                    int totalMultiCount = 0;
                    int totalHosCount = 0;
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
                                    totalMultiCount += x.OtherCount;
                                }
                                if (x.Hoscount > 0)
                                {
                                    sl.SetCellValue(3 + i, 10, x.Hoscount);
                                    sl.SetCellValue(3 + i, 11, x.Hoscount * 50);
                                    totalHosCount += x.Hoscount;
                                }

                                //公基金
                                int pub_count = 0;
                                bool np = false;
                                foreach (var y in Pub_Persons)
                                {
                                    if (y.ID == x.ID)
                                    {
                                        if (y.PublicStation == "專師")
                                        {
                                            pub_count += 100 * (TotalCount - PI_Count);
                                            sl.SetCellValue(3, 20, x.ID);
                                            sl.SetCellValue(3, 21, x.Name);
                                            sl.SetCellValue(3, 22, x.ProName);
                                            np = true;
                                        }
                                        else if (y.PublicStation == "護理部")
                                        {
                                            pub_count += 100 * TotalCount;
                                            sl.SetCellValue(7, 20, x.ID);
                                            sl.SetCellValue(7, 21, x.Name);
                                            sl.SetCellValue(7, 22, x.ProName);
                                        }
                                        else if (StationDatas.ContainsKey(y.PublicStation))
                                        {
                                            pub_count += StationDatas[y.PublicStation] * 200;
                                            if (y.PublicStation == "PI")
                                                pub_count += PI_Count * 100;
                                        }
                                    }
                                }
                                if (pub_count > 0)
                                {
                                    if (np)
                                    {
                                        sl.SetCellValue(3 + i, 12, "=X3");
                                    }
                                    else
                                    {
                                        sl.SetCellValue(3 + i, 12, pub_count);
                                    }
                                }
                                //原始總金額
                                if (np)
                                    sl.SetCellValue(3 + i, 13, $"=E{3 + i} + G{3 + i} + I{3 + i} + K{3 + i} + L{3 + i}");
                                else
                                    sl.SetCellValue(3 + i, 13, x.OpenCount * 50 + x.MultiOpen * 20 + x.RecordCount * 150 + x.MultiRecord * 30 + x.OtherCount * 50 + x.Hoscount * 50 + pub_count);
                                //sl.SetCellValue(3 + i, 13, $"=E{3 + i} + G{3 + i} + I{3 + i} + K{3 + i} + L{3 + i}");
                                //發放總金額
                                sl.SetCellValue(3 + i, 16, $"=M{3 + i} - O{3 + i}");

                                if ((x.MultiOpen + x.MultiRecord) > 0)
                                    sl.SetCellValue(3 + i, 17, x.MultiOpen + x.MultiRecord);
                                if (sl.GetCellValueAsInt32(3 + i, 13) != 0 || np)
                                    i++;
                                /*
                                if (x.ProName.Contains("個管") && x.OtherCount + x.OpenCount + x.RecordCount > 0)
                                {
                                    sl.SetCellValue(j, 13, x.ID);
                                    sl.SetCellValue(j, 14, x.Name);
                                    sl.SetCellValue(j, 15, x.ProName);
                                    sl.SetCellValue(j, 16, x.OtherCount + x.OpenCount + x.RecordCount);
                                    j++;
                                }
                                */
                            }
                            );
                        }
                    }
                    SLWorksheetStatistics wsstats = sl.GetWorksheetStatistics();

                    sl.SetCellValue(2, 20, "員工代號");
                    sl.SetCellValue(2, 21, "員工名稱");
                    sl.SetCellValue(2, 22, "職稱");
                    sl.SetCellValue(2, 23, "件數(NP公基金)");
                    sl.SetCellValue(2, 24, "總金額 (100)");
                    sl.SetCellValue(3, 23, $"={TotalCount - PI_Count} + W4");
                    sl.SetCellValue(3, 24, $"={100 * (TotalCount - PI_Count)} + X4");
                    sl.SetCellValue(4, 23, "=SUM(N3:N100)");
                    sl.SetCellValue(4, 24, "=SUM(O3:O100)");

                    sl.SetCellValue(6, 20, "員工代號");
                    sl.SetCellValue(6, 21, "員工名稱");
                    sl.SetCellValue(6, 22, "職稱");
                    sl.SetCellValue(6, 23, "件數(護理部公基金)");
                    sl.SetCellValue(6, 24, "總金額 (100)");
                    sl.SetCellValue(7, 23, TotalCount);
                    sl.SetCellValue(7, 24, 100 * TotalCount);

                    sl.SetCellValue(10, 20, "獎勵項目");
                    sl.SetCellValue(10, 23, "件數");
                    sl.SetCellValue(10, 24, "總金額");
                    sl.SetCellValue(11, 20, "護理站協助獎勵金");
                    sl.SetCellValue(12, 20, "紀錄完成者者獎勵金");
                    sl.SetCellValue(13, 20, "專師公基金");
                    sl.SetCellValue(14, 20, "護理部公基金");
                    sl.SetCellValue(15, 20, "開案獎勵金");
                    sl.SetCellValue(16, 20, "職類紀錄獎勵金");
                    sl.SetCellValue(17, 20, "安寧緩和獎勵金");
                    sl.SetCellValue(18, 20, "總計");
                    
                    sl.SetCellValue(11, 23, TotalCount);
                    sl.SetCellValue(11, 24, 200 * TotalCount + 100 * PI_Count);
                    sl.SetCellValue(12, 23, "=W11-W4");
                    sl.SetCellValue(12, 24, $"=SUM(G3:G{wsstats.EndRowIndex - 1}) - X4");
                    sl.SetCellValue(13, 23, "=W3");
                    sl.SetCellValue(13, 24, "=X3");
                    sl.SetCellValue(14, 23, TotalCount);
                    sl.SetCellValue(14, 24, 100 * TotalCount);
                    sl.SetCellValue(15, 23, TotalCount);
                    sl.SetCellValue(15, 24, $"=SUM(E3:E{wsstats.EndRowIndex - 1})");
                    sl.SetCellValue(16, 23, totalMultiCount);
                    sl.SetCellValue(16, 24, 50 * totalMultiCount);
                    sl.SetCellValue(17, 23, totalHosCount);
                    sl.SetCellValue(17, 24, 50 * totalHosCount);
                    sl.SetCellValue(18, 24, "=SUM(X11:X17)");

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

        private void Btn_Pub_Click(object sender, RoutedEventArgs e)
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
                    Pub_Persons = new List<Person>();

                    SLWorksheetStatistics wsstats = sl.GetWorksheetStatistics();
                    for (int i = 0; i < wsstats.EndRowIndex; i++)
                    {
                        if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 2, 1)))
                            break;
                        string station = sl.GetCellValueAsString(i + 2, 1);
                        if (!int.TryParse(sl.GetCellValueAsString(i + 2, 2), out int pid))
                            break;
                        string pname = sl.GetCellValueAsString(i + 2, 3);

                        Pub_Persons.Add(new Person()
                        {
                            ID = pid,
                            Name = pname,
                            ProName = "公基金代領",
                            PublicStation = station
                        });
                    }
                    sl.CloseWithoutSaving();
                }
                if (Pub_Persons.Count > 0)
                {
                    Lb_1.Foreground = new SolidColorBrush(Colors.White);
                    Lb_1.Content = "已匯入公基金領取名單!!";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}

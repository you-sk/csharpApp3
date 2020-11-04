using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NL_SUMMARY_TOOL
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }
        /// <summary>
        /// 現行ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button1_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Filter = "Excelファイル(*.xlsx)|*.xlsx|すべてのファイル(*.*)|*.*",
                Title = "現行システム出力ファイルを選択してください",
                InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            };
            //ダイアログ表示→OKならテキストボックス1へファイルパスをセット
            if (ofd.ShowDialog() == DialogResult.OK) { this.textBox1.Text = ofd.FileName; }
            ofd.Dispose();
        }
        /// <summary>
        /// ボタンclick
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button2_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Filter = "Excelファイル(*.xlsx)|*.xlsx|すべてのファイル(*.*)|*.*",
                Title = "出力ファイル(外注)を選択してください",
                InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            };
            //ダイアログ表示→OKならテキストボックス1へファイルパスをセット
            if (ofd.ShowDialog() == DialogResult.OK) { this.textBox2.Text = ofd.FileName; }
            ofd.Dispose();

        }
        /// <summary>
        /// 集計ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button3_Click(object sender, EventArgs e)
        {
            label1.Text = "";
            label1.Update();

            Console.WriteLine("実行ボタンクリック");
            //ファイル存在チェック
            if (!System.IO.File.Exists(textBox1.Text))
            {
                MessageBox.Show("現行システム出力ファイルがありません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (!System.IO.File.Exists(textBox2.Text))
            {
                MessageBox.Show("出力ファイル(xx)がありません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            bool isSplitGaichu = false;
            if (!String.IsNullOrEmpty(textBox3.Text))
            {
                if (!System.IO.File.Exists(textBox3.Text))
                {
                    MessageBox.Show("出力ファイル(zz)がありません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                isSplitGaichu = true;
            }

            //結果ファイル作成
            DateTime dt = System.DateTime.Now;
            var resultFileName = "result_" + dt.ToString("yyyyMMdd") + ".xlsx";

            var ofd = new SaveFileDialog
            {
                Filter = "Excelファイル(*.xlsx)|*.xlsx|すべてのファイル(*.*)|*.*",
                Title = "比較結果ファイルを指定してください",
                FileName = resultFileName,
                InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            };

            string outputFilePath;
            if (ofd.ShowDialog() == DialogResult.OK) { outputFilePath = ofd.FileName; }
            else
            {   //ファイル指定をキャンセルしたら終了
                label1.Text = "";
                return;
            }

            //ダイアログで上書きを指定していたら消す
            if (System.IO.File.Exists(outputFilePath))
            {
                try
                {   //ファイルをExcelが掴んでいたらエラー
                    System.IO.File.Delete(outputFilePath);
                }
                catch (Exception)
                {
                    MessageBox.Show("ファイルを上書きできません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    label1.Text = "";
                    return;
                }

            }


            var oldxlsx = textBox1.Text;
            var newxlsx = textBox2.Text;
            var newxlsx2 = textBox3.Text;

            //
            var isDataFound = false;
            var sqlConnectionSb = new SQLiteConnectionStringBuilder { DataSource = "db.sqlite" };
            using (var cn = new SQLiteConnection(sqlConnectionSb.ToString()))
            {
                cn.Open();

                using (var cmd = new SQLiteCommand(cn))
                {
                    label1.Text = $"現行データ読み込み開始";
                    label1.Update();

                    //テーブル作成
                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS old_dat(" +
                        "no INTEGER NOT NULL PRIMARY KEY," +
                        "ccode TEXT NOT NULL," +
                        "cname TEXT ," +
                        "tcode TEXT NOT NULL," +
                        "tname TEXT ," +
                        "gcode TEXT NOT NULL," +
                        "gname TEXT ," +
                        "uamount TREAL NOT NULL ," +
                        "gamount TREAL NOT NULL )";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "DELETE FROM old_dat";
                    cmd.ExecuteNonQuery();

                    // Excelファイルを開く(既に開いていてもOK）

                    using (System.IO.FileStream s = new System.IO.FileStream(oldxlsx, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite))
                    using (XLWorkbook workbook = new XLWorkbook(s))
                    {
                        try
                        {
                            foreach (var ws in workbook.Worksheets)
                            {
                                Console.WriteLine(ws.Name);
                                if (ws.Name != "T_実績_総計") { continue; }
                                var StartRownum = 2;
                                var lastUsedRownum = ws.LastCellUsed().Address.RowNumber;
                                if (lastUsedRownum < StartRownum) { continue; }

                                using (var ts = cn.BeginTransaction())
                                {

                                    for (int i = StartRownum; i <= lastUsedRownum; i++)
                                    {
                                        var row = ws.Row(i);
                                        //String.Format("{0:0000}", num)
                                        var no = i;
                                        var ccode = Right("000" + row.Cell(2).Value.ToString().Trim(), 3);
                                        var cname = row.Cell(3).Value.ToString().Trim();
                                        var tcode = Right("00000" + row.Cell(5).Value.ToString().Trim(), 5) + "-" + Right("0000" + row.Cell(6).Value.ToString().Trim(), 4);
                                        var tname = row.Cell(7).Value.ToString().Trim();
                                        var gcode = Right("00000" + row.Cell(8).Value.ToString().Trim(), 5);
                                        var gname = row.Cell(9).Value.ToString().Trim();
                                        var uamount = row.Cell(10).Value.ToString().Trim();
                                        var gamount = row.Cell(11).Value.ToString().Trim();

                                        cmd.CommandText = "INSERT INTO old_dat(no, ccode, cname, tcode, tname ,gcode, gname,uamount,gamount ) VALUES(" +
                                            $"{no}, '{ccode}', '{cname}','{tcode}', '{tname}','{gcode}', '{gname}', {uamount}, {gamount})";
                                        cmd.ExecuteNonQuery();
                                        //Console.WriteLine("INSERT:" + no);
                                        label1.Text = $"現行データ読み込み中 {i}/{lastUsedRownum}";
                                        label1.Update();

                                        isDataFound = true;
                                    }

                                    ts.Commit();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("現行データの読み込みに失敗しました\nファイルを確認してください\n" + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            label1.Text = "";
                            return;
                        }
                    }

                    if (!isDataFound)
                    {
                        MessageBox.Show("現行データが見つかりません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        label1.Text = "";
                        return;
                    }
                    isDataFound = false;

                    label1.Text = $"データ読み込み開始";
                    label1.Update();

                    //テーブル作成
                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS new_dat(" +
                        "no INTEGER NOT NULL PRIMARY KEY," +
                        "ccode TEXT NOT NULL," +
                        "cname TEXT ," +
                        "tcode TEXT NOT NULL," +
                        "tname TEXT ," +
                        "gcode TEXT NOT NULL," +
                        "gname TEXT ," +
                        "uamount TREAL NOT NULL ," +
                        "gamount TREAL NOT NULL )";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "DELETE FROM new_dat";
                    cmd.ExecuteNonQuery();

                    int newId = 0;

                    // Excelファイルを開く(既に開いていてもOK）
                    using (System.IO.FileStream s = new System.IO.FileStream(newxlsx, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite))
                    using (XLWorkbook workbook = new XLWorkbook(s))
                    {
                        try
                        {
                            foreach (var ws in workbook.Worksheets)
                            {
                                Console.WriteLine(ws.Name);
                                if (ws.Name != "Data") { continue; }
                                var StartRownum = 2;
                                var lastUsedRownum = ws.LastCellUsed().Address.RowNumber;
                                if (lastUsedRownum < StartRownum) { continue; }
                                using (var ts = cn.BeginTransaction())
                                {
                                    for (int i = StartRownum; i <= lastUsedRownum; i++)
                                    {
                                        var row = ws.Row(i);
                                        //String.Format("{0:0000}", num)
                                        var no = ++newId;
                                        var cell26 = row.Cell(26).Value.ToString().Trim();
                                        var ccode = "";
                                        var cname = "";
                                        if (cell26 != "")
                                        {
                                            ccode = cell26.Substring(0, 3);
                                            cname = cell26.Substring(4, cell26.Length - 4);
                                        }

                                        var tcode = row.Cell(71).Value.ToString().Trim();
                                        var tname = row.Cell(72).Value.ToString().Trim();

                                        var cell77 = row.Cell(77).Value.ToString().Trim();
                                        var gcode = "00000";
                                        var gname = "";
                                        if (cell77 != "")
                                        {
                                            gcode = cell77.Substring(0, 5);
                                            gname = row.Cell(78).Value.ToString().Trim();
                                        }

                                        var uamount = row.Cell(19).Value.ToString().Trim();
                                        var gamount = row.Cell(22).Value.ToString().Trim();

                                        if (isSplitGaichu)
                                        {
                                            gcode = "";
                                            gname = "";
                                            gamount = "0";
                                        }

                                        cmd.CommandText = "INSERT INTO new_dat(no, ccode, cname, tcode, tname ,gcode, gname,uamount,gamount ) VALUES(" +
                                            $"{no}, '{ccode}', '{cname}','{tcode}', '{tname}','{gcode}', '{gname}', {uamount}, {gamount})";
                                        cmd.ExecuteNonQuery();
                                        Console.WriteLine("INSERT:" + no);
                                        isDataFound = true;
                                        label1.Text = $"データ(zz)読み込み中 {i}/{lastUsedRownum}";
                                        label1.Update();
                                    }
                                    ts.Commit();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("データ(zz)の読み込みに失敗しました\nファイルを確認してください\n" + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            label1.Text = "";
                            return;
                        }
                    }
                    if (!isDataFound)
                    {
                        MessageBox.Show("データ(zz)が見つかりません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        label1.Text = "";
                        return;
                    }

                    if (isSplitGaichu)  //外注データを別Excelから生成するパターン（売上０の外注のみデータを作成する）
                    {
                        // Excelファイルを開く(既に開いていてもOK）
                        using (System.IO.FileStream s = new System.IO.FileStream(newxlsx2, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite))
                        using (XLWorkbook workbook = new XLWorkbook(s))
                        {
                            try
                            {
                                foreach (var ws in workbook.Worksheets)
                                {
                                    Console.WriteLine(ws.Name);
                                    if (ws.Name != "Data") { continue; }
                                    var StartRownum = 2;
                                    var lastUsedRownum = ws.LastCellUsed().Address.RowNumber;
                                    if (lastUsedRownum < StartRownum) { continue; }
                                    using (var ts = cn.BeginTransaction())
                                    {
                                        for (int i = StartRownum; i <= lastUsedRownum; i++)
                                        {
                                            var row = ws.Row(i);
                                            //String.Format("{0:0000}", num)
                                            var no = ++newId;
                                            var cell26 = row.Cell(22).Value.ToString().Trim();  //26⇒22
                                            var ccode = "";
                                            var cname = "";
                                            if (cell26 != "")
                                            {
                                                ccode = cell26.Substring(0, 3);
                                                cname = cell26.Substring(4, cell26.Length - 4);
                                            }

                                            var tcode = ""; // row.Cell(71).Value.ToString().Trim();   //⇒ブランク
                                            var tname = ""; // row.Cell(72).Value.ToString().Trim();   //⇒ブランク

                                            var cell77 = row.Cell(30).Value.ToString().Trim();  //77⇒30外注先コード
                                            var gcode = "00000";
                                            var gname = "";
                                            if (cell77 != "")
                                            {
                                                gcode = cell77.Substring(0, 5);
                                                gname = row.Cell(31).Value.ToString().Trim();   //78⇒31
                                            }

                                            var uamount = "0";// row.Cell(19).Value.ToString().Trim(); //⇒0
                                            var gamount = row.Cell(19).Value.ToString().Trim(); //⇒22⇒19

                                            cmd.CommandText = "INSERT INTO new_dat(no, ccode, cname, tcode, tname ,gcode, gname,uamount,gamount ) VALUES(" +
                                                $"{no}, '{ccode}', '{cname}','{tcode}', '{tname}','{gcode}', '{gname}', {uamount}, {gamount})";
                                            cmd.ExecuteNonQuery();
                                            Console.WriteLine("INSERT:" + no);
                                            isDataFound = true;
                                            label1.Text = $"データ(xx)読み込み中 {i}/{lastUsedRownum}";
                                            label1.Update();
                                        }
                                        ts.Commit();
                                    }

                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("データ(zz)の読み込みに失敗しました\nファイルを確認してください\n" + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                label1.Text = "";
                                return;
                            }
                        }
                        if (!isDataFound)
                        {
                            MessageBox.Show("データ(xx)が見つかりません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            label1.Text = "";
                            return;
                        }
                    }


                    label1.Text = "結果Excel作成中";
                    label1.Update();

                    //ビューを用意
                    cmd.CommandText = "DROP VIEW IF EXISTS g_dat";
                    cmd.ExecuteNonQuery();
                    var sb = new StringBuilder();
                    sb.Append("CREATE VIEW g_dat AS ");
                    sb.Append("select a.gcode, ");
                    sb.Append("       a.ccode, ");
                    sb.Append("       (select gname from old_dat where gcode = a.gcode limit 1) gname_old, ");
                    sb.Append("       (select cname from old_dat where ccode = a.ccode limit 1) cname_old, ");
                    sb.Append("       b.amount                                                  old_amount, ");
                    sb.Append("       (select gname from new_dat where gcode = a.gcode limit 1) gname_new, ");
                    sb.Append("       (select cname from new_dat where ccode = a.ccode limit 1) cname_new, ");
                    sb.Append("       c.amount                                                  new_amount, ");
                    sb.Append("       b.amount - c.amount                                       diff ");
                    sb.Append("from ( ");
                    sb.Append("      (select distinct gcode, ccode ");
                    sb.Append("       from ( ");
                    sb.Append("                select gcode, ccode ");
                    sb.Append("                from old_dat where gamount != 0");
                    sb.Append("                union ");
                    sb.Append("                select gcode, ccode ");
                    sb.Append("                from new_dat where gamount != 0");
                    sb.Append("            )) a ");
                    sb.Append(" ");
                    sb.Append("         left join (select ccode, gcode, sum(gamount) amount from old_dat where gamount != 0 group by ccode, gcode) b ");
                    sb.Append("                   on a.gcode = b.gcode and a.ccode = b.ccode ");
                    sb.Append("         left join (select ccode, gcode, sum(gamount) amount from new_dat where gamount != 0 group by ccode, gcode) c ");
                    sb.Append("                   on a.gcode = c.gcode and a.ccode = c.ccode ");
                    sb.Append("    ) ");
                    sb.Append("order by 1, 2");
                    cmd.CommandText = sb.ToString();
                    cmd.ExecuteNonQuery();

                    //ビューを用意
                    cmd.CommandText = "DROP VIEW IF EXISTS u_dat";
                    cmd.ExecuteNonQuery();
                    sb = new StringBuilder();
                    sb.Append("CREATE VIEW u_dat AS ");
                    sb.Append("select a.tcode, ");
                    sb.Append("       a.ccode, ");
                    sb.Append("       (select tname from old_dat where tcode = a.tcode limit 1) tname_old, ");
                    sb.Append("       (select cname from old_dat where ccode = a.ccode limit 1) cname_old, ");
                    sb.Append("       b.amount                                                  old_amount, ");
                    sb.Append("       (select tname from new_dat where tcode = a.tcode limit 1) tname_new, ");
                    sb.Append("       (select cname from new_dat where ccode = a.ccode limit 1) cname_new, ");
                    sb.Append("       c.amount                                                  new_amount, ");
                    sb.Append("       b.amount - c.amount                                       diff ");
                    sb.Append("from ( ");
                    sb.Append("      (select distinct tcode, ccode ");
                    sb.Append("       from ( ");
                    sb.Append("                select tcode, ccode ");
                    sb.Append("                from old_dat where uamount != 0");
                    sb.Append("                union ");
                    sb.Append("                select tcode, ccode ");
                    sb.Append("                from new_dat where uamount != 0");
                    sb.Append("            )) a ");
                    sb.Append(" ");
                    sb.Append("         left join (select ccode, tcode, sum(uamount) amount from old_dat group by ccode, tcode) b ");
                    sb.Append("                   on a.tcode = b.tcode and a.ccode = b.ccode ");
                    sb.Append("         left join (select ccode, tcode, sum(uamount) amount from new_dat group by ccode, tcode) c ");
                    sb.Append("                   on a.tcode = c.tcode and a.ccode = c.ccode ");
                    sb.Append("    ) ");
                    sb.Append("order by 1, 2");
                    cmd.CommandText = sb.ToString();
                    cmd.ExecuteNonQuery();
                }
                //Console.WriteLine("OK!");
            }

            var resiltWb = new XLWorkbook();
            resiltWb.Style.Font.FontName = "ＭＳ ゴシック";
            //売上
            {
                var worksheet = resiltWb.Worksheets.Add("yyy");
                string[] title = {
                    "xxx","xxxx",
                    "xxx","xxxx",
					"dumy"
                };
                var c = 1;
                var r = 1;
                foreach (var item in title)
                {
                    worksheet.Cell(r, c).Value = item;
                    c++;
                }
                using (var cn = new SQLiteConnection(sqlConnectionSb.ToString()))
                {
                    cn.Open();
                    using (var cmd = new SQLiteCommand(cn))
                    {
                        cmd.CommandText = "SELECT * FROM u_dat";
                        using (SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                c = 1;
                                r++;
                                worksheet.Cell(r, c).Style.NumberFormat.Format = "@";
                                worksheet.Cell(r, c++).Value = reader["tcode"].ToString();
                                worksheet.Cell(r, c).Style.NumberFormat.Format = "@";
                                worksheet.Cell(r, c++).Value = reader["ccode"].ToString();
                                worksheet.Cell(r, c++).Value = reader["tname_old"].ToString();
                                worksheet.Cell(r, c++).Value = reader["cname_old"].ToString();
                                worksheet.Cell(r, c++).Value = reader["old_amount"].ToString();
                                worksheet.Cell(r, c++).Value = reader["tname_new"].ToString();
                                worksheet.Cell(r, c++).Value = reader["cname_new"].ToString();
                                worksheet.Cell(r, c++).Value = reader["new_amount"].ToString();
                                worksheet.Cell(r, c).FormulaR1C1 = "=RC[-4]-RC[-1]";
                            }
                            worksheet.ColumnsUsed().AdjustToContents();
                        }

                    }

                }
            }

            {
                var worksheet = resiltWb.Worksheets.Add("xx");
                string[] title = {
                    "xxx","xxxx",
                    "xxx","xxxx",
					"dumy"
                };
                var c = 1;
                var r = 1;
                foreach (var item in title)
                {
                    worksheet.Cell(r, c).Value = item;
                    c++;
                }
                using (var cn = new SQLiteConnection(sqlConnectionSb.ToString()))
                {
                    cn.Open();
                    using (var cmd = new SQLiteCommand(cn))
                    {
                        cmd.CommandText = "SELECT * FROM g_dat";
                        using (SQLiteDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                c = 1;
                                r++;
                                // .Style.NumberFormat.Format = "@";
                                worksheet.Cell(r, c).Style.NumberFormat.Format = "@";
                                worksheet.Cell(r, c++).Value = reader["gcode"].ToString();
                                worksheet.Cell(r, c).Style.NumberFormat.Format = "@";
                                worksheet.Cell(r, c++).Value = reader["ccode"].ToString();
                                worksheet.Cell(r, c++).Value = reader["gname_old"].ToString();
                                worksheet.Cell(r, c++).Value = reader["cname_old"].ToString();
                                worksheet.Cell(r, c++).Value = reader["old_amount"].ToString();
                                worksheet.Cell(r, c++).Value = reader["gname_new"].ToString();
                                worksheet.Cell(r, c++).Value = reader["cname_new"].ToString();
                                worksheet.Cell(r, c++).Value = reader["new_amount"].ToString();
                                worksheet.Cell(r, c).FormulaR1C1 = "=RC[-4]-RC[-1]";
                            }
                            worksheet.ColumnsUsed().AdjustToContents();
                        }

                    }

                }

            }

            resiltWb.SaveAs(outputFilePath);

            MessageBox.Show("終了しました！", "message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            label1.Text = "終了しました";

        }
        /// <summary>
        /// 閉じるボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private static string Right(string str, int len)
        {
            if (len < 0)
            {
                throw new ArgumentException("引数'len'は0以上でなければなりません。");
            }
            if (str == null)
            {
                return "";
            }
            if (str.Length <= len)
            {
                return str;
            }
            return str.Substring(str.Length - len, len);
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Filter = "Excelファイル(*.xlsx)|*.xlsx|すべてのファイル(*.*)|*.*",
                Title = "ファイルを選択してください",
                InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            };
            //ダイアログ表示→OKならテキストボックス1へファイルパスをセット
            if (ofd.ShowDialog() == DialogResult.OK) { this.textBox3.Text = ofd.FileName; }
            ofd.Dispose();
        }
    }
}

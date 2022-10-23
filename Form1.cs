using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using GhostscriptSharp;    //ライブラリ追加
using iTextSharp.text.pdf; //ライブラリ追加
using Microsoft.WindowsAPICodePack.Dialogs;  //ライブラリ追加

namespace PDFCreator
{
    public partial class Form1 : Form
    {
        string path = "C:\\PDF_kks";
        List<string[]> list = new List<string[]>();
        public CancellationTokenSource cancelTokensource;       //キャンセル判定用

        public Form1()
        {
            InitializeComponent();
        }

        //フォームのLoadイベントハンドラ
        private void Form1_Load(object sender, EventArgs e)
        {
            String strThisProcess = System.Diagnostics.Process.GetCurrentProcess().ProcessName;
            if (System.Diagnostics.Process.GetProcessesByName(strThisProcess).Length > 1)
            {
                MessageBox.Show("既に起動済です", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(0x8020);
            }

            ListView1.Activation = ItemActivation.TwoClick;
            ListView1.ItemDrag += new ItemDragEventHandler(ListView1_ItemDrag);
            ImageList1.ImageSize = new Size(180, 220);


            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            if (Directory.EnumerateFileSystemEntries(path).Any())
            {
                Form1 cf = new Form1();
                cf.AllClear();
            }
        }
        private void TextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13) { e.Handled = true; }
        }
        private void TextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13) { e.Handled = true; }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            ListView1.Clear();
            ImageList1.Images.Clear();
            list.Clear();
            if (!CheckBox1.Checked) { TextBox1.Text = ""; }
            if (!CheckBox2.Checked) { TextBox2.Text = ""; }
            Form1 cf = new Form1();
            cf.AllClear();
        }

        public void AllClear()
        {
            try
            {
                DirectoryInfo di = new DirectoryInfo(path);
                FileInfo[] files = di.GetFiles();
                foreach (FileInfo file in files)
                {
                    file.Delete();
                }
            }
            catch
            {
                MessageBox.Show("ファイル削除に失敗しました", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void Button3_Click(object sender, EventArgs e)
        {
            Form1 cf = new Form1();
            cf.AllClear();
            this.Close();
        }

        /* ①ドラッグ時１回発生するイベント
             アイテムを選択してドラッグしたときに発生する ItemDrag イベントで、DoDragDrop メソッドを呼び出して、ドラッグアンドドロップを開始します。
             移動を実装します。そこで、どのドラッグ操作が発生できるかを示す allowedEffect パラメータに、DragDropEffects.Moveを指定します。*/
        private void ListView1_ItemDrag(Object sender, ItemDragEventArgs e)
        {
            ListView1.DoDragDrop((ListViewItem)e.Item, DragDropEffects.Move);
        }

        // ②ドラッグ開始（移動開始）直後１回発生するイベント
        private void ListView1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(ListViewItem)))
            {
                //Console.WriteLine(e.KeyState);
                if ((e.KeyState & 0x1) > 0) //(e.KeyState & 2) == 2, 0x1(1):ﾏｳｽ左，0x2(2):ﾏｳｽ右,0x4(4):Shift,0x8(8):CTRL,0x10(16):ﾏｳｽ中央,0x20(32):ALT
                {
                    e.Effect = DragDropEffects.Move;
                }
            }
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) //ドラッグ中のデータ（e.Data）の形式がファイル（DataFormats.FileDrop）であることをGetDataPresentメソッドにより確認
            {
                e.Effect = DragDropEffects.Copy; // ドラッグドロップ操作のEffect(e.Effect)を設定する。ここでEffectを変更しないと、以降のイベント（Drop）は発生しない。コピーを許可するようにドラッグ元に通知する。
            }
        }

        // ③アイテムドラッグ中に繰り返し発生するイベント
        // InsertionMark（挿入位置）を入れる
        private void ListView1_DragOver(object sender, DragEventArgs e)
        {
            if (ListView1.SelectedItems.Count != 0)
            {
                Point target = ListView1.PointToClient(new Point(e.X, e.Y));
                int InsertIndex = ListView1.InsertionMark.NearestIndex(target);
                Rectangle itemBounds = ListView1.GetItemRect(0);
                int PointIndex;

                if (InsertIndex > -1) // 自分自身が一番近い場合は「-1」が返ってくる
                {
                    itemBounds = ListView1.GetItemRect(InsertIndex);
                    PointIndex = InsertIndex;
                }
                else
                {
                    int x = target.X;
                    int y = target.Y;
                    ListViewItem lvi = ListView1.GetItemAt(x, y);
                    if (lvi == null)
                    {
                        return;
                    }
                    InsertIndex = ListView1.GetItemAt(x, y).Index;
                    PointIndex = InsertIndex;
                }
                if (target.X < itemBounds.Left + (itemBounds.Width / 2))
                {
                    if (InsertIndex != 0)
                    {
                        InsertIndex--;
                        ListView1.InsertionMark.Index = InsertIndex;
                        ListView1.InsertionMark.AppearsAfterItem = true;
                    }
                    else
                    {
                        ListView1.InsertionMark.Index = InsertIndex;
                        ListView1.InsertionMark.AppearsAfterItem = false;
                    }

                }
                else
                {
                    ListView1.InsertionMark.Index = InsertIndex;
                    ListView1.InsertionMark.AppearsAfterItem = true;
                }

                int OneQuarterSizeHeight = itemBounds.Height / 4;
                int OneEighthSizeHeight = OneQuarterSizeHeight / 2;
                int ListViewHeight = ListView1.Size.Height;
                int ListViewWidth = ListView1.Size.Width;
                int ListViewItemCount = ListView1.Items.Count;

                if ((ListViewHeight - OneQuarterSizeHeight) < target.Y)
                {
                    if ((PointIndex + (ListViewWidth / itemBounds.Width)) < ListViewItemCount - 1)
                    {
                        ListView1.EnsureVisible(PointIndex + (ListViewWidth / itemBounds.Width));
                    }
                    else
                    {
                        ListView1.EnsureVisible(ListViewItemCount - 1);
                    }
                    if ((ListViewHeight - OneEighthSizeHeight) > target.Y)
                    {
                        System.Threading.Thread.Sleep(100);
                    }
                }
                else if (OneQuarterSizeHeight > target.Y)
                {
                    if ((PointIndex - (ListViewWidth / itemBounds.Width)) > 0)
                    {
                        ListView1.EnsureVisible(PointIndex - (ListViewWidth / itemBounds.Width));
                    }
                    else
                    {
                        ListView1.EnsureVisible(0);
                    }
                    if ((OneEighthSizeHeight) < target.Y)
                    {
                        System.Threading.Thread.Sleep(100);
                    }
                }
            }
        }
        // ④ドロップ時１回発生するイベント
        private void ListView1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                List<string> Filename = new List<string>();
                Filename.AddRange((string[])e.Data.GetData(DataFormats.FileDrop));

                cancelTokensource = new CancellationTokenSource();
                var cToken = cancelTokensource.Token;

                using (Form3 f3 = new Form3())
                {
                    f3.f1 = this;
                    Task.Run(() => PdfExport(Filename, cToken)).ContinueWith(_ => f3.Invoke((MethodInvoker)(() => f3.Close())));
                    f3.ShowDialog();
                    GC.Collect();
                }
                if (!cToken.IsCancellationRequested)
                {
                    MessageBox.Show("取り込み完了", "ドラッグ＆ドロップ");
                }
                else
                {
                    MessageBox.Show("キャンセルしました", "中断", MessageBoxButtons.OK);
                }
            }

            //public bool GetDataPresent (Type format); Typeオブジェクトで指定された型でデータを使用、または変換できるかを決定する。
            //typeofで型を取得している。if (e.Data.GetType() == typeof(ListViewItem))でも同じ。
            if (e.Data.GetDataPresent(typeof(ListViewItem)))
            {
                Point targetPoint = ListView1.PointToClient(new Point(e.X, e.Y));

                int InsertIndex = ListView1.InsertionMark.NearestIndex(targetPoint); // 近い位置
                int item_cnt = ListView1.SelectedItems.Count; 　　　　　　　　　　　 // 選択しているアイテム
                int SelectIndex = ListView1.SelectedItems[item_cnt - 1].Index;  　　 //一番最後のインデックス
                int c = ListView1.Items.Count;　　　　　　　　　　　　　　　　　　 　//ListViewItemの数
                string[,] Box = new string[item_cnt, 7];                             //前情報
                int list_Index;
                int k = 0;
                int s = 0;

                if (InsertIndex > -1)
                {
                    Rectangle itemBounds = ListView1.GetItemRect(InsertIndex);
                    if (targetPoint.X > itemBounds.Left + (itemBounds.Width / 2))
                    {
                        InsertIndex++;
                    }
                    for (int i = 0; i < item_cnt; i++)  //前情報
                    {
                        list_Index = ListView1.SelectedItems[i].Index;
                        for (int j = 0; j < 7; j++)
                        {
                            Box[i, j] = list[list_Index][j];
                        }
                    }
                    for (int i = 0; i < item_cnt; i++) //Insertより手前
                    {
                        if (ListView1.SelectedItems[i].Index < InsertIndex)
                        {
                            list.Insert(InsertIndex, new string[7] { Box[i, 0], Box[i, 1], Box[i, 2], Box[i, 3], Box[i, 4], Box[i, 5], Box[i, 6] });
                            list.RemoveAt(ListView1.SelectedItems[i].Index - i);
                        }
                        else
                        {
                            s = i;
                            break;
                        }
                    }
                    for (int i = item_cnt - 1; i >= s; i--)　//Insertより後
                    {
                        if (ListView1.SelectedItems[i].Index > InsertIndex)
                        {
                            list.Insert(InsertIndex, new string[7] { Box[i, 0], Box[i, 1], Box[i, 2], Box[i, 3], Box[i, 4], Box[i, 5], Box[i, 6] });
                            list.RemoveAt(ListView1.SelectedItems[i].Index + k + 1);
                            k++;
                        }
                    }

                    ListView1.BeginUpdate();
                    ListView1.Items.Clear();
                    for (int i = 0; i < list.Count; i++) // 再描画
                    {
                        ListView1.Items.Add(list[i][1] + "_" + (i + 1), int.Parse(list[i][4])); // 回転情報、Iconname、pdfのパス、元のpdfのページ番号、ImageIndex、jpgPath,Size比較結果
                    }
                    ListView1.EndUpdate();
                }
                else
                {
                    ListView1.InsertionMark.Index = -1;
                }
            }
        }

        private void Button4_Click(object sender, EventArgs e)  　　　　　　　　　// ページ削除
        {
            int item_cnt = ListView1.SelectedItems.Count;             　　　　　 // 選択しているアイテム
            if (item_cnt > 0)
            {
                for (int i = item_cnt - 1; i >= 0; i--) //←反対からに修正
                {
                    list.RemoveAt(ListView1.SelectedItems[i].Index);
                }
                ListView1.Items.Clear();
                for (int i = 0; i < list.Count; i++) // 再描画
                {
                    ListView1.Items.Add(list[i][1] + "_" + (i + 1), int.Parse(list[i][4])); // 回転情報、Iconname、pdfのパス、元のpdfのページ番号、ImageIndex、jpgPath,Size比較結果
                }
            }

        }
        private void Button2_Click(object sender, EventArgs e)                               //結合
        {
            if (TextBox1.Text == "" && TextBox2.Text == "") { MessageBox.Show("ファルダ名とファイル名を記入してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
            else if (TextBox1.Text == "") { MessageBox.Show("ファルダ名を記入してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
            else if (TextBox2.Text == "") { MessageBox.Show("ファイル名を記入してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
            if (!Directory.Exists(TextBox1.Text)) { MessageBox.Show("保存先のフォルダが存在しません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
            if (list.Count > 300) { MessageBox.Show("保存枚数は300枚以下まで", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
            if (list.Count == 0) { MessageBox.Show("保存するPDFがありません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

            string joinPdfPath = TextBox1.Text + "\\" + TextBox2.Text + ".pdf";
            string PreName = "";

            if (File.Exists(joinPdfPath))
            {
                DialogResult result = MessageBox.Show("同一のファイル名が保存先にありますが上書きしますか？", "ファイルの上書き確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2); if (result == DialogResult.No) { return; } else { PreName = joinPdfPath; joinPdfPath = TextBox1.Text + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".pdf"; }
            }

            cancelTokensource = new CancellationTokenSource();
            var cToken = cancelTokensource.Token;
            Boolean Error_flag = false;

            using (Form3 f3 = new Form3())
            {
                f3.f1 = this;
                Task.Run(() => PdfSave(joinPdfPath, cToken, ref Error_flag)).ContinueWith(_ => f3.Invoke((MethodInvoker)(() => f3.Close())));
                f3.ShowDialog();
                if (cToken.IsCancellationRequested)
                {
                    MessageBox.Show("キャンセルしました", "中断", MessageBoxButtons.OK);
                    File.Delete(joinPdfPath);
                }
                else
                {
                    if (PreName.Length != 0)
                    {
                        try
                        {
                            File.Delete(PreName);
                            File.Move(joinPdfPath, PreName);
                        }
                        catch
                        {
                            File.Delete(joinPdfPath);
                            MessageBox.Show("上書き保存元のファイルを閉じてください" + Environment.NewLine + PreName, "中断", MessageBoxButtons.OK);
                            return;
                        }

                    }
                    if (!Error_flag) { MessageBox.Show("保存しました", "保存"); }
                }
                GC.Collect();
            }
        }
        private void PdfSave(string JoinPdfPath, CancellationToken cancelToken, ref Boolean error_flag)
        {
            using (iTextSharp.text.Document doc = new iTextSharp.text.Document())                   // Documentクラスのインスタンスを作成
            {
                try
                {
                    using (FileStream fs = new FileStream(JoinPdfPath, FileMode.Create, FileAccess.Write))     // 結合先PDFファイルの作成
                    {
                        try
                        {
                            using (PdfCopy copy = new PdfCopy(doc, fs))
                            {
                                doc.Open();                                                                      // 結合先PDFオープン(文章の出力開始)
                                PdfReader pdf;
                                PdfDictionary page;
                                PdfNumber newRotation;
                                int oldRotation;

                                for (int i = 0; i < list.Count; i++)
                                {
                                    if (cancelToken.IsCancellationRequested)
                                    {
                                        return;
                                    }
                                    pdf = new PdfReader(list[i][2]);
                                    page = pdf.GetPageN(int.Parse(list[i][3]));
                                    oldRotation = pdf.GetPageRotation(int.Parse(list[i][3]));
                                    newRotation = new PdfNumber((int.Parse(list[i][0]) + (oldRotation)) % 360);
                                    page.Put(PdfName.ROTATE, newRotation);
                                    copy.AddPage(copy.GetImportedPage(pdf, int.Parse(list[i][3])));          //copy.AddDocument(pdf) 全ページ
                                    pdf.Close();
                                }
                            }
                        }
                        catch (OutOfMemoryException)
                        {
                            MessageBox.Show("メモリ不足です。\n保存枚数を減らすか他のタスクを閉じてください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error); error_flag = true; return;
                        }
                        catch
                        {
                            MessageBox.Show("保存に失敗しました", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error); error_flag = true; return;
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("保存先と同一ファイル名のPDFを閉じてください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error); error_flag = true; return;
                }
            }
        }
        private void Button7_Click(object sender, EventArgs e)  //フォルダから一括取得
        {
            using (CommonOpenFileDialog dialog = new CommonOpenFileDialog()
            {
                Title = "フォルダを選択してください",
                RestoreDirectory = true,
                IsFolderPicker = true,
            })
            {
                if (dialog.ShowDialog() != CommonFileDialogResult.Ok)
                {
                    return;
                }

                List<string> Filename = new List<string>();
                Filename.AddRange(Directory.GetFiles(dialog.FileName, "*.xls*")); //フォルダ下のEXCELファイル
                Filename.AddRange(Directory.GetFiles(dialog.FileName, "*.pdf"));  //フォルダ下のPDFファイル

                if (Filename.Count > 0)
                {
                    cancelTokensource = new CancellationTokenSource();
                    var cToken = cancelTokensource.Token;

                    using (Form3 f3 = new Form3())
                    {
                        f3.f1 = this;
                        Task.Run(() => PdfExport(Filename, cToken)).ContinueWith(_ => f3.Invoke((MethodInvoker)(() => f3.Close())));
                        f3.ShowDialog();
                        GC.Collect();
                    }
                    if (!cToken.IsCancellationRequested)
                    {
                        MessageBox.Show("取り込み完了", "ファイル一括取得");
                    }
                    else
                    {
                        MessageBox.Show("キャンセルしました", "中断", MessageBoxButtons.OK);
                    }
                }
                else
                {
                    MessageBox.Show("ファイルが見つかりませんでした", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private void PdfExport(List<string> FILENAME, CancellationToken cancelToken)
        {
            foreach (string fileName in FILENAME)
            {
                string ext = Path.GetExtension(fileName);                       //拡張子取得
                string fn = Path.GetFileNameWithoutExtension(fileName);         //拡張子抜きのファイル名

                if (cancelToken.IsCancellationRequested)
                {
                    return;
                }

                if (ext.Contains(".xlsx") || ext.Contains(".xls") || ext.Contains(".xlsm") || ext.Contains(".pdf"))
                {
                    string pdfPath;
                    string jpgPath;
                    string DateTimefn = DateTime.Now.ToString("yyyyMMddHHmmssfff");

                    if (TextBox1.Text == "")
                    {
                        Invoke(new Action<string>(delegate (string FileName) { TextBox1.Text = Path.GetDirectoryName(FileName); }), fileName);
                    }
                    if (TextBox2.Text == "")
                    {
                        Invoke(new Action<string>(delegate (string FileName) { TextBox2.Text = fn; }), fn);
                    }

                    if (ext.Contains(".xlsx") || ext.Contains(".xls") || ext.Contains(".xlsm"))
                    {
                        try
                        {
                            Excel.Application app = null;                                       //Microsoft.Office.Interop.Excel名前空間に属するApplocation型のapp。
                            Excel.Workbook wb = null;
                            try
                            {
                                pdfPath = path + "\\" + DateTimefn + ".pdf";                 //出力するpdfのパス

                                app = new Excel.Application();                              //インスタンス生成
                                app.Visible = false;                                        //非表示
                                wb = app.Workbooks.Open(fileName);

                                wb.ExportAsFixedFormat(                                     //EXCELのpdf出力
                                Excel.XlFixedFormatType.xlTypePDF,
                                pdfPath,                                                    //PDFファイルの出力パス
                                Excel.XlFixedFormatQuality.xlQualityStandard                //出力のクオリティ
                                );

                                wb.Close();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                                wb = null;
                                app.Quit();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                                app = null;
                            }
                            catch
                            {
                                if (wb != null)
                                {
                                    wb.Close();
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                                    wb = null;
                                }
                                if (app != null)
                                {
                                    app.Quit();
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                                    app = null;
                                }
                                MessageBox.Show(fileName + " のPDF出力に失敗しました", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }
                        catch
                        {
                            MessageBox.Show("Microsoft Excelがインストールされてないかバージョンが対応していません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    else
                    {
                        pdfPath = path + "\\" + DateTimefn + ".pdf";
                        File.Copy(fileName, pdfPath);
                        if (TextBox1.Text == "")
                        {
                            Invoke(new Action<string>(delegate (string FileName) { TextBox1.Text = Path.GetDirectoryName(FileName); }), fileName);
                        }
                    }

                    try
                    {
                        PdfReader pdf = new PdfReader(pdfPath);
                        float Height;
                        float Width;
                        Bitmap canvas;
                        Bitmap img = new Bitmap(180, 220);
                        Graphics g = Graphics.FromImage(img);
                        string HW;

                        for (int j = 1; j <= pdf.NumberOfPages; j++)
                        {
                            if (cancelToken.IsCancellationRequested)
                            {
                                g.Dispose();
                                img.Dispose();
                                pdf.Close();
                                return;
                            }

                            jpgPath = path + "\\" + DateTimefn + "_" + j + "_0.jpg";
                            Height = pdf.GetPageSizeWithRotation(j).Height;
                            Width = pdf.GetPageSizeWithRotation(j).Width;
                            try
                            {
                                GhostscriptWrapper.GeneratePageThumb(pdfPath, jpgPath, j, 300, 300);
                            }
                            catch
                            {
                                MessageBox.Show(fileName + " の画像変換に失敗しました。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            canvas = new Bitmap(jpgPath);
                            g.FillRectangle(Brushes.White, g.VisibleClipBounds); //TransparentでもOK

                            if (Height > Width)
                            {
                                g.DrawImage(canvas, 0, 0, 180, 220);
                                HW = "H";
                            }
                            else
                            {
                                g.DrawImage(canvas, 0, 20, 180, 180);
                                HW = "W";
                            }
                            Invoke(new Action<Bitmap>(delegate (Bitmap IMG) { ImageList1.Images.Add(IMG); }), img);
                            canvas.Dispose();
                            Invoke(new Action<string>(delegate (string FN) { ListView1.Items.Add(FN + "_" + (list.Count + 1), ImageList1.Images.Count - 1); }), fn);
                            Invoke(new Action<string, string, int, string, string>(delegate (string FN, string PDFPath, int J, string JPGPath, string hw)
                            { list.Add(new String[7] { "0", FN, PDFPath, J.ToString(), (ImageList1.Images.Count - 1).ToString(), JPGPath.Replace("_0.jpg", ""), hw }); }), fn, pdfPath, j, jpgPath, HW);
                            // 回転情報、Iconname、pdfのパス、元のpdfのページ番号、ImageIndex、jpgPath,Size比較結果
                        }
                        g.Dispose();
                        img.Dispose();
                        pdf.Close();
                    }
                    catch
                    {
                        MessageBox.Show(fileName + " の取り込みに失敗しました。\n" + fileName + " にパスワードがかかっているか、中身が破損しています。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        File.Delete(pdfPath);
                        return;
                    }
                }
            }
        }

        private void Button5_Click(object sender, EventArgs e)     //左回転
        {
            int item_cnt = ListView1.SelectedItems.Count;             　　　　　 // 選択しているアイテム
            if (item_cnt > 0)
            {
                int SelectIndex;
                string jpgPath;
                Bitmap canvas;
                Image img;
                Bitmap bmimg = new Bitmap(180, 220);
                Graphics g = Graphics.FromImage(bmimg);

                for (int i = 0; i < item_cnt; i++)
                {
                    SelectIndex = ListView1.SelectedItems[i].Index;
                    img = Image.FromFile(list[SelectIndex][5] + "_0.jpg");
                    if (list[SelectIndex][0] == "0")
                    {
                        img.RotateFlip(RotateFlipType.Rotate270FlipNone);
                        list[SelectIndex][0] = "270";
                    }
                    else if (list[SelectIndex][0] == "270")
                    {
                        img.RotateFlip(RotateFlipType.Rotate180FlipNone);
                        list[SelectIndex][0] = "180";
                    }
                    else if (list[SelectIndex][0] == "180")
                    {
                        img.RotateFlip(RotateFlipType.Rotate90FlipNone);
                        list[SelectIndex][0] = "90";
                    }
                    else if (list[SelectIndex][0] == "90")
                    {
                        list[SelectIndex][0] = "0";
                    }

                    if (File.Exists(list[SelectIndex][5] + "_" + list[SelectIndex][0] + ".jpg") == false)
                    {
                        img.Save(list[SelectIndex][5] + "_" + list[SelectIndex][0] + ".jpg", ImageFormat.Jpeg);
                    }
                    jpgPath = list[SelectIndex][5] + "_" + list[SelectIndex][0] + ".jpg";

                    canvas = new Bitmap(jpgPath);
                    g.FillRectangle(Brushes.White, g.VisibleClipBounds); //Transparent

                    if (((list[SelectIndex][0] == "90" || list[SelectIndex][0] == "270") && list[SelectIndex][6] == "H") || ((list[SelectIndex][0] == "0" || list[SelectIndex][0] == "180") && list[SelectIndex][6] == "W"))
                    {
                        g.DrawImage(canvas, 0, 20, 180, 180);
                    }
                    else
                    {
                        g.DrawImage(canvas, 0, 0, 180, 220);
                    }

                    ImageList1.Images[int.Parse(list[SelectIndex][4])] = bmimg;

                    canvas.Dispose();
                    img.Dispose();
                }
                bmimg.Dispose();
                g.Dispose();

                ListView1.BeginUpdate();
                ListView1.Items.Clear();
                for (int i = 0; i < list.Count; i++) // 再描画
                {
                    ListView1.Items.Add(list[i][1] + "_" + (i + 1), int.Parse(list[i][4]));  // 回転情報、Iconname、pdfのパス、元のpdfのページ番号、ImageIndex、jpgPath,Size比較結果
                }
                ListView1.EndUpdate();
            }
        }

        private void Button6_Click(object sender, EventArgs e)　   //右回転
        {
            int item_cnt = ListView1.SelectedItems.Count;             　　　　　 // 選択しているアイテム
            if (item_cnt > 0)
            {
                int SelectIndex;
                string jpgPath;
                Bitmap canvas;
                Image img;
                Bitmap bmimg = new Bitmap(180, 220);
                Graphics g = Graphics.FromImage(bmimg);

                for (int i = 0; i < item_cnt; i++)
                {
                    SelectIndex = ListView1.SelectedItems[i].Index;
                    img = Image.FromFile(list[SelectIndex][5] + "_0.jpg");
                    if (list[SelectIndex][0] == "0")
                    {
                        img.RotateFlip(RotateFlipType.Rotate90FlipNone);
                        list[SelectIndex][0] = "90";
                    }
                    else if (list[SelectIndex][0] == "90")
                    {
                        img.RotateFlip(RotateFlipType.Rotate180FlipNone);
                        list[SelectIndex][0] = "180";
                    }
                    else if (list[SelectIndex][0] == "180")
                    {
                        img.RotateFlip(RotateFlipType.Rotate270FlipNone);
                        list[SelectIndex][0] = "270";
                    }
                    else if (list[SelectIndex][0] == "270")
                    {
                        list[SelectIndex][0] = "0";
                    }

                    if (File.Exists(list[SelectIndex][5] + "_" + list[SelectIndex][0] + ".jpg") == false)
                    {
                        img.Save(list[SelectIndex][5] + "_" + list[SelectIndex][0] + ".jpg", ImageFormat.Jpeg);
                    }
                    jpgPath = list[SelectIndex][5] + "_" + list[SelectIndex][0] + ".jpg";

                    canvas = new Bitmap(jpgPath);
                    g.FillRectangle(Brushes.White, g.VisibleClipBounds); //Transparent

                    if (((list[SelectIndex][0] == "90" || list[SelectIndex][0] == "270") && list[SelectIndex][6] == "H") || ((list[SelectIndex][0] == "0" || list[SelectIndex][0] == "180") && list[SelectIndex][6] == "W"))
                    {
                        g.DrawImage(canvas, 0, 20, 180, 180);
                    }
                    else
                    {
                        g.DrawImage(canvas, 0, 0, 180, 220);
                    }
                    ImageList1.Images[int.Parse(list[SelectIndex][4])] = bmimg;
                    canvas.Dispose();
                    img.Dispose();
                }
                bmimg.Dispose();
                g.Dispose();

                ListView1.BeginUpdate();
                ListView1.Items.Clear();
                for (int i = 0; i < list.Count; i++) // 再描画
                {
                    ListView1.Items.Add(list[i][1] + "_" + (i + 1), int.Parse(list[i][4])); // 回転情報、Iconname、pdfのパス、元のpdfのページ番号、ImageIndex、jpgPath,Size比較結果
                }
                ListView1.EndUpdate();
            }
        }

        private void ListView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            int SelectIndex = ListView1.SelectedItems[0].Index;
            Form2 f = new Form2(list[SelectIndex][5], list[SelectIndex][0]);
            f.Show();
            GC.Collect();
        }

        private void ListView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofDialog = new OpenFileDialog()
            {
                Title = "ファイルを選択してください",
                RestoreDirectory = true,
                Filter = "Excel・PDF(*.xls*;*.pdf)|*.xls*;*.pdf"
            })
            {
                {
                    if (ofDialog.ShowDialog() == DialogResult.Cancel) { return; }
                }
                string ext = Path.GetExtension(ofDialog.FileName);

                if (ext.Contains(".xlsx") || ext.Contains(".xls") || ext.Contains(".xlsm") || ext.Contains(".pdf"))
                {
                    List<string> Filename = new List<string>() { ofDialog.FileName };
                    if (Filename.Count > 0)
                    {
                        cancelTokensource = new CancellationTokenSource();
                        var cToken = cancelTokensource.Token;

                        using (Form3 f3 = new Form3())
                        {
                            f3.f1 = this;
                            Task.Run(() => PdfExport(Filename, cToken)).ContinueWith(_ => f3.Invoke((MethodInvoker)(() => f3.Close())));
                            f3.ShowDialog();
                            GC.Collect();
                        }
                        if (!cToken.IsCancellationRequested)
                        {
                            MessageBox.Show("取り込み完了", "ファイル取得");
                        }
                        else
                        {
                            MessageBox.Show("キャンセルしました", "中断", MessageBoxButtons.OK);
                        }
                    }
                }
            }
        }

        private void Button9_Click(object sender, EventArgs e)
        {
            using (CommonOpenFileDialog dialog = new CommonOpenFileDialog()
            {
                Title = "保存先を選択してください",
                RestoreDirectory = true,
                IsFolderPicker = true,
            })
            {
                if (dialog.ShowDialog() != CommonFileDialogResult.Ok)
                {
                    return;
                }
                TextBox1.Text = dialog.FileName;
            }
        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (CheckBox1.Checked)
            {
                if (TextBox1.Text != "")
                {
                    if (!Directory.Exists(TextBox1.Text)) { MessageBox.Show("保存先のフォルダが存在しません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error); CheckBox1.Checked = false; return; }
                }
                else
                {
                    MessageBox.Show("フォルダ名を記入してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error); CheckBox1.Checked = false; return;
                }
                TextBox1.ReadOnly = true;
            }
            else
            {
                TextBox1.ReadOnly = false;
            }
        }

        private void CheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (CheckBox2.Checked)
            {
                if (TextBox2.Text == "")
                {
                    MessageBox.Show("ファイル名を記入してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error); CheckBox2.Checked = false; return;
                }
                TextBox2.ReadOnly = true;
            }
            else
            {
                TextBox2.ReadOnly = false;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            ListView1.Clear();
            ImageList1.Images.Clear();
            list.Clear();
            if (!CheckBox1.Checked) { TextBox1.Text = ""; }
            if (!CheckBox2.Checked) { TextBox2.Text = ""; }
            Form1 cf = new Form1();
            cf.AllClear();
        }
    }
}
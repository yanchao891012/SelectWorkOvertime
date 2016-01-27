using System;
using System.Collections.Generic;
using System.Windows;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System.Threading;
using System.Globalization;
using System.ComponentModel;

namespace SelectWorkOvertime
{    
    /// <summary>
    /// ExportWindow.xaml 的交互逻辑
    /// </summary>
    public partial class ExportWindow : System.Windows.Window
    {
        BackgroundWorker bgWait;
        List<OverTimeModel> otmList = new List<OverTimeModel>();

        public ExportWindow(List<OverTimeModel> list)
        {
            //日历控件初始化
            Thread.CurrentThread.CurrentCulture = new CultureInfo("zh-CN");
            Thread.CurrentThread.CurrentCulture = (CultureInfo)Thread.CurrentThread.CurrentCulture.Clone();
            Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern = "yyyy-MM-dd";

            otmList = list;
            InitializeComponent();
        }

        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="overTimeList"></param>
        /// <param name="fileName"></param>
        public void ExportWord(List<OverTimeModel> overTimeList, string fileName)
        {
            Microsoft.Office.Interop.Word.Application app = null;
            Microsoft.Office.Interop.Word.Document doc = null;

            try
            {
                int rows = overTimeList.Count + 4;//表格行数
                int cols = 6;//表格列数
                object oMissing = System.Reflection.Missing.Value;
                object Visible = true;
                app = new Microsoft.Office.Interop.Word.Application();//创建Word应用程序
                doc = app.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref Visible);//添加Word文档

                //设置页边距
                doc.PageSetup.LeftMargin = 71;
                doc.PageSetup.RightMargin = 71;
                doc.PageSetup.TopMargin = 56.8f;
                doc.PageSetup.BottomMargin = 42.6f;

                //设置页眉页脚为1cm
                app.ActiveDocument.PageSetup.HeaderDistance = app.CentimetersToPoints(1f);
                app.ActiveDocument.PageSetup.FooterDistance = app.CentimetersToPoints(1f);

                //设置页眉内容
                string filename = System.Windows.Forms.Application.StartupPath + "\\logo.png";//设置图片路径
                object linkToFile = false;//定义该插入的图片是否为外部链接
                object saveWithDocument = true;//定义要插入的图片是否随Word文档一起保存
                object headerFooterRange;//在Word中插入的位置
                foreach (Section section in doc.Sections)
                {
                    foreach (HeaderFooter headerFooter in section.Headers)
                    {
                        headerFooterRange = headerFooter.Range;//在页眉中插入
                        headerFooter.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;//设置图片位置为居左                        
                        doc.InlineShapes.AddPicture(filename, ref linkToFile, ref saveWithDocument, ref headerFooterRange);//插入图片
                    }
                }

                //输入大标题加粗加大字号水平居中
                app.Selection.Font.Bold = 700;
                app.Selection.Font.Size = 15;
                app.Selection.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                app.Selection.Text = "加班申请单";

                //换行添加表格
                object line = WdUnits.wdLine;
                app.Selection.MoveDown(ref line, oMissing, oMissing);
                app.Selection.TypeParagraph();//换行
                Range range = app.Selection.Range;
                Table table = app.Selection.Tables.Add(range, rows, cols, ref oMissing, ref oMissing);

                table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;

                //设置表格的字体大小
                table.Range.Font.Size = 10.5f;
                table.Range.Font.Bold = 0;

                //设置表格中的第二行为加粗
                table.Rows[2].Range.Font.Bold = 700;

                //设置表格的行高
                table.Rows.Height = 25f;
                table.Rows.HeightRule = WdRowHeightRule.wdRowHeightExactly;

                //设置表格列的宽度
                table.Columns[1].Width = 75f;
                table.Columns[2].Width = 95f;
                table.Columns[3].Width = 95f;
                table.Columns[4].Width = 95f;
                table.Columns[5].Width = 70f;
                table.Columns[6].Width = 85f;

                //设置Table居中
                table.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;

                //设置第一行
                int rowIndex = 1;
                table.Cell(rowIndex, 1).Range.Text = "申请人";
                //table.Cell(rowIndex, 2).Range.Text = txtSqr.Text;
                this.txtSqr.Dispatcher.Invoke(new Action(() => table.Cell(rowIndex, 2).Range.Text = this.txtSqr.Text));
                table.Cell(rowIndex, 3).Range.Text = "所属部门";
                //table.Cell(rowIndex, 4).Range.Text = txtSsbm.Text;
                this.txtSsbm.Dispatcher.Invoke(new Action(() => table.Cell(rowIndex, 4).Range.Text = this.txtSsbm.Text));
                table.Cell(rowIndex, 5).Range.Text = "申请时间";
                //table.Cell(rowIndex, 6).Range.Text = dpSqrq.Text;
                this.dpSqrq.Dispatcher.Invoke(new Action(() => table.Cell(rowIndex, 6).Range.Text = this.dpSqrq.Text));

                //设置第二行
                rowIndex = 2;
                table.Cell(rowIndex, 1).Range.Text = "姓名";
                table.Cell(rowIndex, 2).Range.Text = "加班类型";
                table.Cell(rowIndex, 3).Range.Text = "加班时间起";
                table.Cell(rowIndex, 4).Range.Text = "加班时间止";
                table.Cell(rowIndex, 5).Range.Text = "小时数";
                table.Cell(rowIndex, 6).Range.Text = "备注";

                //循环添加数据                
                foreach (var i in overTimeList)
                {
                    rowIndex++;
                    table.Cell(rowIndex, 1).Range.Text = i.Name;
                    table.Cell(rowIndex, 2).Range.Text = i.OvertimeType;
                    table.Cell(rowIndex, 3).Range.Text = i.OvertimeStart;
                    table.Cell(rowIndex, 4).Range.Text = i.OvertimeEnd;
                    table.Cell(rowIndex, 5).Range.Text = i.OvertimeHours;
                    table.Cell(rowIndex, 6).Range.Text = i.Remark;
                }

                //倒数第二行
                rowIndex++;
                table.Cell(rowIndex, 1).Range.Text = "部门经理";
                table.Cell(rowIndex, 2).Merge(table.Cell(rowIndex, 3));
                table.Cell(rowIndex, 3).Range.Text = "副总经理";
                table.Cell(rowIndex, 4).Merge(table.Cell(rowIndex, 5));
                //倒数第一行
                rowIndex++;
                table.Cell(rowIndex, 1).Range.Text = "常务副总经理";
                table.Cell(rowIndex, 2).Merge(table.Cell(rowIndex, 3));
                table.Cell(rowIndex, 3).Range.Text = "总经理";
                table.Cell(rowIndex, 4).Merge(table.Cell(rowIndex, 5));

                doc.SaveAs(fileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                MessageBox.Show("导出成功！");
                this.Dispatcher.Invoke(new Action(() => this.Close()));
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close();
                }
                if (app != null)
                {
                    app.Quit();
                }
            }
        }

        SaveFileDialog sfd = new SaveFileDialog();
        
        /// <summary>
        /// 导出按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnEx_Click(object sender, RoutedEventArgs e)
        {            
            sfd.Filter = "Word文档(*.doc)|*.doc";
            sfd.FileName = "加班申请单" + dpSqrq.Text;
            if (sfd.ShowDialog() == true)
            {
                bgWait = new BackgroundWorker();
                bgWait.WorkerReportsProgress = true;
                bgWait.DoWork += BgWait_DoWork;
                bgWait.RunWorkerCompleted += BgWait_RunWorkerCompleted;
                bgWait.RunWorkerAsync();                
            }            
        }

        private void BgWait_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            DpWait.Visibility = Visibility.Collapsed;
            //throw new NotImplementedException();
        }


        private void BgWait_DoWork(object sender, DoWorkEventArgs e)
        {
            this.Dispatcher.Invoke(new Action(
                () =>
                {
                    DpWait.Visibility = Visibility.Visible;
                }));
            ExportWord(otmList, sfd.FileName);
            //throw new NotImplementedException();
        }
    }
}

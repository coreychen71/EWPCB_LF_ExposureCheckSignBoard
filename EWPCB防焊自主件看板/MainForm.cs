using System;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace EWPCB防焊自主件看板
{
    public partial class MainForm : Form
    {
        string strCon = "server=EWNAS;database=ME;uid=me;pwd=2dae5na";
        string strComm = "";
        string LogPath = Directory.GetCurrentDirectory() + @"\ErrLog.txt";
        DataTable srcData = new DataTable();
        DateTime date = DateTime.Now;
        DateTime updateTime = DateTime.Now;
        int setTime = 180;
        int clock = 0;
        StreamWriter writeLog;
        public MainForm()
        {
            InitializeComponent();
            dgvData.ReadOnly = true;
            DataRefresh(date.ToString("yyyy-MM-dd 00:00:00"));
            clock = setTime;
            Text = "防焊曝光自主件檢驗看板(" + clock + ") - " + updateTime.ToString("yyyy-MM-dd HH:mm:ss");
            tmrRefresh.Interval = 1000;
            tmrRefresh.Start();
        }

        private void DataRefresh(string date)
        {
            srcData.Clear();
            writeLog = File.AppendText(LogPath);
            strComm = "select partnum as '料號',machineno+'-'+empname as '曝光手',workqnty as '檢修數',qcresult as '結果'," +
                "qcman as '檢驗人',CONVERT(char(19), starttime, 120) as '放板時間',CONVERT(char(19),endtime,120) " +
                "as '結束時間' from drymcse where departname = 'LF' and process = '自主件' and todo = 1 and " +
                "starttime >='" + date + "' order by starttime desc";
            using (SqlConnection sqlcon = new SqlConnection(strCon))
            {
                using (SqlCommand sqlcomm = new SqlCommand(strComm, sqlcon))
                {
                    try
                    {
                        sqlcon.Open();
                        SqlDataReader read = sqlcomm.ExecuteReader();
                        srcData.Load(read);
                        /*
                        writeLog.WriteLine(updateTime + "     " + "Update OK！");
                        writeLog.Flush();
                        */
                        writeLog.Close();
                    }
                    catch (Exception ex)
                    {
                        writeLog.WriteLine(updateTime + "     " + ex.Message);
                        writeLog.Flush();
                        writeLog.Close();
                    }
                }
            }
            //1920*1080
            dgvData.DataSource = srcData;
            dgvData.RowHeadersWidth = 70;
            dgvData.Columns["料號"].Width = 285;
            dgvData.Columns["曝光手"].Width = 175;
            dgvData.Columns["檢修數"].Width = 175;
            dgvData.Columns["結果"].Width = 250;
            dgvData.Columns["檢驗人"].Width = 200;
            dgvData.Columns["放板時間"].Width = 400;
            dgvData.Columns["結束時間"].Width = 400;
            dgvData.DataBindingComplete += ChangRowColor;
        }

        //在資料繫結後，所觸發的事件
        private void ChangRowColor(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            foreach (DataGridViewRow row in dgvData.Rows)
            {
                if (Convert.ToString(row.Cells["結果"].Value).Trim().ToUpper() == "OK")
                {
                    row.DefaultCellStyle.BackColor = Color.Lime;
                    row.DefaultCellStyle.SelectionBackColor = Color.Lime;
                    row.DefaultCellStyle.SelectionForeColor = Color.Black;
                    row.Height = 50;
                }
                else if (Convert.ToString(row.Cells["結果"].Value).Trim().Contains("曝偏"))
                {
                    row.DefaultCellStyle.BackColor = Color.Magenta;
                    row.DefaultCellStyle.SelectionBackColor = Color.Magenta;
                    row.DefaultCellStyle.SelectionForeColor = Color.Black;
                    row.Height = 50;
                }
                else if (Convert.ToString(row.Cells["結果"].Value).Trim().ToUpper().Contains("SMD") &&
                    Convert.ToString(row.Cells["結果"].Value).Trim().Contains("smd"))
                {
                    row.DefaultCellStyle.BackColor = Color.Red;
                    row.DefaultCellStyle.SelectionBackColor = Color.Red;
                    row.DefaultCellStyle.SelectionForeColor = Color.Black;
                    row.Height = 50;
                }
                else if (Convert.ToString(row.Cells["結果"].Value).Trim().ToUpper() != "")
                {
                    row.DefaultCellStyle.BackColor = Color.LightSalmon;
                    row.DefaultCellStyle.SelectionBackColor = Color.LightSalmon;
                    row.DefaultCellStyle.SelectionForeColor = Color.Black;
                    row.Height = 50;
                }
                else
                {
                    row.DefaultCellStyle.BackColor = Color.White;
                    row.DefaultCellStyle.SelectionBackColor = Color.White;
                    row.DefaultCellStyle.SelectionForeColor = Color.Black;
                    row.Height = 50;
                }
            }
        }

        private void tmrRefresh_Tick(object sender, EventArgs e)
        {
            clock--;
            Text = "防焊曝光自主件檢驗看板(" + clock + ") - " + updateTime.ToString("yyyy-MM-dd HH:mm:ss");
            if (clock == 0)
            {
                date = DateTime.Now;
                updateTime = DateTime.Now;
                DataRefresh(date.ToString("yyyy-MM-dd 00:00:00"));
                clock = setTime;
                Text = "防焊曝光自主件檢驗看板(" + clock + ") - " + updateTime.ToString("yyyy-MM-dd HH:mm:ss");
            }
        }
    }
}


using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Linq;
using OpcRcw.Da;
using System.IO;
using System.Diagnostics;


namespace ProgOpe
{
    public partial class ProgOpeForm : Form
    {
        // OPCサーバーアクセスインスタンス
        DxpSimpleAPI.DxpSimpleClass opc = new DxpSimpleAPI.DxpSimpleClass();

        // ステップ数
        const int STEP_NUM = 20;

        // 格納先レジスタ起点No
        const int TEM_POS = 1500;
        const int MST_POS = 1550;
        const int SEC_POS = 1700;
        const int UNIT_BIT_POS = 1500;
        const int ENABLE_BIT_POS = 1650;
        const int LOWER_TIME_POS = 1600;
        const int UPPER_TIME_POS = 1650;
        const int STEP_OFFSET = 50;
        // 2014/08/14 T.Yamanaka 繰り返し回数&ポーズ追加 開始
        const int PAUSE_BIT_POS = 1700;
        const int REP_POS = 2050;
        // 2014/08/14 T.Yamanaka 繰り返し回数&ポーズ追加 終了

        // スケールレンジ
        const double TMP_LOW = -100.0;
        const double TMP_HIGH = 100.0;
        const double MST_LOW = -0.0;
        const double MST_HIGH = 100.0;

        // 入力可能レンジ
        const double TMP_MIN = -45.0, TMP_MAX = 65.0;
        const double MST_MIN =  30.0, MST_MAX = 95.0;

        // OPCサーバーアクセス用固定文字列
        const string DEV_NAME = "DEV1";
        const string WORD_REG_PREFIX = "D";
        const string BIT_REG_PREFIX = "M";

        // DCSアクセス用変数
        const int MAX_DIGIT = 10000;

        // 処理円滑化用配列
        List<ComboBox> unitCmbs = new List<ComboBox>();
        List<TextBox> tmpTxts = new List<TextBox>();
        List<TextBox> mstTxts = new List<TextBox>();
        List<TextBox> timTxts = new List<TextBox>();
        // 2014/08/14 T.Yamanaka 繰り返し回数&ポーズ追加 開始
        List<ComboBox> pauseCmbs = new List<ComboBox>();
        // 2014/08/14 T.Yamanaka 繰り返し回数&ポーズ追加 終了

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public ProgOpeForm()
        {
            InitializeComponent();

            // 処理円滑化用配列を構築
            SetUnitCmbAry();
            SetTmpTxtAry();
            SetMstTxtAry();
            SetTimTxtAry();
            // 2014/08/14 T.Yamanaka 繰り返し回数&ポーズ追加 開始
            SetPauseCmbAry();
            // 2014/08/14 T.Yamanaka 繰り返し回数&ポーズ追加 終了

            for (int i = 0; i < STEP_NUM; i++)
            {
                // 対応ステップナンバーをタグで登録
                // 2014/08/14 T.Yamanaka 繰り返し回数&ポーズ追加 開始
                // timTxts[i].Tag = tmpTxts[i].Tag = unitCmbs[i].Tag = i + 1;
                timTxts[i].Tag = tmpTxts[i].Tag = unitCmbs[i].Tag = pauseCmbs[i].Tag = i + 1;
                // 2014/08/14 T.Yamanaka 繰り返し回数&ポーズ追加 終了
                timTxts[i].TextAlign = tmpTxts[i].TextAlign = HorizontalAlignment.Right;

                // ボックスバリデーション用イベントハンドラー追加
                unitCmbs[i].SelectedIndexChanged += UnitCmb_SelectedIndexChanged;
                timTxts[i].Validating += timTxt_Validating;
                tmpTxts[i].Validating += tmpTxt_Validating;
            
                if (i > 0)  // 湿度のみ最初１つがReadonlyのため専用処理
                {
                    mstTxts[i - 1].Tag = i + 1;
                    mstTxts[i - 1].TextAlign = HorizontalAlignment.Right;
                    mstTxts[i - 1].Validating += mstTxt_Validating;
                }
            }
            ResetTable();
        }

        private void SetPauseCmbAry()
        {
            pauseCmbs.Add(PauseCmb1);
            pauseCmbs.Add(PauseCmb2);
            pauseCmbs.Add(PauseCmb3);
            pauseCmbs.Add(PauseCmb4);
            pauseCmbs.Add(PauseCmb5);
            pauseCmbs.Add(PauseCmb6);
            pauseCmbs.Add(PauseCmb7);
            pauseCmbs.Add(PauseCmb8);
            pauseCmbs.Add(PauseCmb9);
            pauseCmbs.Add(PauseCmb10);
            pauseCmbs.Add(PauseCmb11);
            pauseCmbs.Add(PauseCmb12);
            pauseCmbs.Add(PauseCmb13);
            pauseCmbs.Add(PauseCmb14);
            pauseCmbs.Add(PauseCmb15);
            pauseCmbs.Add(PauseCmb16);
            pauseCmbs.Add(PauseCmb17);
            pauseCmbs.Add(PauseCmb18);
            pauseCmbs.Add(PauseCmb19);
            pauseCmbs.Add(PauseCmb20);
        }

        /// <summary>
        /// 温度テキストボックスバリデーション
        /// </summary>
        /// <param name="sender">メッセージ元</param>
        /// <param name="e">イベント</param>
        private void tmpTxt_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            string errorTxt = string.Format("{0}～{1}の数値を入力してください。", TMP_MIN, TMP_MAX);
            TextBox tb = sender as TextBox;
            if (tb == null)
            {
                Debug.Assert(false, "テキストボックス認識失敗");
                return;
            }

            double ret = 0;
            if (!double.TryParse(tb.Text, out ret))
            {
                //  e.Cancel = true;
                tb.Text = "0";
                errorProvider.SetError(tb, errorTxt);
            }
            else
            {
                if (ret < TMP_MIN || TMP_MAX < ret)
                {
                    // e.Cancel = true;
                    tb.Text = "0";
                    errorProvider.SetError(tb, errorTxt);
                }
                else
                {
                    errorProvider.SetError(tb, null);
                }
            }
        }

        /// <summary>
        /// 時間テキストボックスバリデーション
        /// </summary>
        /// <param name="sender">メッセージ元</param>
        /// <param name="e">イベント</param>
        private void timTxt_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            string errorTxt = "8時間を超えない数値を入力してください。";

            int[] MAX_TIME = { 28800, 480, 8 };
            TextBox tb = sender as TextBox;
            if (tb == null)
            {
                Debug.Assert(false, "テキストボックス認識失敗");
                return;
            }


            int ret = 0;
            if (!int.TryParse(tb.Text, out ret))
            {
                //e.Cancel = true;
                tb.Text = "0";
                errorProvider.SetError(tb, errorTxt);
            }
            else
            {
                int n = Convert.ToInt32(tb.Tag)-1;
                int s = unitCmbs[n].SelectedIndex;
                if (ret > MAX_TIME[s])
                {
                    //e.Cancel = true;
                    tb.Text = "0";
                    errorProvider.SetError(tb, errorTxt);
                }
                else
                {
                    errorProvider.SetError(tb, null);
                }
            }
        }

        /// <summary>
        /// 湿度テキストボックスバリデーション
        /// </summary>
        /// <param name="sender">メッセージ元</param>
        /// <param name="e">イベント</param>
        private void mstTxt_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            string errorTxt = string.Format("{0}～{1}の数値、又は湿度運転無しを示す数値0を入力してください。", MST_MIN, MST_MAX);
            
            TextBox tb = sender as TextBox;
            if (tb == null)
            {
                Debug.Assert(false, "テキストボックス認識失敗");
                return;
            }

            double ret = 0;
            if (!double.TryParse(tb.Text, out ret))
            {
                //  e.Cancel = true;
                tb.Text = "0";
                errorProvider.SetError(tb, errorTxt);
            }
            else
            {
                if ( ret <= 0 ) // 湿度設定無し
                {
                    errorProvider.SetError(tb, null);
                }
                else if (ret < MST_MIN || MST_MAX < ret)
                {
                    // e.Cancel = true;
                    tb.Text = "0";
                    errorProvider.SetError(tb, errorTxt);
                }
                else            // 湿度設定有り
                {
                    errorProvider.SetError(tb, null);
                }
            }
        }

        /// <summary>
        /// 時間単位設定コンボボックス変化イベント
        /// </summary>
        /// <param name="sender">メッセージ元</param>
        /// <param name="e">イベント</param>
        private void UnitCmb_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = sender as ComboBox;
            if (cb == null)
            {
                Debug.Assert(false, "コンボボックス認識失敗");
                return;
            }

            int ret = Convert.ToInt32(cb.Tag) - 1;
            timTxts[ret].Text = "0";
        }

        /// <summary>
        /// 設定値初期化
        /// </summary>
        private void ResetTable()
        {
            foreach (ComboBox cmb in unitCmbs)
            {
                cmb.SelectedIndex = 0;
                errorProvider.SetError(cmb, null);
            }
            foreach (ComboBox cmb in pauseCmbs)
            {
                cmb.SelectedIndex = 0;
                errorProvider.SetError(cmb, null);
            }
            ResetTexts(tmpTxts, "0");
            ResetTexts(mstTxts, "0");
            ResetTexts(timTxts, "0");
        }

        private void ResetTexts(List<TextBox> txts, string val)
        {
            foreach (TextBox txt in txts)
            {
                txt.Text = val;
                errorProvider.SetError(txt, null);
            }
        }

        /// <summary>
        /// 時間設定用テキストボックスを配列登録
        /// </summary>
        private void SetTimTxtAry()
        {
            timTxts.Add(TimTxt1);
            timTxts.Add(TimTxt2);
            timTxts.Add(TimTxt3);
            timTxts.Add(TimTxt4);
            timTxts.Add(TimTxt5);
            timTxts.Add(TimTxt6);
            timTxts.Add(TimTxt7);
            timTxts.Add(TimTxt8);
            timTxts.Add(TimTxt9);
            timTxts.Add(TimTxt10);
            timTxts.Add(TimTxt11);
            timTxts.Add(TimTxt12);
            timTxts.Add(TimTxt13);
            timTxts.Add(TimTxt14);
            timTxts.Add(TimTxt15);
            timTxts.Add(TimTxt16);
            timTxts.Add(TimTxt17);
            timTxts.Add(TimTxt18);
            timTxts.Add(TimTxt19);
            timTxts.Add(TimTxt20);
        }

        /// <summary>
        /// 湿度設定用テキストボックスを配列登録
        /// </summary>
        private void SetMstTxtAry()
        {
            //mstTxts.Add(MstTxt1); // MstTxt1はReadOnly=trueな為スキップ
            mstTxts.Add(MstTxt2);
            mstTxts.Add(MstTxt3);
            mstTxts.Add(MstTxt4);
            mstTxts.Add(MstTxt5);
            mstTxts.Add(MstTxt6);
            mstTxts.Add(MstTxt7);
            mstTxts.Add(MstTxt8);
            mstTxts.Add(MstTxt9);
            mstTxts.Add(MstTxt10);
            mstTxts.Add(MstTxt11);
            mstTxts.Add(MstTxt12);
            mstTxts.Add(MstTxt13);
            mstTxts.Add(MstTxt14);
            mstTxts.Add(MstTxt15);
            mstTxts.Add(MstTxt16);
            mstTxts.Add(MstTxt17);
            mstTxts.Add(MstTxt18);
            mstTxts.Add(MstTxt19);
            mstTxts.Add(MstTxt20);
        }

        /// <summary>
        /// 時間単位設定用コンボボックスを配列登録
        /// </summary>
        private void SetUnitCmbAry()
        {
            unitCmbs.Add(UnitCmb1);
            unitCmbs.Add(UnitCmb2);
            unitCmbs.Add(UnitCmb3);
            unitCmbs.Add(UnitCmb4);
            unitCmbs.Add(UnitCmb5);
            unitCmbs.Add(UnitCmb6);
            unitCmbs.Add(UnitCmb7);
            unitCmbs.Add(UnitCmb8);
            unitCmbs.Add(UnitCmb9);
            unitCmbs.Add(UnitCmb10);
            unitCmbs.Add(UnitCmb11);
            unitCmbs.Add(UnitCmb12);
            unitCmbs.Add(UnitCmb13);
            unitCmbs.Add(UnitCmb14);
            unitCmbs.Add(UnitCmb15);
            unitCmbs.Add(UnitCmb16);
            unitCmbs.Add(UnitCmb17);
            unitCmbs.Add(UnitCmb18);
            unitCmbs.Add(UnitCmb19);
            unitCmbs.Add(UnitCmb20);
        }

        /// <summary>
        /// 温度設定用テキストボックスを配列登録
        /// </summary>
        private void SetTmpTxtAry()
        {
            tmpTxts.Add(TmpTxt1);
            tmpTxts.Add(TmpTxt2);
            tmpTxts.Add(TmpTxt3);
            tmpTxts.Add(TmpTxt4);
            tmpTxts.Add(TmpTxt5);
            tmpTxts.Add(TmpTxt6);
            tmpTxts.Add(TmpTxt7);
            tmpTxts.Add(TmpTxt8);
            tmpTxts.Add(TmpTxt9);
            tmpTxts.Add(TmpTxt10);
            tmpTxts.Add(TmpTxt11);
            tmpTxts.Add(TmpTxt12);
            tmpTxts.Add(TmpTxt13);
            tmpTxts.Add(TmpTxt14);
            tmpTxts.Add(TmpTxt15);
            tmpTxts.Add(TmpTxt16);
            tmpTxts.Add(TmpTxt17);
            tmpTxts.Add(TmpTxt18);
            tmpTxts.Add(TmpTxt19);
            tmpTxts.Add(TmpTxt20);
        }


        // ロード時OPCサーバーリストを更新する
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                btnConnect.Enabled = true;
                btnDisconnect.Enabled = false;

                if (opc.Connect(txtNode.Text, cmbServerList.Text))
                {
                    btnListRefresh.Enabled = false;
                    btnDisconnect.Enabled = true;
                    btnConnect.Enabled = false;
                }
                else
                {
                    DataHelper.ErrorLog("OPC connect() failed in Load()");
                }
            }
            catch (Exception ex)
            {
                // Log Error
                DataHelper.ErrorLog( ex.Message );
            }
        }

        // 接続
        private void btnConnect_Click(object sender, EventArgs e)
        {
            try
            {
                if (opc.Connect(txtNode.Text, cmbServerList.Text))
                {
                    btnListRefresh.Enabled = false;
                    btnDisconnect.Enabled = true;
                    btnConnect.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                // Log Error
                DataHelper.ErrorLog( ex.Message );
            }
        }

        // 接続解除
        private void btnDisconnect_Click(object sender, EventArgs e)
        {
            try
            {
                if (opc.Disconnect())
                {
                    btnConnect.Enabled = true;
                    btnListRefresh.Enabled = true;
                    btnDisconnect.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                // Log Error
                DataHelper.ErrorLog(ex.Message);
            }
        }

        // ノードに対するサーバーリストのリフレッシュ
        private void btnListRefresh_Click(object sender, EventArgs e)
        {
            cmbServerList.Items.Clear();
            string[] ServerNameArray;

            try
            {
                opc.EnumServerList(txtNode.Text, out ServerNameArray);

                for (int a = 0; a < ServerNameArray.Count<string>(); a++)
                {
                    cmbServerList.Items.Add(ServerNameArray[a]);
                }
                cmbServerList.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                // Log Error
                DataHelper.ErrorLog(ex.Message);
            }
        }

        /// <summary>
        /// フォーム表示時のイベント
        /// </summary>
        /// <param name="sender">メッセージ元</param>
        /// <param name="e">イベント</param>
        private void Form1_Shown(object sender, EventArgs e)
        {
            if (!CheckStatus())
            {
                this.Close();
            }

            ReadValues(tmpTxts, TEM_POS, TMP_LOW, TMP_HIGH);
            foreach (TextBox tb in tmpTxts)
            {
                if (Convert.ToDouble(tb.Text) <= TMP_LOW)
                {
                    tb.Text = "0";
                }
            }

            ReadValues(mstTxts, MST_POS, MST_LOW, MST_HIGH);
            foreach (TextBox tb in mstTxts)
            {
                if (Convert.ToDouble(tb.Text) <= MST_LOW)
                {
                    tb.Text = "0";
                }
            }

            ReadTimeBits();
            ReadTimeValues();

            List<string> targetRegs = new List<string>();
            object[] readVals;
            short[] qlty;
            int[] errs;
            FILETIME[] ft;

            targetRegs.Add(string.Format("{0}.{1}{2}", DEV_NAME, WORD_REG_PREFIX, 1062));
            if (opc.Read(targetRegs.ToArray(), out readVals, out qlty, out ft, out errs))
            {
                try
                {
                    int wk = Convert.ToInt32(readVals[0]);
                    int currentStep = ( wk >= STEP_NUM ) ? STEP_NUM : wk;
                    
                    for (int step = 0; step < currentStep; step++)
                    {
                        if (step >= 1)
                        {
                            mstTxts[step - 1].Enabled = false;
                        }
                        tmpTxts[step].Enabled = timTxts[step].Enabled = unitCmbs[step].Enabled = false;
                    }
                }
                catch (Exception ex)
                {
                    DataHelper.ErrorLog(ex.Message);
                }
            }
        }

        private bool CheckStatus()
        {
            if (Properties.Settings.Default.SimMode == true)
            {
                return true;
            }

            List<string> targetRegs = new List<string>();
            object[] readVals;
            short[] qlty;
            int[] errs;
            FILETIME[] ft;

            targetRegs.Add(string.Format("{0}.{1}{2}", DEV_NAME, WORD_REG_PREFIX, 1100));
            targetRegs.Add(string.Format("{0}.{1}{2}", DEV_NAME, BIT_REG_PREFIX, 901));
            targetRegs.Add(string.Format("{0}.{1}{2}", DEV_NAME, BIT_REG_PREFIX, 2050));
            targetRegs.Add(string.Format("{0}.{1}{2}", DEV_NAME, BIT_REG_PREFIX, 2051));

            if (opc.Read(targetRegs.ToArray(), out readVals, out qlty, out ft, out errs))
            {
                try {
                    int currentUser = Convert.ToInt32(readVals[0]);
                    if (currentUser != Properties.Settings.Default.OpcNo)
                    {
                        MessageBox.Show("操作権がありません。");
                        return false;
                    }
                    if (Convert.ToInt32(readVals[1]) == 0)
                    {
                        MessageBox.Show("操作モードが手元です。");
                        return false;
                    }
                    //if ( (Convert.ToInt32(readVals[2]) != 0) &&
                    //     (Convert.ToInt32(readVals[3]) == 0) )
                    if (Convert.ToInt32(readVals[2]) != 0)
                    {
                        MessageBox.Show("プログラム運転中です。");
                        return false;                        
                    }

                    return true;
                }
                catch (Exception ex)
                {
                    DataHelper.ErrorLog(ex.Message);
                }
            }
            else
            {
                // Log Error
                DataHelper.ErrorLog("Reading Failed in CheckStatus()");
            }

            return false;
        }

        private void ReadTimeBits()
        {
            List<string> targetRegs = new List<string>();
            object[] readVals;
            short[] qlty;
            int[] errs;
            FILETIME[] ft;

            foreach (ComboBox cb in unitCmbs)
            {
                int i = Convert.ToInt32(cb.Tag) + UNIT_BIT_POS + STEP_OFFSET;
                targetRegs.Add(string.Format("{0}.{1}{2}", DEV_NAME, BIT_REG_PREFIX, i));
                int j = Convert.ToInt32(cb.Tag) + UNIT_BIT_POS + STEP_OFFSET * 2;
                targetRegs.Add(string.Format("{0}.{1}{2}", DEV_NAME, BIT_REG_PREFIX, j));
            }

            if (opc.Read(targetRegs.ToArray(), out readVals, out qlty, out ft, out errs))
            {
                Debug.WriteLine("Reading Succeed in ReadValues()");

                int i = 0;
                foreach (ComboBox cb in unitCmbs)
                {
                    int ret1 = Convert.ToInt32(readVals[i]);
                    if (ret1 != 0)
                    {
                        cb.SelectedIndex = 1;
                    }
                    int ret2 = Convert.ToInt32(readVals[i + 1]);
                    if (ret2 != 0)
                    {
                        cb.SelectedIndex = 2;
                    }

                    i += 2;
                }
            }
            else
            {
                // Log Error
                DataHelper.ErrorLog("Reading Failed in ReadValues()");
            }
        }

        private void ReadValues(List<TextBox> tmps, int tmpPos, double minValue, double maxValue)
        {
            List<string> targetRegs = new List<string>();
            object[] readVals;
            short[] qlty;
            int[] errs;
            FILETIME[] ft;

            foreach (TextBox tb in tmps)
            {
                int i = Convert.ToInt32(tb.Tag) + tmpPos;
                targetRegs.Add(string.Format("{0}.{1}{2}", DEV_NAME, WORD_REG_PREFIX, i));
            }

            if (opc.Read(targetRegs.ToArray(), out readVals, out qlty, out ft, out errs))
            {
                Debug.WriteLine("Reading Succeed in ReadValues()");

                try
                {
                    int i = 0;
                    foreach (TextBox tb in tmps)
                    {
                        double rawValue = Convert.ToDouble(readVals[i]);
                        //入力値0-MAX_DIGITをminValue-maxValueへ変換
                        // y = EMin + ((ConvertedMax - ConvertedMin) * (x - RawMin)) / (RawMax-RawMin);
                        double newVal = minValue + ((maxValue - minValue) * (rawValue - 0)) / (MAX_DIGIT - 0);

                        tb.Text = newVal.ToString();
                        i++;
                    }
                }
                catch (Exception ex)
                {
                    DataHelper.ErrorLog(ex.Message);
                }
            }
            else
            {
                // Log Error
                DataHelper.ErrorLog("Reading Failed in ReadValues()");
            }

        }

        /// <summary>
        /// 時間値レジスタより取り出し
        /// </summary>
        private void ReadTimeValues()
        {
            List<string> targetRegs = new List<string>();
            object[] readVals;
            short[] qlty;
            int[] errs;
            FILETIME[] ft;

            foreach (TextBox tb in timTxts)
            {
                int i = Convert.ToInt32(tb.Tag) + LOWER_TIME_POS;
                targetRegs.Add(string.Format("{0}.{1}{2}", DEV_NAME, WORD_REG_PREFIX, i));
                i = Convert.ToInt32(tb.Tag) + UPPER_TIME_POS;
                targetRegs.Add(string.Format("{0}.{1}{2}", DEV_NAME, WORD_REG_PREFIX, i));
            }

            if (opc.Read(targetRegs.ToArray(), out readVals, out qlty, out ft, out errs))
            {
                Debug.WriteLine("Reading Succeed in ReadTimeValues()");

                int i = 0;
                try
                {
                    foreach (TextBox tb in timTxts)
                    {
                        int upperTime = Convert.ToInt32( readVals[i + 1] );
                        if (upperTime <= 0)
                        {
                            tb.Text = readVals[i].ToString();
                        }
                        else
                        {
                            tb.Text = string.Format( "{0}{1:0000}", upperTime, readVals[i] );
                        }
                        i += 2;
                    }
                }
                catch (Exception ex)
                {
                    DataHelper.ErrorLog(ex.Message);
                }
            }
            else
            {
                // Log Error
                DataHelper.ErrorLog("Reading Failed in ReadTimeValues()");
            }
        }

        /// <summary>
        /// フォームクローズ時にOPCサーバー接続を切断
        /// </summary>
        /// <param name="sender">メッセージ元</param>
        /// <param name="e">イベント</param>
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // opc.Disconnect();
        }

        /// <summary>
        /// 戻るボタンクリックでプログラム終了
        /// </summary>
        /// <param name="sender">メッセージ元</param>
        /// <param name="e">イベント</param>
        private void BackBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 全削除ボタンで値をクリア
        /// </summary>
        /// <param name="sender">メッセージ元</param>
        /// <param name="e">イベント</param>
        private void AllDelBtn_Click(object sender, EventArgs e)
        {
            ResetTable();
        }

        /// <summary>
        /// 設定ボタンで値をPLC書き込み
        /// </summary>
        /// <param name="sender">メッセージ元</param>
        /// <param name="e">イベント</param>
        private void SetBtn_Click(object sender, EventArgs e)
        {
            int[] errs;

            if (!CheckStatus())
            {
                return;
            }

            if ((Convert.ToInt32(timTxts[0].Text) <= 0) && timTxts[0].Enabled == true)
            {
                MessageBox.Show( "STEP 1の時間は0に出来ません。");
                return;
            }

            bool mstON = false;
            foreach (TextBox tb in mstTxts)
            {
                if (Convert.ToDouble(tb.Text) != 0)
                {
                    const double CTRL_TMP_MIN = 15.0;
                    const double CTRL_TMP_MAX = 60.0;
                    int nm = Convert.ToInt32(tb.Tag);
                    double ret;
                    double.TryParse(tmpTxts[nm - 1].Text, out ret);
                    if (ret < CTRL_TMP_MIN || CTRL_TMP_MAX < ret)
                    {
                        MessageBox.Show(string.Format(
                                "湿度設定がされているSTEP {0} の温度設定が {1} ～ {2} の範囲にありません。",
                                nm, CTRL_TMP_MIN, CTRL_TMP_MAX
                            ));
                        return;
                    }
                    else
                    {
                        mstON = true;
                    }
                }
            }

            bool tmpHigh = false;
            foreach (TextBox tb in tmpTxts) {
                const double CTR_BOIL_TMP = 15.0;                
                double ret;
                double.TryParse(tb.Text, out ret);
                if (CTR_BOIL_TMP < ret) {
                    tmpHigh = true;
                }
 
            }

            if ((mstON == true) || (tmpHigh == true))
            {
                List<string> targetRegs = new List<string>();
                object[] readVals;
                short[] qlty;
                FILETIME[] ft;

                targetRegs.Add(string.Format("{0}.{1}{2}", DEV_NAME, BIT_REG_PREFIX, 900));

                if (opc.Read(targetRegs.ToArray(), out readVals, out qlty, out ft, out errs))
                {
                    try
                    {
                        int boilerStatus = Convert.ToInt32(readVals[0]);
                        if (boilerStatus == 0)
                        {
                            MessageBox.Show("15℃以上設定の温度運転、及び湿度運転にはボイラー起動が必要です。その運転のステップ開始までにボイラー起動してください。");
                        }
                    }
                    catch (Exception ex)
                    {
                        DataHelper.ErrorLog(ex.Message);
                    }
                }
                else
                {
                    // Log Error
                    DataHelper.ErrorLog("Reading Failed in boiler check.");
                }
            }

            // 温度設定
            errs = WriteValues(TEM_POS, tmpTxts, TMP_LOW, TMP_HIGH);
            // 湿度設定
            errs = WriteValues(MST_POS, mstTxts, MST_LOW, MST_HIGH);
            // 時間
            errs = WriteTimeValues();
            // 繰返し回数
            errs = WriteRepeatValue();

            List<ComboBox> tmps = unitCmbs;
            errs = ClearBits(UNIT_BIT_POS, unitCmbs);
            errs = ClearBits(UNIT_BIT_POS + STEP_OFFSET, unitCmbs);
            errs = ClearBits(UNIT_BIT_POS + STEP_OFFSET * 2, unitCmbs);
            errs = ClearBits(UNIT_BIT_POS + STEP_OFFSET * 3, unitCmbs);     // 運転有効／無効フラグのクリア(時間単位初期化処理を流用)
            errs = SetTimeBits(UNIT_BIT_POS, STEP_OFFSET, unitCmbs);
            errs = SetEnableBits(ENABLE_BIT_POS, timTxts);                  // 運転有効／無効フラグのセット(更新時間配列を流用)
            // 2014/08/14 T.Yamanaka 繰り返し回数&ポーズ追加 開始
            errs = SetPauseBits(PAUSE_BIT_POS, pauseCmbs);
            // 2014/08/14 T.Yamanaka 繰り返し回数&ポーズ追加 終了

            WriteSecTimeValues();

            this.Close();
        }

        private int[] WriteRepeatValue()
        {
            int[] errs;
            string[] writeReg = new string[] { DEV_NAME+"."+WORD_REG_PREFIX+REP_POS };
            object[] writeVal = new object[] { numericUpDown.Value };

            if (opc.Write(writeReg, writeVal, out errs))
            {
                Debug.WriteLine("Set Writing Succeed in WriteTimeValues()");
            }
            else
            {
                // Log Error
                DataHelper.ErrorLog("Set Writing Failed in WriteTimeValues()");
            }
            return errs;
        }

        /// <summary>
        /// プログラム運転ステップ毎秒をレジスタ書き込み
        /// </summary>
        /// <returns>エラー状態配列</returns>
        private int[] WriteSecTimeValues()
        {
            List<string> targetRegs = new List<string>();
            List<object> writeVals = new List<object>();
            int[] errs;
            int[] timeTypes = { 1, 60, 3600 };  // 秒、分、時間

            for (int i = 0; i < timTxts.Count; i++ )
            {
                TextBox tb = timTxts[i];
                int wk = Convert.ToInt32(tb.Tag) + SEC_POS;
                targetRegs.Add(string.Format("{0}.{1}{2}", DEV_NAME, WORD_REG_PREFIX, wk));

                ComboBox cb = unitCmbs[i];
                int val = Convert.ToInt32(tb.Text) * timeTypes[ cb.SelectedIndex ];
                writeVals.Add(val);
            }

            if (opc.Write(targetRegs.ToArray(), writeVals.ToArray(), out errs))
            {
                Debug.WriteLine("Setting Writing Succeed in WriteSecTimeValues()");
            }
            else
            {
                // Log Error
                DataHelper.ErrorLog("Set Writing Failed in WriteSecTimeValues()");
            }
            return errs;
        }

        /// <summary>
        /// プログラム運転ステップ情報ビットをレジスタ書き込み
        /// </summary>
        /// <param name="bitPos">書き込み先レジスタ起点</param>
        /// <param name="tmps">書き込み元コンボボックス配列</param>
        /// <returns>エラー状態配列</returns>
        private int[] SetEnableBits(int bitPos, List<TextBox> tmps)
        {
            List<string> targetRegs = new List<string>();
            List<object> writeVals = new List<object>();
            int[] errs;

            foreach (TextBox tb in tmps)
            {
                int i = Convert.ToInt32(tb.Tag) + bitPos;
                targetRegs.Add(string.Format("{0}.{1}{2}", DEV_NAME, BIT_REG_PREFIX, i));
                int val = Convert.ToInt32(tb.Text);
                if (val <= 0)
                {
                    writeVals.Add(0);
                }
                else
                {
                    writeVals.Add(1);
                }
            }

            if (opc.Write(targetRegs.ToArray(), writeVals.ToArray(), out errs))
            {
                Debug.WriteLine("Bit Set Writing Succeed in SetEnableBits()");
            }
            else
            {
                // Log Error
                DataHelper.ErrorLog("Bit Set Writing Failed in SetEnableBits()");
            }
            return errs;
        }

        private int[] SetPauseBits(int bitPos, List<ComboBox> tmps)
        {
            List<string> targetRegs = new List<string>();
            List<object> writeVals = new List<object>();
            int[] errs;

            foreach (ComboBox tb in tmps)
            {
                int i = Convert.ToInt32(tb.Tag) + bitPos;
                targetRegs.Add(string.Format("{0}.{1}{2}", DEV_NAME, BIT_REG_PREFIX, i));
                int val = tb.SelectedIndex;
                if (val <= 0)
                {
                    writeVals.Add(0);
                }
                else
                {
                    writeVals.Add(1);
                }
            }

            if (opc.Write(targetRegs.ToArray(), writeVals.ToArray(), out errs))
            {
                Debug.WriteLine("Bit Set Writing Succeed in SetEnableBits()");
            }
            else
            {
                // Log Error
                DataHelper.ErrorLog("Bit Set Writing Failed in SetEnableBits()");
            }
            return errs;
        }

        /// <summary>
        /// 時間単位ビットのセット処理
        /// </summary>
        /// <param name="bitPos">書き込み先レジスタ起点</param>
        /// <param name="offset">レジスタオフセット</param>
        /// <param name="tmps">書き込み用コンボボックス配列</param>
        /// <returns>エラー状態配列</returns>
        private int[] SetTimeBits(int bitPos, int offset, List<ComboBox> tmps)
        {
            List<string> targetRegs = new List<string>();
            List<object> writeVals = new List<object>();
            int[] errs;

            foreach (ComboBox cb in tmps)
            {
                int i = Convert.ToInt32(cb.Tag) + bitPos + offset * cb.SelectedIndex;
                targetRegs.Add(string.Format("{0}.{1}{2}", DEV_NAME, BIT_REG_PREFIX, i));
                writeVals.Add(1);
            }

            if (opc.Write(targetRegs.ToArray(), writeVals.ToArray(), out errs))
            {
                Debug.WriteLine("Bit Set Writing Succeed in SetTimeBits()");
            }
            else
            {
                // Log Error
                DataHelper.ErrorLog("Bit Set Writing Failed in SetTimeBits()");
            }
            return errs;
        }

        /// <summary>
        /// 時間単位ビットのクリア処理
        /// </summary>
        /// <param name="bitPos">書き込み先レジスタ起点</param>
        /// <param name="tmps">書き込み用コンボボックス配列</param>
        /// <returns>エラー状態配列</returns>
        private int[] ClearBits(int bitPos, List<ComboBox> tmps)
        {
            List<string> targetRegs = new List<string>();
            List<object> writeVals = new List<object>();
            int[] errs;

            foreach (ComboBox cb in tmps)
            {
                int i = Convert.ToInt32(cb.Tag) + bitPos;
                targetRegs.Add(string.Format("{0}.{1}{2}", DEV_NAME, BIT_REG_PREFIX, i));
                writeVals.Add(0);
            }

            if (opc.Write(targetRegs.ToArray(), writeVals.ToArray(), out errs))
            {
                Debug.WriteLine("Bit Clear Writing Succeed in ClearBits()");
            }
            else
            {
                // Log Error
                DataHelper.ErrorLog("Bit Clear Writing Failed in ClearBits()");
            }
            return errs;
        }

        /// <summary>
        /// DCS表示用の時間をレジスタ書き込み
        /// </summary>
        /// <returns>エラー状態配列</returns>
        private int[] WriteTimeValues()
        {
            List<string> targetRegs = new List<string>();
            List<object> writeVals = new List<object>();
            int[] errs;

            foreach (TextBox tb in timTxts)
            {
                int i = Convert.ToInt32(tb.Tag) + LOWER_TIME_POS;
                targetRegs.Add(string.Format("{0}.{1}{2}", DEV_NAME, WORD_REG_PREFIX, i));
                writeVals.Add(Convert.ToInt32(tb.Text) % MAX_DIGIT);
                i = Convert.ToInt32(tb.Tag) + UPPER_TIME_POS;
                targetRegs.Add(string.Format("{0}.{1}{2}", DEV_NAME, WORD_REG_PREFIX, i));
                writeVals.Add(Convert.ToInt32(tb.Text) / MAX_DIGIT);
            }

            if (opc.Write(targetRegs.ToArray(), writeVals.ToArray(), out errs))
            {
                Debug.WriteLine("Set Writing Succeed in WriteTimeValues()");
            }
            else
            {
                // Log Error
                DataHelper.ErrorLog("Set Writing Failed in WriteTimeValues()");
            }
            return errs;
        }

        /// <summary>
        /// DCS用に値を換算してレジスタ書き込み
        /// </summary>
        /// <param name="tmpPos">書き込みレジスタ起点</param>
        /// <param name="tmps">書き込み用テキストボックス配列</param>
        /// <param name="rangeLow">変換前下限値</param>
        /// <param name="rangeHigh">変換前上限値</param>
        /// <returns></returns>
        private int[] WriteValues(int tmpPos, List<TextBox> tmps, double rangeLow, double rangeHigh)
        {
            List<string> targetRegs = new List<string>();
            List<object> writeVals = new List<object>();
            int[] errs;

            foreach (TextBox tb in tmps)
            {
                int i = Convert.ToInt32(tb.Tag) + tmpPos;
                targetRegs.Add(string.Format("{0}.{1}{2}", DEV_NAME, WORD_REG_PREFIX, i));
                //入力値を0-10000digitへ変換
                // y = EMin + ((ConvertedMax - ConvertedMin) * (x - RawMin)) / (RawMax-RawMin);
                double rawDigit = 0 + ((MAX_DIGIT - 0) * (Convert.ToDouble(tb.Text) - rangeLow)) / (rangeHigh - rangeLow);
                writeVals.Add(Convert.ToInt32(rawDigit));
            }

            if (opc.Write(targetRegs.ToArray(), writeVals.ToArray(), out errs))
            {
                Debug.WriteLine("Set Writing Succeed in WriteValues()");
            }
            else
            {
                // Log Error
                DataHelper.ErrorLog("Set Writing Failed in WriteValues()");
            }
            return errs;
        }

        /// <summary>
        /// デフォーカス時のクローズ用イベントハンドラー
        /// </summary>
        /// <param name="sender">メッセージ元</param>
        /// <param name="e">イベント</param>
        private void Form1_Deactivate(object sender, EventArgs e)
        {
            //this.Close();
        }

        /// <summary>
        /// テスト用現在レジスタ値読み出し用メソッド
        /// </summary>
        /// <param name="sender">メッセージ元</param>
        /// <param name="e">イベント</param>
        private void button1_Click(object sender, EventArgs e)
        {
            ReadValues(tmpTxts, TEM_POS, TMP_LOW, TMP_HIGH);
            ReadValues(mstTxts, MST_POS, MST_LOW, MST_HIGH);
            ReadTimeBits();
            ReadTimeValues();
        }

        private void PatternReadBtn_Click(object sender, EventArgs e)
        {
            openFileDialog.FileName = PatternTxt.Text;
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                PatternTxt.Text = Path.GetFileNameWithoutExtension(openFileDialog.FileName);

                string[] lines = File.ReadAllLines(openFileDialog.FileName);
                if (lines.Length == (tmpTxts.Count + mstTxts.Count + timTxts.Count
                                      + unitCmbs.Count + pauseCmbs.Count + 1))
                {
                    int count = 0;
                    numericUpDown.Value = Convert.ToInt32(lines[count++]);

                    foreach (var tmp in tmpTxts)
                    {
                        tmp.Text = lines[count++];
                    }
                    foreach (var mst in mstTxts)
                    {
                        mst.Text = lines[count++];
                    }
                    foreach (var tim in timTxts)
                    {
                        tim.Text = lines[count++];
                    }
                    foreach (var uni in unitCmbs)
                    {
                        uni.SelectedIndex = Convert.ToInt32(lines[count++]);
                    }
                    foreach (var pos in pauseCmbs)
                    {
                        pos.SelectedIndex = Convert.ToInt32(lines[count++]);
                    }
                }
                else
                {
                    MessageBox.Show("読み込まれたパターン・ファイルに誤りがあります。",
                        "パターンファイル読み込みエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void PatternWriteBtn_Click(object sender, EventArgs e)
        {
            saveFileDialog.FileName = PatternTxt.Text;
            if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(numericUpDown.Value.ToString());
                foreach (var tmp in tmpTxts)
                {
                    sb.AppendLine(tmp.Text);
                }
                foreach (var mst in mstTxts)
                {
                    sb.AppendLine(mst.Text);
                }
                foreach (var tim in timTxts)
                {
                    sb.AppendLine(tim.Text);
                }
                foreach (var uni in unitCmbs)
                {
                    sb.AppendLine(uni.SelectedIndex.ToString());
                }
                foreach (var pos in pauseCmbs)
                {
                    sb.AppendLine(pos.SelectedIndex.ToString());
                }
                File.WriteAllText(saveFileDialog.FileName, sb.ToString());
            }
        }

    }
}

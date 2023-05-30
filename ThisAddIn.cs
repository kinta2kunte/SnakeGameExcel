using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace SnakeGameExcel
{
    public partial class ThisAddIn
    {
        public Excel.Application oExcel = null;
        private Excel._Worksheet objSheet = null;

        // 壁マップ
        private int[,] wallPos = new int[,] {
            // 1:無描画 2:縦 3:横 4:2重縦 未実装=>5:2重横<= 6:左上角丸 7:左下角丸 8:右上角丸 9:右下角丸
            //0---------+---------*---------+---------*---------+---------*---  
            {6,3,3,3,3,3,3,3,3,3,3,3,3,3,3,8,6,3,3,3,3,3,3,3,3,3,3,3,3,3,3,8},
            {2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,2,2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,2},
            {2,0,6,3,3,3,8,0,6,3,3,3,3,8,0,2,2,0,6,3,3,3,3,8,0,6,3,3,3,8,0,2},
            {2,0,2,1,1,1,2,0,2,1,1,1,1,2,0,2,2,0,2,1,1,1,1,2,0,2,1,1,1,2,0,2},
            {2,0,2,1,1,1,2,0,2,1,1,1,1,2,0,2,2,0,2,1,1,1,1,2,0,2,1,1,1,2,0,2},
            {2,0,7,3,3,3,9,0,7,3,3,3,3,9,0,7,9,0,7,3,3,3,3,9,0,7,3,3,3,9,0,2},
            {2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,2},
            {2,0,6,3,3,3,8,0,6,8,0,6,3,3,3,3,3,3,3,3,8,0,6,8,0,6,3,3,3,8,0,2},
            {2,0,7,3,3,3,9,0,2,2,0,7,3,3,3,8,6,3,3,3,9,0,2,2,0,7,3,3,3,9,0,2},
            {2,0,0,0,0,0,0,0,2,2,0,0,0,0,0,2,2,0,0,0,0,0,2,2,0,0,0,0,0,0,0,2},
            {7,3,3,3,3,3,8,0,2,7,3,3,3,8,0,2,2,0,6,3,3,3,9,2,0,6,3,3,3,3,3,9},
            {1,1,1,1,1,1,2,0,2,6,3,3,3,9,0,7,9,0,7,3,3,3,8,2,0,2,1,1,1,1,1,1},
            {1,1,1,1,1,1,2,0,2,2,0,0,0,0,0,0,0,0,0,0,0,0,2,2,0,2,1,1,1,1,1,1},
            {1,1,1,1,1,1,2,0,2,2,0,6,3,3,3,3,3,3,3,3,8,0,2,2,0,2,1,1,1,1,1,1},
            {3,3,3,3,3,3,9,0,7,9,0,2,1,1,1,1,1,1,1,1,2,0,7,9,0,7,3,3,3,3,3,3},
            {0,0,0,0,0,0,0,0,0,0,0,2,1,1,1,1,1,1,1,1,2,0,0,0,0,0,0,0,0,0,0,0},
            {3,3,3,3,3,3,8,0,6,8,0,2,1,1,1,1,1,1,1,1,2,0,6,8,0,6,3,3,3,3,3,3},
            {1,1,1,1,1,1,2,0,2,2,0,7,3,3,3,3,3,3,3,3,9,0,2,2,0,2,1,1,1,1,1,1},
            {1,1,1,1,1,1,2,0,2,2,0,0,0,0,0,0,0,0,0,0,0,0,2,2,0,2,1,1,1,1,1,1},
            {1,1,1,1,1,1,2,0,2,2,0,6,3,3,3,3,3,3,3,3,8,0,2,2,0,2,1,1,1,1,1,1},
            {6,3,3,3,3,3,9,0,7,9,0,7,3,3,3,8,6,3,3,3,9,0,7,9,0,7,3,3,3,3,3,8},
            {2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,2,2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,2},
            {2,0,6,3,3,3,8,0,6,3,3,3,3,8,0,2,2,0,6,3,3,3,3,8,0,6,3,3,3,8,0,2},
            {2,0,7,3,3,8,2,0,7,3,3,3,3,9,0,7,9,0,7,3,3,3,3,9,0,2,6,3,3,9,0,2},
            {2,0,0,0,0,2,2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,2,2,0,0,0,0,2},
            {7,3,3,8,0,2,2,0,6,8,0,6,3,3,3,3,3,3,3,3,8,0,6,8,0,2,2,0,6,3,3,9},
            {6,3,3,9,0,7,9,0,2,2,0,7,3,3,3,8,6,3,3,3,9,0,2,2,0,7,9,0,7,3,3,8},
            {2,0,0,0,0,0,0,0,2,2,0,0,0,0,0,2,2,0,0,0,0,0,2,2,0,0,0,0,0,0,0,2},
            {2,0,6,3,3,3,3,3,9,7,3,3,3,8,0,2,2,0,6,3,3,3,9,7,3,3,3,3,3,8,0,2},
            {2,0,7,3,3,3,3,3,3,3,3,3,3,9,0,7,9,0,7,3,3,3,3,3,3,3,3,3,3,9,0,2},
            {2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,2},
            {7,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,9}
        };

        private int _width = 32;
        private int _height = 32;

        private bool bInit = false;
        private bool bStartFlg = false;

        private int nScore = 0;                  // スコア
        private int nSpeed = 400;                // 進行速度
        //int nAniFlg = 0;                      // アニメーション
        public int nNextDirectionMove = 1;      // 先行入力用移動方向
        private int nDirectionMove = 0;          // 移動方向

        // 効果音再生用
        private System.Media.SoundPlayer _player = new System.Media.SoundPlayer("c:\\temp\\se_get_1.wav");

        // タイマー
        System.Windows.Forms.Timer _timer = null;

        /// <summary>
        /// 座標用クラス
        /// </summary>
        class snakePos
        {
            public int x { get; set; }
            public int y { get; set; }
            public snakePos(int _x, int _y)
            {
                x = _x;
                y = _y;
            }
        }        // スネーク座標
        List<snakePos> pSnakePos = new List<snakePos> {
            new snakePos( 15, 24 ),
            new snakePos( 16, 24 ),
            new snakePos( 17, 24 ),
            new snakePos( 18, 24 ) };
        // 餌座標
        snakePos pFoodPos = new snakePos(1, 1);

        // キーフック関連処理
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod,
            uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        public delegate int LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private static LowLevelKeyboardProc _proc = HookCallback;
        private static IntPtr _hookID = IntPtr.Zero;

        //declare the mouse hook constant.
        //For other hook types, you can obtain these values from Winuser.h in the Microsoft SDK.

        private const int WH_KEYBOARD = 2; // mouse
        private const int HC_ACTION = 0;

        private const int WH_KEYBOARD_LL = 13; // keyboard
        private const int WM_KEYDOWN = 0x0100;

        public static void SetHook()
        {
            // Ignore this compiler warning, as SetWindowsHookEx doesn't work with ManagedThreadId
#pragma warning disable 618
            _hookID = SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, (uint)AppDomain.GetCurrentThreadId());
#pragma warning restore 618

        }

        public static void ReleaseHook()
        {
            UnhookWindowsHookEx(_hookID);
        }

        //Note that the custom code goes in this method the rest of the class stays the same.
        //It will trap if BOTH keys are pressed down.
        private static int HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0 )
            {
                return (int)CallNextHookEx(_hookID, nCode, wParam, lParam);
            }
            else
            {

                //if (nCode == HC_ACTION)
                //{
                //    Keys keyData = (Keys)wParam;
                //}
                Keys keyData = (Keys)wParam;
            }
            return (int)CallNextHookEx(_hookID, nCode, wParam, lParam);
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            oExcel = (Excel.Application)Globals.ThisAddIn.Application;
            if (oExcel == null)
            {
                MessageBox.Show("アプリケーション接続エラー", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // キーボードフック処理
            //SetHook();

            Globals.Ribbons.Ribbon1.btnStart.Enabled = false;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public void btnInit()
        {
            if( _timer != null )
            { 
                _timer.Stop();
                _timer.Dispose();
                _timer = null;
                nNextDirectionMove = 1;      // 先行入力用移動方向
                nDirectionMove = 0;          // 移動方向
                pSnakePos.Clear();
                pSnakePos = new List<snakePos> {
                    new snakePos( 15, 24 ),
                    new snakePos( 16, 24 ),
                    new snakePos( 17, 24 ),
                    new snakePos( 18, 24 ) };
                bInit = false;
                bStartFlg = false;

                nScore = 0;                  // スコア
            }

            //アクティブシート(の直前)にシートを追加
            if( objSheet != null )
                objSheet.Delete();

            //ws = oExcel.ActiveWorkbook.ActiveSheet; 
            oExcel.ActiveWorkbook.Worksheets.Add(Type.Missing, oExcel.ActiveWorkbook.ActiveSheet, 1);

            objSheet = oExcel.ActiveWorkbook.ActiveSheet;

            objSheet.Name = "スネークゲームTR+";

            objSheet.Columns.ColumnWidth = 1.5;
            objSheet.Columns.RowHeight = 14;

            //範囲指定
            Excel.Range range_color;
            range_color = objSheet.Range[objSheet.Cells[1, 1], objSheet.Cells[32, 32]];
            //背景色変更
            range_color.Interior.Color = Color.FromArgb(0, 0, 0);

            drawMap();

            drawMe();

            drawScore();

            makeFood();
            drawFood();

            Globals.Ribbons.Ribbon1.btnStart.Enabled = true;
            bInit = true;

        }
        private void drawMap()
        {
            for (int x = 0; x < _width; x++)
            {
                for (int y = 0; y < _height; y++)
                {
                    if(wallPos[y, x] > 1)
                    {
                        Excel.Range range_color;
                        range_color = objSheet.Range[objSheet.Cells[y + 1, x + 1], objSheet.Cells[y+1, x+1]];
                        //背景色変更
                        range_color.Interior.Color = Color.FromArgb(0, 0, 255);

                    }
                    else
                    {
                        Excel.Range range_color;
                        range_color = objSheet.Range[objSheet.Cells[y + 1, x + 1], objSheet.Cells[y + 1, x + 1]];
                        //背景色変更
                        range_color.Interior.Color = Color.FromArgb(0, 0, 0);
                        range_color.Value2 = "";
                    }
                }
            }

        }
        private void drawScore()
        {
            Excel.Range rng = objSheet.Range[objSheet.Cells[1, _width + 1], objSheet.Cells[1, _width + 1]];
            rng.Value = "Score:";
            Excel.Range rng2 = objSheet.Range[objSheet.Cells[1, _width + 8], objSheet.Cells[1, _width + 8]];
            rng2.Value = nScore.ToString();
        }
        private void drawMe()
        {
            Excel.Range rng = objSheet.Range[objSheet.Cells[pSnakePos[0].y + 1, pSnakePos[0].x + 1], objSheet.Cells[pSnakePos[0].y + 1, pSnakePos[0].x + 1]];
            rng.Value = "●";
            rng.Interior.Color = Color.FromArgb(0, 0, 0);
            rng.Font.Color = Color.FromArgb(255, 255, 0);
            for (int i = 1; i < pSnakePos.Count; i++)
            {
                Excel.Range rng2 = objSheet.Range[objSheet.Cells[pSnakePos[i].y + 1, pSnakePos[i].x + 1], objSheet.Cells[pSnakePos[i].y + 1, pSnakePos[i].x + 1]];
                rng2.Value = "■";
                rng2.Interior.Color = Color.FromArgb(0, 0, 0);
                rng2.Font.Color = Color.FromArgb(255, 128, 128);
            }
        }
        private void eraseMe()
        {
            // 先頭を尻尾にする
            Excel.Range rng = objSheet.Range[objSheet.Cells[pSnakePos[0].y + 1, pSnakePos[0].x + 1], objSheet.Cells[pSnakePos[0].y + 1, pSnakePos[0].x + 1]];
            rng.Value = "■";
            rng.Interior.Color = Color.FromArgb(0, 0, 0);
            rng.Font.Color = Color.FromArgb(255, 128, 128);
            // 最後尾を消す
            Excel.Range rng2 = objSheet.Range[objSheet.Cells[pSnakePos[pSnakePos.Count-1].y + 1, pSnakePos[pSnakePos.Count-1].x + 1], objSheet.Cells[pSnakePos[pSnakePos.Count-1].y + 1, pSnakePos[pSnakePos.Count-1].x + 1]];
            rng2.Value = "■";
            rng2.Interior.Color = Color.FromArgb(0, 0, 0);
            rng2.Font.Color = Color.FromArgb(0, 0, 0);
        }

        public void btnStart()
        {
            if (bStartFlg == false)
            {
                // タイマーの間隔(ミリ秒)
                _timer = new System.Windows.Forms.Timer();
                _timer.Tick += new EventHandler(tickHandler);
                _timer.Interval = nSpeed;
                _timer.Start();
                bStartFlg = true;
                //btnStart.Text = "停止";

                // 開始初期値＝左方向
                nNextDirectionMove = 1;
                // 餌座標演算
                //makeFood();
                Globals.Ribbons.Ribbon1.btnStart.Enabled = false;
            }
            else
            {
                _timer.Stop();
                _timer.Dispose();
                _timer = null;
                bStartFlg = false;
                //btnStart.Text = "開始";
            }
        }
        /// <summary>
        /// タイマー割り込みイベント
        /// メイン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tickHandler(object sender, EventArgs e)
        {
            try
            {
                snakePos pos = pSnakePos[0];
                int mx = pos.x;
                int my = pos.y;

                //先行入力処理
                switch (nNextDirectionMove)
                {
                    case 1:                                 // 左方向
                        if (wallPos[pos.y, pos.x - 1] == 0)         // 道判定（壁判定）
                        {
                            nNextDirectionMove = 0;
                            nDirectionMove = 1;
                        }
                        break;
                    case 2:                                 // 右方向
                        if (wallPos[pos.y, pos.x + 1] == 0)        // 道判定（壁判定）
                        {
                            nNextDirectionMove = 0;
                            nDirectionMove = 2;
                        }
                        break;
                    case 3:                                 // 上方向
                        if (wallPos[pos.y - 1, pos.x] == 0)        // 道判定（壁判定）
                        {
                            nNextDirectionMove = 0;
                            nDirectionMove = 3;
                        }
                        break;
                    case 4:                                 // 下方向
                        if (wallPos[pos.y + 1, pos.x] == 0)        // 道判定（壁判定）
                        {
                            nNextDirectionMove = 0;
                            nDirectionMove = 4;
                        }
                        break;

                }
                // 移動
                switch (nDirectionMove)
                {
                    case 1:
                        if (pos.x <= 0)                        // ワープ判定
                            mx = _width - 1;
                        else if (wallPos[pos.y, pos.x - 1] == 0)
                            mx -= 1;
                        break;
                    case 2:
                        if (pos.x >= _width - 1)                        // ワープ判定
                            mx = 0;
                        else if (wallPos[pos.y, pos.x + 1] == 0)
                            mx += 1;
                        break;
                    case 3:
                        if (wallPos[pos.y - 1, pos.x] == 0)
                            my -= 1;
                        break;
                    case 4:
                        if (wallPos[pos.y + 1, pos.x] == 0)
                            my += 1;
                        break;
                }
                // 移動処理
                if (mx != pos.x || my != pos.y)
                {
                    eraseMe();
                    snakePos p = new snakePos(mx, my);
                    pSnakePos.Insert(0, p);
                    pSnakePos.RemoveAt(pSnakePos.Count - 1);
                    //pictureBox1.Invalidate();        //PictureBox更新
                    drawMe();
                }

                // 当たり判定
                if (bingo() == true)
                {
                    if (bStartFlg == true)
                    {
                        _timer.Stop();
                        _timer.Dispose();
                        _timer = null;
                        bStartFlg = false;
                        // ゲームオーバー
                        MessageBox.Show("ゲームオーバー", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

            }
            catch (Exception ex)
            {
                // 例外処理
                return;
            }
        }
        /// <summary>
        /// 食べ物位置計算
        /// </summary>
        private void makeFood()
        {
            List<snakePos> fPos = new List<snakePos>();

            // 位置演算
            for (int x = 0; x < _width; x++)
            {
                for (int y = 0; y < _height; y++)
                {
                    if (wallPos[y, x] == 0)
                    {
                        // スネーク自分判定
                        int myp = 0;
                        for (int i = 0; i < pSnakePos.Count - 1; i++)
                        {
                            if (pSnakePos[i].x == x && pSnakePos[i].y == y)
                                myp++;
                        }
                        if (myp == 0)
                        {
                            // 道（map=0）を抽出
                            snakePos p = new snakePos(x, y);
                            fPos.Add(p);
                        }
                    }
                }
            }
            // 餌位置を道（map=0）からランダムに選択
            var rand = new Random();
            int num = rand.Next(minValue: 0, maxValue: fPos.Count - 1);
            snakePos pos = fPos[num];
            pFoodPos.x = pos.x;
            pFoodPos.y = pos.y;
        }
        /// <summary>
        /// 餌描画
        /// </summary>
        private void drawFood()
        {
            Excel.Range rng = objSheet.Range[objSheet.Cells[pFoodPos.y + 1, pFoodPos.x + 1], objSheet.Cells[pFoodPos.y + 1, pFoodPos.x + 1]];
            rng.Value = "＊";
            rng.Interior.Color = Color.FromArgb(0, 0, 0);
            rng.Font.Color = Color.FromArgb(255, 0, 0);
            //g.FillEllipse(Brushes.Red, pFoodPos.x * _size + (_size / 3), pFoodPos.y * _size + (_size / 3), _size / 2, _size / 2);
        }

        /// <summary>
        /// 当たり判定
        /// </summary>
        private bool bingo()
        {
            snakePos myPos = pSnakePos[0];

            // 餌・当たり判定
            if (pFoodPos.x == myPos.x && pFoodPos.y == myPos.y)
            {
                snakePos p = pSnakePos[pSnakePos.Count - 1];
                pSnakePos.Add(p);

                // スコア加算
                nScore++;
                //lblScore.Text = nScore.ToString();
                // ゲームスピード調整
                nSpeed -= 5;
                _timer.Interval = nSpeed;
                //document.getElementById("gameScore").textContent = " Score: " + score + "  Speed:" + sp;

                // 効果音
                //_player.Play();

                // 新たな餌演算
                makeFood();
                drawFood();
                drawScore();
            }

            // スネーク・当たり判定
            for (int i = 1; i < pSnakePos.Count; i++)
            {
                snakePos p = pSnakePos[i];
                if (p.x == myPos.x && p.y == myPos.y)
                {
                    return true;
                }
            }
            return false;
        }

        public void btnHelp()
        {
            help hp = new help();
            hp.Show();
        }
        public void btnSettings()
        {
            frmSettings st = new frmSettings();
            if( st.ShowDialog() == DialogResult.OK )
            {
                switch(st.gameMode)
                {
                    case 0:
                        nSpeed = 400;
                        break;
                    case 1:
                        nSpeed = 300;
                        break;
                    case 2:
                        nSpeed = 200;
                        break;
                    case 3:
                        nSpeed = 100;
                        break;

                }
                if( _timer != null ) _timer.Interval = nSpeed;
            }
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}

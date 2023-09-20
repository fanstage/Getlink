using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Net;
using System.Net.Mail;
using System.Text;
namespace Getlink
{
    public partial class Form1 : Form
    {
        private SQLiteConnection sqlite;
        private int currentId = 1;
        [DllImport("user32.dll")]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);

        private const uint MOUSEEVENTF_LEFTDOWN = 0x02;
        private const uint MOUSEEVENTF_LEFTUP = 0x04;
        // 模拟鼠标左键点击事件

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;        // x position of upper-left corner
            public int Top;         // y position of upper-left corner
            public int Right;       // x position of lower-right corner
            public int Bottom;      // y position of lower-right corner
        }
        public static void FixWindowPositionAndSize(string windowTitle, int x, int y, int width, int height)
        {
            // 查找窗口
            IntPtr hWnd = FindWindow(null, windowTitle);
            if (hWnd == IntPtr.Zero)
            {
                Console.WriteLine("未找到指定标题的窗口");
                return;
            }

            // 将窗口移动到指定位置
            SetWindowPos(hWnd, IntPtr.Zero, x, y, 0, 0, 0x0040 | 0x0080);

            // 设置窗口大小
            SetWindowPos(hWnd, IntPtr.Zero, x, y, width, height, 0x0040 | 0x0080);
            SetForegroundWindow(hWnd);
        }
        //根据坐标点击位置
        public static void SendMail(string recipent,string subject,string body)
        { }

        public Form1()
        {
            InitializeComponent();
            if (!File.Exists("./Links.db"))
            {
                SQLiteConnection.CreateFile("Links.db");
            }
            sqlite = new SQLiteConnection("Data Source=Links.db;");
            sqlite.Open();
        }
        private void Click_1(int x, int y)
        {
            SetCursorPos(157, 450);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            Thread.Sleep(1000);
            SetCursorPos(x, y);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            Thread.Sleep(1000);
            SetCursorPos(608, 103);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            string clipboardText = Clipboard.GetText();
            textBox1.Text += clipboardText;
            textBox1.Text += "\n";
            /*File.AppendAllText("link.txt", Clipboard.GetText());
            File.AppendAllText("link.txt", "\n");*/
            string sql = $"SELECT Link FROM LinksTable WHERE Id = {currentId};";
            SQLiteCommand command = new SQLiteCommand(sql, sqlite);
            object result = command.ExecuteScalar();

            if (result != null)
            {
                string existingLink = result.ToString();
                if (!existingLink.Equals(clipboardText))
                {
                    sql = $"UPDATE LinksTable SET Link = '{clipboardText}' WHERE Id = {currentId};";
                    command = new SQLiteCommand(sql, sqlite);
                    command.ExecuteNonQuery();
                }
            }

            currentId++;
            Thread.Sleep(1000);
            SetCursorPos(1163, 70);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            Thread.Sleep(1000);
        }
        private void Click_2(int x, int y)
        {
            SetCursorPos(277, 450);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            Thread.Sleep(1000);
            SetCursorPos(x, y);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            Thread.Sleep(1000);
            SetCursorPos(608, 103);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            string clipboardText = Clipboard.GetText();
            textBox1.Text += clipboardText;
            textBox1.Text += "\n";
            /*File.AppendAllText("link.txt", Clipboard.GetText());
            File.AppendAllText("link.txt", "\n");*/
            string sql = $"SELECT Link FROM LinksTable WHERE Id = {currentId};";
            SQLiteCommand command = new SQLiteCommand(sql, sqlite);
            object result = command.ExecuteScalar();

            if (result != null)
            {
                string existingLink = result.ToString();
                if (!existingLink.Equals(clipboardText))
                {
                    sql = $"UPDATE LinksTable SET Link = '{clipboardText}' WHERE Id = {currentId};";
                    command = new SQLiteCommand(sql, sqlite);
                    command.ExecuteNonQuery();
                }
            }

            currentId++;
            Thread.Sleep(1000);
            SetCursorPos(1163, 70);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            Thread.Sleep(1000);
        }
        private void Click_3(int x, int y)
        {
            SetCursorPos(396, 450);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            Thread.Sleep(1000);
            SetCursorPos(x, y);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            Thread.Sleep(1000);
            SetCursorPos(608, 103);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            string clipboardText = Clipboard.GetText();
            textBox1.Text += clipboardText;
            textBox1.Text += "\n";
            /*File.AppendAllText("link.txt", Clipboard.GetText());
            File.AppendAllText("link.txt", "\n");*/
            string sql = $"SELECT Link FROM LinksTable WHERE Id = {currentId};";
            SQLiteCommand command = new SQLiteCommand(sql, sqlite);
            object result = command.ExecuteScalar();

            if (result != null)
            {
                string existingLink = result.ToString();
                if (!existingLink.Equals(clipboardText))
                {
                    sql = $"UPDATE LinksTable SET Link = '{clipboardText}' WHERE Id = {currentId};";
                    command = new SQLiteCommand(sql, sqlite);
                    command.ExecuteNonQuery();
                }
            }

            currentId++;
            Thread.Sleep(1000);
            SetCursorPos(1163, 70);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            Thread.Sleep(1000);
        }
        private void Start_Click(object sender, EventArgs e)
        {
            FixWindowPositionAndSize("上海海事大学", 100, 100, 400, 374);
            Thread.Sleep(1000);
            //SimulateClick(); 
            Click_1(154, 408);
            Click_1(154, 375);
            //Click_1(154, 342);此处为小程序
            Click_1(154, 310);
            Click_1(154, 275);
            Click_2(277, 407);
            Click_2(277, 375);
            Click_2(277, 342);
            Click_2(277, 310);
            Click_2(277, 275);
            Click_3(396, 407);
            Click_3(396, 375);
            Click_3(396, 342);
            Click_3(396, 310);
            Click_3(396, 275);
        }
    }
}
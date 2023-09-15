using System.Runtime.InteropServices;
namespace Getlink
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        [StructLayout(LayoutKind.Sequential)]
        public struct Point
        {
            public int X;
            public int Y;
        }
        public static void SimulateClick(Control control,Point clickCoordinates)
        {
            control.MousePosition = clickCoordinates; // �������λ��Ϊָ������
            Control.SendMouseClick(Control.ActiveForm, 0, 0, 0); // ģ�������
        }
        public static void FixWindowPositionAndSize(string windowTitle, int x, int y, int width, int height)
        {
            // ���Ҵ���
            IntPtr hWnd = FindWindow(null, windowTitle);
            if (hWnd == IntPtr.Zero)
            {
                Console.WriteLine("δ�ҵ�ָ������Ĵ���");
                return;
            }

            // �������ƶ���ָ��λ��
            SetWindowPos(hWnd, IntPtr.Zero, x, y, 0, 0, 0x0040 | 0x0080);

            // ���ô��ڴ�С
            SetWindowPos(hWnd, IntPtr.Zero, x, y, width, height, 0x0040 | 0x0080);
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void Start_Click(object sender, EventArgs e)
        {
            FixWindowPositionAndSize("�Ϻ����´�ѧ", 100, 100, 400, 374);
        }
    }
}
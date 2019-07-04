using System;
using System.Collections.Generic;
using System.Windows.Forms;
using UsbHid;
using UsbHid.USB.Classes.Messaging;

namespace USB2550HidTest.Forms
{
    public partial class FrmMain : Form
    {
        public UsbHidDevice Device;

        public byte[] READER_T_STARTSCAN = new byte[] { 0x07, 0x11, 0x00, 0x86, 0x00, 0x02, 0x00, 0x00, 0x00, 0x0D, 0x8C, 0x00, 0x05, 0x00, 0x00, 0x01, 0x01, 0x00, 0x01,0x06 };
        public byte[] READER_T_STOPSCAN = new byte[] { 0x08, 0x0A, 0x00, 0x8C, 0x00, 0x05, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00 };
        public byte[] READER_T_TEST = new byte[] { 0x07, 0x04, 0x03, 0x03, 0x01, 0x04 };
        //'Select dedicate EPC 81 1A 00 06 00 15 00 00 02 00 00 01 20 00 60 00    99 99 99 99 99 99 99 99 99 99 99 99    AA
        public byte[] READER_T_EPC = new byte[] { 0x81, 0x1A, 0x00, 0x06, 0x00, 0x15, 0x00, 0x00, 0x02, 0x00, 0x00, 0x01, 0x20, 0x00, 0x60, 0x00 };
        public byte[] READER_R_RSV = new byte[] { 0x82, 0x0E, 0x00, 0x08, 0x00, 0x09, 0x00, 0x81, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
        public byte[] READER_R_EPC = new byte[] { 0x82, 0x0E, 0x00, 0x08, 0x00, 0x09, 0x00, 0x81, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
        public byte[] READER_R_TID = new byte[] { 0x82, 0x0E, 0x00, 0x08, 0x00, 0x09, 0x00, 0x81, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
        public byte[] READER_R_USR = new byte[] { 0x82, 0x0E, 0x00, 0x08, 0x00, 0x09, 0x00, 0x81, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
        //'82 12 00 07 00 0D 00 02 (Bank Area 03) (Begin Addr 00 00 00 00)(TAG_PASSWS 00 00 00 00) (Write Data 11 11 11 11)
        public byte[] READER_W_RSV_ACCESS_PW = new byte[] { 0x82, 0x12, 0x00, 0x07, 0x00, 0x0D, 0x00, 0x02, 0x00, 0x02, 0x00, 0x00, 0x00 };
        public byte[] READER_W_RSV_KILL_PW = new byte[] { 0x82, 0x12, 0x00, 0x07, 0x00, 0x0D, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00 };
        //'public byte[] READER_W_EPC = new byte[] {82120007000D00020100000000};
        public byte[] READER_W_EPC = new byte[] {0x82, 0x1A, 0x00, 0x07, 0x00, 0x15, 0x00, 0x02, 0x01, 0x02, 0x00, 0x00, 0x00 };   //'4f1a00070015000201 02000000 00000000 e2003000390701110610d482 cbcbcb
        //'public byte[] READER_W_TID = new byte[] {82120007000D000202000000000000000011111111};
        public byte[] READER_W_USR = new byte[] { 0x82, 0x12, 0x00, 0x07, 0x00, 0x0D, 0x00, 0x02, 0x03, 0x00, 0x00, 0x00, 0x00 };
        public byte[] READER_W47B_USR = new byte[] { 0x82, 0x3D, 0x00, 0x07, 0x00, 0x49, 0x00, 0x02, 0x03, 0x00, 0x00, 0x00, 0x00 };
        public byte[] READER_W17B_USR = new byte[] { 0x82, 0x11, 0x00 };
        //'25 3d 00 07  00 49 00 02 03 00 00 00  00 00 00 00 00 b1 b2 b3  11 11 11 11 b8 00 00 00  00 00 00 00


        public FrmMain()
        {
            InitializeComponent();
        }

        private void FrmMainLoad(object sender, EventArgs e)
        {
            Device = new UsbHidDevice(0x1325, 0xC02E);
            Device.OnConnected += DeviceOnConnected;
            Device.OnDisConnected += DeviceOnDisConnected;
            Device.DataReceived += DeviceDataReceived;
            Device.Connect();
            initialCombobox();    
            // QueryDeviceCapabilities();
        }

        private void DeviceDataReceived(byte[] data)
        {
            AppendText(ByteArrayToString(data));
        }

        private void AppendText(string p)
        {
            ThreadSafe(() => textBox1.AppendText(p + Environment.NewLine));
        }

        private void DeviceOnDisConnected()
        {
            ThreadSafe(() => checkBox1.Enabled = false);

        }

        private void DeviceOnConnected()
        {
            ThreadSafe(() => checkBox1.Enabled = true);
        }

        private void ThreadSafe(MethodInvoker method)
        {
            if (InvokeRequired)
                Invoke(method);
            else
                method();
        }


        private static string ByteArrayToString(ICollection<byte> input)
        {
            var result = string.Empty;

            if (input != null && input.Count > 0)
            {
                var isFirst = true;
                foreach (var b in input)
                {
                    result += isFirst ? string.Empty : ",";
                    result += b.ToString("X2");
                    isFirst = false;
                }
            }
            return result;
        }

        private static byte[] StringToByteArray(string input,int byteLength)
        {
            input = input.Replace(" ", "").Replace(",", "");
            if (input.Length%2==1) return null;
            int len = input.Length / 2;
            byte[] result = new byte[len];
            for (int ix = 0;ix<len ; ix++)
            {
                result[ix] = Convert.ToByte("0x" + input.Substring(ix * 2, 2), 16); 
            }
            return result;
        }

        private void Button1Click(object sender, EventArgs e)
        {
            SendCommand(READER_T_TEST);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SendCommand(READER_T_STARTSCAN);
            //SendCommand(READER_T_EPC);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SendCommand(READER_T_STOPSCAN);
        }

        private void SendCommand(byte[] cmd)
        {
            if (!Device.IsDeviceConnected) return;
            var command = new CommandMessage(0x86, cmd);
            Device.SendMessage(command);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SendCommand(StringToByteArray(comboBox1.Text,0));
        }

        private void initialCombobox()
        {
            //comboBox1.Items.Add("20 11 00 86 00 02 00 00 00 0d 8c 00 05 00 00 01 01 00 01 06 cb cb cb cb  cb cb cb cb  cb cb cb cb");
            //comboBox1.Items.Add("21 0a 00 8c 00 05 00 00 01 00 00 00 00 cb cb cb");
            comboBox1.Items.Add("07 04 03 03 01 04");
            comboBox1.Items.Add("07 11 00 86 00 02 00 00 00 0D 8C 00 05 00 00 01 01 00 01 06");
            comboBox1.Items.Add("08 0A 00 8C 00 05 00 00 01 00 00 00 00");
            //comboBox1.Items.Add("81 1A 00 06 00 15 00 00 02 00 00 01 20 00 60 00 99 99 99 99 99 99 99");
            //comboBox1.Items.Add("");
            //comboBox1.Items.Add("");
            //comboBox1.Items.Add("");
            //comboBox1.Items.Add("");
            //comboBox1.Items.Add("");
            //comboBox1.Items.Add("");
            //comboBox1.Items.Add("");
            //comboBox1.Items.Add("");
            //comboBox1.Items.Add("");
            //comboBox1.Items.Add("");
            //comboBox1.Items.Add("");

        
        }

        private void button5_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
        }

    }
}

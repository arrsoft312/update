using System;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Ports;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

[assembly: AssemblyVersion("1.0")]
[assembly: ComVisible(false)]

[assembly: AssemblyTitle("update")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("Artur Kurpukov")]
[assembly: AssemblyProduct("update")]
[assembly: AssemblyCopyright("Copyright (C) 2019-2021 Artur Kurpukov")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]
[assembly: AssemblyFileVersion("1.0.0")]

class update:Form {
    private enum Board {
        ArduinoUNO,
        ArduinoMEGA2560,
    }
    
    private string fileName;
    private bool isHexFile;
    
    private Board board = Board.ArduinoUNO;
    
    private string portName = "COM1";
    private int baudRate = 115200;
    
    private UInt16 CRC16(byte ch, UInt16 oldCRC) {
        uint m = (((uint)oldCRC << 8) | ch);
        for (int n = 0; n < 8; n++) {
            m <<= 1;
            if ((m & 0x1000000) != 0) {
                m ^= 0x800500;
            }
        }
        return (UInt16)(m >> 8);
    }
    
    private void Form1DragEnter(object sender, DragEventArgs e) {
        if (backgroundWorker1.IsBusy) {
            return;
        }
        
        string[] fileNames = (string[])e.Data.GetData(DataFormats.FileDrop, false);
        if (fileNames.Length == 1) {
            e.Effect = DragDropEffects.All;
        } else {
            e.Effect = DragDropEffects.None;
        }
    }
    
    private void Form1DragDrop(object sender, DragEventArgs e) {
        if (backgroundWorker1.IsBusy) {
            return;
        }
        
        string[] fileNames = (string[])e.Data.GetData(DataFormats.FileDrop, false);
        if (fileNames.Length != 1) {
            return;
        }
        
        textBox1.Text = fileName = fileNames[0];
        comboBox1.Enabled = isHexFile = CultureInfo.InvariantCulture.CompareInfo.IsSuffix(fileName, ".hex", CompareOptions.IgnoreCase);
        comboBox2.Enabled = true;
        comboBox3.Enabled = true;
        progressBar1.Enabled = true;
        button2.Enabled = true;
    }
    
    private void Button1Click(object sender, EventArgs e) {
        if (openFileDialog1.ShowDialog(this) != DialogResult.OK) {
            return;
        }
        
        textBox1.Text = fileName = openFileDialog1.FileName;
        comboBox1.Enabled = isHexFile = CultureInfo.InvariantCulture.CompareInfo.IsSuffix(fileName, ".hex", CompareOptions.IgnoreCase);
        comboBox2.Enabled = true;
        comboBox3.Enabled = true;
        progressBar1.Enabled = true;
        button2.Enabled = true;
    }
    
    private void ComboBox1SelectedIndexChanged(object sender, EventArgs e) {
        board = (Board)((ComboBox)sender).SelectedItem;
    }
    
    private void ComboBox2DropDown(object sender, EventArgs e) {
        ((ComboBox)sender).Items.Clear();
        ((ComboBox)sender).Items.AddRange(SerialPort.GetPortNames());
    }
    
    private void ComboBox2TextChanged(object sender, EventArgs e) {
        portName = ((ComboBox)sender).Text;
    }
    
    private void ComboBox3TextChanged(object sender, EventArgs e) {
        try {
            baudRate = (int)UInt32.Parse(((ComboBox)sender).Text);
            ((Control)sender).BackColor = Color.Empty;
        } catch {
            ((Control)sender).BackColor = Color.DarkSalmon;
        }
    }
    
    private void Button2Click(object sender, EventArgs e) {
        button1.Enabled = false;
        comboBox1.Enabled = false;
        comboBox2.Enabled = false;
        comboBox3.Enabled = false;
        button2.Enabled = false;
        
        backgroundWorker1.RunWorkerAsync();
    }
    
    private void BackgroundWorker1ProgressChanged(object sender, ProgressChangedEventArgs e) {
        progressBar1.Value = e.ProgressPercentage;
    }
    
    private void BackgroundWorker1DoWork(object sender, DoWorkEventArgs e) {
        SerialPort comPort = new SerialPort(portName, baudRate);
        comPort.ReadTimeout = 3000;
        
        try {
            if (!isHexFile) {
                
                const byte Error_INCOMPATIBLE = 0x88;
                const byte Error_OK = 0x11;
                const byte Error_CRC = 0x22;
                
                FileStream inFile;
                try {
                    inFile = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
                } catch {
                    throw new Exception("Couldn't find '" + fileName + "'.");
                }
                
                byte[] buf = new byte[262144];
                int size;
                
                try {
                    size = inFile.Read(buf, 0, 262144);
                    
                    UInt32 crc32 = 0xFFFFFFFF;
                    for (int i = 0; i < size; i++) {
                        crc32 ^= buf[i];
                        for (int n = 0; n < 8; n++) {
                            if ((crc32 & 0x00000001) != 0) {
                                crc32 = ((crc32 >> 1) ^ 0xEDB88320);
                            } else {
                                crc32 >>= 1;
                            }
                        }
                    }
                    
                    if (crc32 != 0) {
                        throw new Exception("'" + fileName + "' is damaged.");
                    }
                } finally {
                    inFile.Close();
                }
                
                comPort.Open();
                
                size -= 4;
                for (int index = 0, retries = 0; index < size;) {
                    ((BackgroundWorker)sender).ReportProgress(1000 * index / size);
                    
                    int frameSize = 0;
                    
                    UInt16 crc = 0x0000;
                    do {
                        crc = CRC16(buf[index++], crc);
                    } while (++frameSize < 16 || index < 48);
                    
                    crc = CRC16(0, crc);
                    crc = CRC16(0, crc);
                    
                    comPort.Write(new byte[] { (byte)(frameSize+2), }, 0, 1);
                    comPort.Write(buf, (index-frameSize), frameSize);
                    comPort.Write(new byte[] { (byte)(crc >> 8), (byte)crc, }, 0, 2);
                    
                    byte resp = (byte)comPort.ReadByte();
                    if (resp == Error_OK) {
                        retries = 0;
                    } else {
                        if (resp == Error_INCOMPATIBLE) {
                            if (frameSize == 48) {
                                throw new WarningException("Unable to initialize bootloader!");
                            }
                        }
                        if (++retries >= 3) {
                            throw new WarningException("Data error (cyclic redundancy check).");
                        }
                        index -= frameSize;
                    }
                }
                
            } else if (board == Board.ArduinoUNO) {
                
                const byte Resp_STK_OK = 0x10;
                const byte Resp_STK_INSYNC = 0x14;
                
                const byte Sync_CRC_EOP = 0x20;
                
                const byte Cmnd_STK_LOAD_ADDRESS = 0x55;
                const byte Cmnd_STK_PROG_PAGE = 0x64;
                const byte Cmnd_STK_READ_PAGE = 0x74;
                const byte Cmnd_STK_READ_SIGN = 0x75;
                
                const int SPM_PAGESIZE = 128;
                
                byte sig0 = 0x1e;
                byte sig1 = 0x95;
                byte sig2 = 0x0f;
                
                MemoryMap flash = new MemoryMap(fileName);
                int size = flash.Size;
                
                comPort.Open();
                
                comPort.DtrEnable = true;
                comPort.RtsEnable = true;
                Thread.Sleep(250);
                
                //comPort.DtrEnable = false;
                //comPort.RtsEnable = false;
                //Thread.Sleep(50);
                
                comPort.DiscardInBuffer();
                
                byte[] buf = new byte[SPM_PAGESIZE];
                byte resp, resp2;
                
                comPort.Write(new byte[] { Cmnd_STK_READ_SIGN, }, 0, 1);
                comPort.Write(new byte[] { Sync_CRC_EOP, }, 0, 1);
                
                resp = (byte)comPort.ReadByte();
                sig0 ^= (byte)comPort.ReadByte();
                sig1 ^= (byte)comPort.ReadByte();
                sig2 ^= (byte)comPort.ReadByte();
                resp2 = (byte)comPort.ReadByte();
                
                if (resp != Resp_STK_INSYNC || resp2 != Resp_STK_OK) {
                    throw new WarningException("STK500 protocol error: Invalid response!");
                }
                
                if (sig0 != 0 || sig1 != 0 || sig2 != 0) {
                    throw new WarningException("Yikes! Invalid device signature.");
                }
                
                bool verify = false;
                for (int addr = 0; addr < size; addr += SPM_PAGESIZE) {
                    ((BackgroundWorker)sender).ReportProgress(1000 * addr / size);
                    
                    comPort.Write(new byte[] { Cmnd_STK_LOAD_ADDRESS, (byte)(addr >> 1), (byte)(addr >> 9), }, 0, 3);
                    comPort.Write(new byte[] { Sync_CRC_EOP, }, 0, 1);
                    
                    resp = (byte)comPort.ReadByte();
                    resp2 = (byte)comPort.ReadByte();
                    
                    if (resp != Resp_STK_INSYNC || resp2 != Resp_STK_OK) {
                        throw new WarningException("STK500 protocol error: Invalid response!");
                    }
                    
                    bool flag = false;
                    int i = 0;
                    
                    if (verify) {
                        comPort.Write(new byte[] { Cmnd_STK_READ_PAGE, (byte)(SPM_PAGESIZE >> 8), (byte)SPM_PAGESIZE, (byte)'F', }, 0, 4);
                    } else {
                        for (; i < SPM_PAGESIZE; i++) {
                            buf[i] = flash[addr+i];
                        }
                        
                        comPort.Write(new byte[] { Cmnd_STK_PROG_PAGE, (byte)(SPM_PAGESIZE >> 8), (byte)SPM_PAGESIZE, (byte)'F', }, 0, 4);
                        comPort.Write(buf, 0, SPM_PAGESIZE);
                    }
                    comPort.Write(new byte[] { Sync_CRC_EOP, }, 0, 1);
                    
                    resp = (byte)comPort.ReadByte();
                    for (; i < SPM_PAGESIZE; i++) {
                        if (comPort.ReadByte() != buf[i]) {
                            flag = true;
                        }
                    }
                    resp2 = (byte)comPort.ReadByte();
                    
                    if (resp != Resp_STK_INSYNC || resp2 != Resp_STK_OK) {
                        throw new WarningException("STK500 protocol error: Invalid response!");
                    }
                    
                    if (flag) {
                        throw new WarningException("Verification error: Content mismatch!");
                    }
                    
                    verify = !verify;
                    if (verify) {
                        addr -= SPM_PAGESIZE;
                    }
                }
                
                comPort.Write(new byte[] { (byte)'Q', }, 0, 1);
                comPort.Write(new byte[] { Sync_CRC_EOP, }, 0, 1);
                
                resp = (byte)comPort.ReadByte();
                resp2 = (byte)comPort.ReadByte();
                
                if (resp != Resp_STK_INSYNC || resp2 != Resp_STK_OK) {
                    throw new WarningException("STK500 protocol error: Invalid response!");
                }
                
            }
        } catch (TimeoutException) {
            throw new WarningException("Target is not responding!");
        } finally {
            comPort.Close();
        }
        
        ((BackgroundWorker)sender).ReportProgress(1000);
    }
    
    private void BackgroundWorker1RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
        Exception err = e.Error;
        if (err != null) {
            if (err is WarningException) {
                MessageBox.Show(this, err.Message, "update", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            } else {
                MessageBox.Show(this, err.Message, "update", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        } else {
            MessageBox.Show(this, "Target updated successfully!", "update", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        button1.Enabled = true;
        comboBox1.Enabled = isHexFile;
        comboBox2.Enabled = true;
        comboBox3.Enabled = true;
        progressBar1.Value = 0;
        button2.Enabled = true;
    }
    
    [STAThread]
    internal static void Main() {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new update());
    }
    
    public update() {
        tableLayoutPanel1 = new TableLayoutPanel();
        
        tableLayoutPanel2 = new TableLayoutPanel();
        label1 = new Label();
        textBox1 = new TextBox();
        button1 = new Button();
        
        tableLayoutPanel3 = new TableLayoutPanel();
        label2 = new Label();
        comboBox1 = new ComboBox();
        
        tableLayoutPanel4 = new TableLayoutPanel();
        label3 = new Label();
        comboBox2 = new ComboBox();
        label4 = new Label();
        comboBox3 = new ComboBox();
        
        tableLayoutPanel5 = new TableLayoutPanel();
        label5 = new Label();
        progressBar1 = new ProgressBar();
        button2 = new Button();
        
        openFileDialog1 = new OpenFileDialog();
        openFileDialog1.Title = "update";
        openFileDialog1.ReadOnlyChecked = true;
        openFileDialog1.Filter = "All files (*.*)|*.*";
        
        backgroundWorker1 = new BackgroundWorker();
        backgroundWorker1.WorkerReportsProgress = true;
        backgroundWorker1.DoWork += BackgroundWorker1DoWork;
        backgroundWorker1.ProgressChanged += BackgroundWorker1ProgressChanged;
        backgroundWorker1.RunWorkerCompleted += BackgroundWorker1RunWorkerCompleted;
        
        tableLayoutPanel2.SuspendLayout();
        tableLayoutPanel3.SuspendLayout();
        tableLayoutPanel4.SuspendLayout();
        tableLayoutPanel5.SuspendLayout();
        tableLayoutPanel1.SuspendLayout();
        this.SuspendLayout();
        
        tableLayoutPanel1.AutoSize = true;
        tableLayoutPanel1.ColumnCount = 1;
        tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle());
        tableLayoutPanel1.Controls.Add(tableLayoutPanel2, 0, 0);
        tableLayoutPanel1.Controls.Add(tableLayoutPanel3, 0, 1);
        tableLayoutPanel1.Controls.Add(tableLayoutPanel4, 0, 2);
        tableLayoutPanel1.Controls.Add(tableLayoutPanel5, 0, 3);
        tableLayoutPanel1.Dock = DockStyle.Fill;
        tableLayoutPanel1.RowCount = 4;
        tableLayoutPanel1.RowStyles.Add(new RowStyle());
        tableLayoutPanel1.RowStyles.Add(new RowStyle());
        tableLayoutPanel1.RowStyles.Add(new RowStyle());
        tableLayoutPanel1.RowStyles.Add(new RowStyle());
        tableLayoutPanel1.TabIndex = 0;
        
        tableLayoutPanel2.AutoSize = true;
        tableLayoutPanel2.ColumnCount = 3;
        tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70F));
        tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle());
        tableLayoutPanel2.Controls.Add(label1, 0, 0);
        tableLayoutPanel2.Controls.Add(textBox1, 1, 0);
        tableLayoutPanel2.Controls.Add(button1, 2, 0);
        tableLayoutPanel2.Dock = DockStyle.Fill;
        tableLayoutPanel2.Margin = new Padding(3, 3, 3, 0);
        tableLayoutPanel2.RowCount = 1;
        tableLayoutPanel2.RowStyles.Add(new RowStyle());
        tableLayoutPanel2.TabIndex = 0;
        
        label1.AutoSize = true;
        label1.Dock = DockStyle.Fill;
        label1.TabIndex = 0;
        label1.Text = "&Firmware:";
        label1.TextAlign = ContentAlignment.MiddleLeft;
        
        textBox1.Dock = DockStyle.Fill;
        textBox1.ReadOnly = true;
        textBox1.TabIndex = 1;
        textBox1.Text = fileName;
        
        button1.AutoSize = true;
        button1.Dock = DockStyle.Fill;
        button1.Enabled = true;
        button1.Size = new Size();
        button1.TabIndex = 2;
        button1.Text = "...";
        button1.UseVisualStyleBackColor = true;
        button1.Click += Button1Click;
        
        tableLayoutPanel3.AutoSize = true;
        tableLayoutPanel3.ColumnCount = 3;
        tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70F));
        tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180F));
        tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        tableLayoutPanel3.Controls.Add(label2, 0, 0);
        tableLayoutPanel3.Controls.Add(comboBox1, 1, 0);
        tableLayoutPanel3.Dock = DockStyle.Fill;
        tableLayoutPanel3.Margin = new Padding(3, 0, 3, 0);
        tableLayoutPanel3.RowCount = 1;
        tableLayoutPanel3.RowStyles.Add(new RowStyle());
        tableLayoutPanel3.TabIndex = 1;
        
        label2.AutoSize = true;
        label2.Dock = DockStyle.Fill;
        label2.TabIndex = 0;
        label2.Text = "Bo&ard:";
        label2.TextAlign = ContentAlignment.MiddleLeft;
        
        comboBox1.Dock = DockStyle.Fill;
        comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
        comboBox1.Enabled = false;
        comboBox1.Items.AddRange(new object[] { Board.ArduinoUNO, });
        comboBox1.SelectedItem = board;
        comboBox1.TabIndex = 1;
        comboBox1.SelectedIndexChanged += ComboBox1SelectedIndexChanged;
        
        tableLayoutPanel4.AutoSize = true;
        tableLayoutPanel4.ColumnCount = 6;
        tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70F));
        tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80F));
        tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20F));
        tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 50F));
        tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120F));
        tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        tableLayoutPanel4.Controls.Add(label3, 0, 0);
        tableLayoutPanel4.Controls.Add(comboBox2, 1, 0);
        tableLayoutPanel4.Controls.Add(label4, 3, 0);
        tableLayoutPanel4.Controls.Add(comboBox3, 4, 0);
        tableLayoutPanel4.Dock = DockStyle.Fill;
        tableLayoutPanel4.Margin = new Padding(3, 0, 3, 0);
        tableLayoutPanel4.RowCount = 1;
        tableLayoutPanel4.RowStyles.Add(new RowStyle());
        tableLayoutPanel4.TabIndex = 2;
        
        label3.AutoSize = true;
        label3.Dock = DockStyle.Fill;
        label3.TabIndex = 0;
        label3.Text = "&Port:";
        label3.TextAlign = ContentAlignment.MiddleLeft;
        
        comboBox2.Dock = DockStyle.Fill;
        comboBox2.Enabled = false;
        comboBox2.TabIndex = 1;
        comboBox2.Text = portName;
        comboBox2.DropDown += ComboBox2DropDown;
        comboBox2.TextChanged += ComboBox2TextChanged;
        
        label4.AutoSize = true;
        label4.Dock = DockStyle.Fill;
        label4.TabIndex = 2;
        label4.Text = "&Baud:";
        label4.TextAlign = ContentAlignment.MiddleLeft;
        
        comboBox3.Dock = DockStyle.Fill;
        comboBox3.Enabled = false;
        comboBox3.Items.AddRange(new object[] { 19200, 38400, 57600, 115200, 230400, });
        comboBox3.TabIndex = 3;
        comboBox3.Text = baudRate.ToString();
        comboBox3.TextChanged += ComboBox3TextChanged;
        
        tableLayoutPanel5.AutoSize = true;
        tableLayoutPanel5.ColumnCount = 3;
        tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70F));
        tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle());
        tableLayoutPanel5.Controls.Add(label5, 0, 0);
        tableLayoutPanel5.Controls.Add(progressBar1, 1, 0);
        tableLayoutPanel5.Controls.Add(button2, 2, 0);
        tableLayoutPanel5.Dock = DockStyle.Fill;
        tableLayoutPanel5.Margin = new Padding(3, 0, 3, 3);
        tableLayoutPanel5.RowCount = 1;
        tableLayoutPanel5.RowStyles.Add(new RowStyle());
        tableLayoutPanel5.TabIndex = 3;
        
        label5.AutoSize = true;
        label5.Dock = DockStyle.Fill;
        label5.TabIndex = 0;
        label5.Text = "Progress:";
        label5.TextAlign = ContentAlignment.MiddleLeft;
        
        progressBar1.Dock = DockStyle.Fill;
        progressBar1.Enabled = false;
        progressBar1.Maximum = 1000;
        progressBar1.Style = ProgressBarStyle.Continuous;
        progressBar1.TabIndex = 1;
        
        button2.AutoSize = true;
        button2.Dock = DockStyle.Fill;
        button2.Enabled = false;
        //button2.Size = new Size();
        button2.TabIndex = 2;
        button2.Text = "&Upload";
        button2.UseVisualStyleBackColor = true;
        button2.Click += Button2Click;
        
        this.AllowDrop = true;
        this.AutoScaleDimensions = new SizeF(6F, 13F);
        this.AutoScaleMode = AutoScaleMode.Font;
        this.AutoSize = true;
        this.ClientSize = new Size(380, -1);
        this.Controls.Add(tableLayoutPanel1);
        this.FormBorderStyle = FormBorderStyle.FixedSingle;
        this.MaximizeBox = false;
        this.Text = "update";
        this.DragDrop += Form1DragDrop;
        this.DragEnter += Form1DragEnter;
        
        this.ResumeLayout(false);
        this.PerformLayout();
        tableLayoutPanel1.ResumeLayout(false);
        tableLayoutPanel1.PerformLayout();
        tableLayoutPanel2.ResumeLayout(false);
        tableLayoutPanel2.PerformLayout();
        tableLayoutPanel3.ResumeLayout(false);
        tableLayoutPanel3.PerformLayout();
        tableLayoutPanel4.ResumeLayout(false);
        tableLayoutPanel4.PerformLayout();
        tableLayoutPanel5.ResumeLayout(false);
        tableLayoutPanel5.PerformLayout();
    }
    
    private BackgroundWorker backgroundWorker1;
    private OpenFileDialog openFileDialog1;
    private Button button2;
    private ProgressBar progressBar1;
    private Label label5;
    private TableLayoutPanel tableLayoutPanel5;
    private ComboBox comboBox3;
    private Label label4;
    private ComboBox comboBox2;
    private Label label3;
    private TableLayoutPanel tableLayoutPanel4;
    private ComboBox comboBox1;
    private Label label2;
    private TableLayoutPanel tableLayoutPanel3;
    private Button button1;
    private TextBox textBox1;
    private Label label1;
    private TableLayoutPanel tableLayoutPanel2;
    private TableLayoutPanel tableLayoutPanel1;
}

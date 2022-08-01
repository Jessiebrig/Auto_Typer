using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using AutoItX3Lib;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Keys = System.Windows.Forms.Keys;

namespace Auto_Typer
{
    public partial class Form1 : Form
    {
        //WebDriver_____WebDriver_____WebDriver_____WebDriver_____WebDriver_____WebDriver_____
        //IWebDriver Chrome = new ChromeDriver();
        IWebDriver Chrome;
        //Chrome_Profile_____Chrome_Profile_____Chrome_Profile_____Chrome_Profile_____Chrome_Profile_____
        public void Profile_With_Coordinates_Extension()
        {
            Chrome.Quit();
            ChromeOptions options = new ChromeOptions();
            options.AddArgument(@"user-data-dir=D:\chrome-dev-profile");
            Chrome = new ChromeDriver(options);
            AddMsg("Loading Custom Profile (With_Coordinates_Extension) Successful");
        }
        public void Chome_Taburnok_Profile()
        {
            Chrome.Quit();
            ChromeOptions options = new ChromeOptions();
            options.AddArgument(@"user-data-dir=D:\Chome_Taburnok_Profile");
            Chrome = new ChromeDriver(options);
            AddMsg("Loading Custom Profile (With_Coordinates_Extension) Successful");
        }
        
        public void Chome_Jessie_Profile()
        {
            Chrome.Quit();
            ChromeOptions options = new ChromeOptions();
            options.AddArgument(@"user-data-dir=C:\Users\User\AppData\Local\Google\Chrome\User Data");
            Chrome = new ChromeDriver(options);
            AddMsg("Loading Custom Profile (Chome_Jessie_Profile) Successful");
        }
        //
        public AutoItX3 Auto = new AutoItX3();
        public delegate void DELETEGATE();
        public String Status;
        public string text_box1;
        public string text_box2;
        public string text_box3;
        public string text_box5;
        public String Words;
        public String[] Word_Array;
        public static string ElementText;
        public static String Value = ".//span[@data-reactid='.0.2.0.0.$=12.0.$=10.1.0.$0']";
        public String Span_Status;
        public int Min;
        public int Max;

        public void MinMax()
        {
            Min = int.Parse(textBox7.Text);
            Max = int.Parse(textBox6.Text);
        }
        public void SetMinMax(int min, int max)
        {
            textBox7.Text = min.ToString();
            textBox6.Text = max.ToString();
            Min = min;
            Max = max;
        }
        //-------------------------------------
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string sClassname, string sAppName);
        [DllImport("user32.dll")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private IntPtr thiswindow;
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            thiswindow = FindWindow(null, "Form1");
            RegisterHotKey(thiswindow, 1, (uint)FsModifiers.Control, (uint)Keys.Add);
            RegisterHotKey(thiswindow, 2, (uint)FsModifiers.Control, (uint)Keys.N);
            RegisterHotKey(thiswindow, 3, (uint)FsModifiers.Control, (uint)Keys.Subtract);
            RegisterHotKey(thiswindow, 4, (uint)FsModifiers.Control, (uint)Keys.Space);
            RegisterHotKey(thiswindow, 5, (uint)FsModifiers.Control, (uint)Keys.Left);
            RegisterHotKey(thiswindow, 6, (uint)FsModifiers.Control, (uint)Keys.Enter);
            RegisterHotKey(thiswindow, 7, (uint)FsModifiers.Control, (uint)Keys.Right);
            textBox3.Text = "Ready";
            Status = "Stop";
            SetMinMax(56, 98);
        }
        public enum FsModifierss
        {
            Alt = 0x0001,
            Control = 0x0002,
            Shift = 0x0004,
            Window = 0x0008,
        }
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            UnregisterHotKey(thiswindow, 1);
        }
        protected override void WndProc(ref Message keyPressed)
        {
            if (keyPressed.Msg == 0x0312)
            {
                int id = keyPressed.WParam.ToInt32();
                switch (id)
                {
                    case 1:
                        Sendkeys();
                        break;
                    case 2:
                        textBox1.AppendText(Environment.NewLine);
                        Newline();
                        break;
                    case 3:
                        Undo();
                        break;
                    case 4:
                        {
                            if (Status == "Stop")
                            {
                                Status = "Type";
                                Access("TB5", " Typing");
                                Auto_Type();
                            }
                            else if (Status == "Type")
                            {
                                Status = "Stop";
                                Access("TB5", " Stopped");
                            }
                            else
                            {
                                Access("TB5", " Unkwon Status");
                            }
                            ON_OFF_TB6_N_TB7();
                        }
                        break;
                    case 5:
                        break;
                    case 6:
                        Space();
                        break;
                    case 7:
                        {
                            if (Status == "Stop")
                            {
                                Status = "Type";
                                Access("TB5", " Typing");
                                GetNSend();
                            }
                            else if (Status == "Type")
                            {
                                Status = "Stop";
                                Access("TB5", " Stopped");
                            }
                            else
                            {
                                Access("TB5", " Unkwon Status");
                            }
                            ON_OFF_TB6_N_TB7();
                        }
                        break;
                }
            }
            base.WndProc(ref keyPressed);
        }
        //Buttons_____Buttons_____Buttons_____Buttons_____Buttons_____Buttons_____Buttons_____Buttons_____
        private void WebDriver_Click(object sender, EventArgs e)
        {
            Chome_Jessie_Profile();
        }
        private void GetText_Click(object sender, EventArgs e)
        {
            Get_Element_Text();
            textBox1.Text = ElementText;
        }
        private void Clear_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            SetMinMax(0, 0);
        }
        private void NextLine_Click(object sender, EventArgs e)
        {
            Newline();
        }
        //Methods_____Methods_____Methods_____Methods_____Methods_____Methods_____Methods_____Methods_____
        //For https://www.livechatinc.com/typing-speed-test/#/
        private void Test_Click(object sender, EventArgs e)
        {
            AddMsg("");
        }
        public void ON_OFF_TB6_N_TB7()
        {
            if (Status == "Stop")
            {
                textBox7.Enabled = true;
                textBox6.Enabled = true;
            }
            else if (Status == "Type")
            {
                textBox7.Enabled = false;
                textBox6.Enabled = false;
            }
            
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            Thread thread = new Thread(() => { Navi(); });
            thread.Start();
        }
        public void Navi()
        {
            Chrome.Navigate().GoToUrl("https://www.livechatinc.com/typing-speed-test/#/");
        }
        public void GetNSend()
        {
            Thread thread = new Thread(() => { Get_Send(); });
            thread.Start();
        }
        public void Get_Element_Text()
        {
            ElementText = Chrome.FindElement(By.XPath(Value)).Text;
            AddMsg(Value + ": " + ElementText);
        }
        public void Get_Send()
        {
            if (Status == "Type")
            {
            Start_Again:
                if (Status == "Type")
                {
                    Span_Status = "Clear";
                    Get_Element_Text();
                    Words = ElementText;
                    Access("TB1", Words);
                Again:
                    Get_Sendkeys();
                    Access("Pause", "S_");
                    if (Span_Status == "End")
                    {
                        AddMsg("End of span");
                        Auto.Send(" ");
                        Access("TB2", " ");
                    }
                    else
                    {
                        goto Again;
                    }
                    goto Start_Again;
                }
                else if (Status == "Stop")
                {
                }
            }
            else if (Status == "Stop")
            {
                AddMsg("Typing Stop.");
            }
        }
        public void Get_Sendkeys()
        {
            try
            {
                char[] array = new char[1];
                Words.CopyTo(0, array, 0, 1);
                Access("TB3", array[0].ToString());
                Auto.Send(text_box3);
                Access("TB2", text_box3);
                Access("TB1_R1st_letter", null);
                Access("TB2_ScrollDown", null);
            }
            catch (ArgumentOutOfRangeException)
            {
                AddMsg("Index Out Of Range Exception or TextBox is Empty.");
                //Status = "Stop";
                Access("TB5", "Stopped");
                Span_Status = "End";
                Space();
            }
            catch (NullReferenceException)
            {
                AddMsg("Null Reference Exception");
            }
        }
        //
        //Auto_Type_____Auto_Type_____Auto_Type_____Auto_Type_____Auto_Type_____Auto_Type_____Auto_Type_____Auto_Type_____
        public void Auto_Type()
        {
            Thread thread = new Thread(() => { Type(); });
            thread.Start();
        }
        public void Space()
        {
            Auto.Send(" ");
        }
        public void Type()
        {
        again:
            if (Status == "Type")
            {
                if (text_box3 == " ")
                {
                    Sendkeys();
                    Access("Pause", "S_");
                    goto again;
                }
                else
                {
                    Sendkeys();
                    Access("Pause", null);
                    goto again;
                }
            }
            else if (Status == "Stop")
            {
                AddMsg("Typing Stop.");
            }
        }
        private void Stop_Click(object sender, EventArgs e)
        {
            Status = "Stop";
            Access("TB5", "Stopped");
            ON_OFF_TB6_N_TB7();
        }
        public void Sendkeys()
        {
            try
            {
                char[] array = new char[1];
                Words.CopyTo(0, array, 0, 1);//copy index_0 of Words to array
                Access("TB3", array[0].ToString());//Store array[0] to TextBox3 and text_box3
                Auto.Send(text_box3);//Send char
                Access("TB2", text_box3);//Append Sended
                Access("TB1_R1st_letter", null);//Remove index_0 in TextBox1
                Access("TB2_ScrollDown", null);//Scroll Down
            }
            catch (ArgumentOutOfRangeException)
            {
                AddMsg("Index Out Of Range Exception or TextBox is Empty.");
                Status = "Stop";
            }
            catch (NullReferenceException)
            {
                AddMsg("Null Reference Exception");
            }
        }
        public void Newline()
        {
            char[] array = new char[1];
            Words.CopyTo(0, array, 0, 1);
            textBox3.Text = array[0].ToString();
            textBox2.Text += textBox3.Text;
            String after = textBox1.Text.Remove(0, 1);
            textBox1.Text = after;
            //---------------------------------------
            int numOfLines = 1;
            var lines = this.textBox1.Lines;
            Word_Array = this.textBox1.Lines;
            var newLines = lines.Skip(numOfLines);
            this.textBox1.Lines = newLines.ToArray();
            textBox2.AppendText(Environment.NewLine);
        }
        public void Undo()
        {
            try
            {
                var words = textBox2.Text.Last();
                textBox3.Text = words.ToString();
                var test = textBox1.Text.Insert(0, textBox3.Text);
                this.textBox1.Text = test;
                //-------------------------------------------------
                var lastword = textBox2.Text.LastIndexOf(words);
                string after = textBox2.Text.Remove(lastword, 1);
                textBox2.Text = after;
            }
            catch (InvalidOperationException)
            {
                AddMsg("Invalid Operation Exception");
            }
        }
        //
        //TextBox_____TextBox_____TextBox_____TextBox_____TextBox_____TextBox_____TextBox_____
        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            Words = textBox1.Text;
        }
        public void AddMsg(string Msg)
        {
            Delegate del = new DELETEGATE(() => 
            {
                textBox4.Text += Msg;
                textBox4.SelectionStart = textBox4.TextLength;
                textBox4.ScrollToCaret();
                textBox4.AppendText(Environment.NewLine);
            });
            this.Invoke(del);
        }
        public void Access(string code, string text)
        {
            if (code == "TB1")
            {
                Delegate del = new DELETEGATE(() => { textBox1.Text += text; text_box1 = text; });
                this.Invoke(del);
            }
            if (code == "TB2")
            {
                Delegate del = new DELETEGATE(() => { textBox2.Text += text; text_box2 = text; });
                this.Invoke(del);
            }
            if (code == "TB3")
            {
                Delegate del = new DELETEGATE(() => { textBox3.Text = text; text_box3 = text; });
                this.Invoke(del);
            }
            if (code == "TB5")
            {
                Delegate del = new DELETEGATE(() => { textBox5.Text = text; text_box5 = text; });
                this.Invoke(del);
            }
            if (code == "Pause")
            {
                Delegate del = new DELETEGATE(() => 
                {
                    MinMax();
                    Random time = new Random();
                    int num = time.Next(Min, Max);
                    textBox4.Text += text + num.ToString() + ", ";
                    Thread.Sleep(num);
                });
                this.Invoke(del);
            }
            //
            if (code == "TB1_R1st_letter") //Remove 1st letter of the textbox1
            {
                Delegate del = new DELETEGATE(() => { textBox1.Text = textBox1.Text.Remove(0, 1); });
                this.Invoke(del);
            }
            if (code == "TB2_ScrollDown") //Scroll textbox2
            {
                Delegate del = new DELETEGATE(() => { textBox2.SelectionStart = textBox2.TextLength; textBox2.ScrollToCaret(); });
                this.Invoke(del);
            }

        }
    }
}

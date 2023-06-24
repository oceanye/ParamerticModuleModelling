using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Timers;

namespace ParamerticModuleModelling
{
    public partial class TimerForm : Form
    {
        public TimerForm()
        {
            InitializeComponent();
        }

        System.Timers.Timer MyTimer;

        public delegate void SetControlValue(long value);
        public long TimeCount;

        private void TimerForm_Load(object sender, EventArgs e)
        {
            int interval = 1000;

            MyTimer = new System.Timers.Timer(interval);


            MyTimer.AutoReset = true;

            MyTimer.Elapsed += new ElapsedEventHandler(MyTimer_tick);
            MyTimer.Start();
            TimeCount = 0;
        }

       

        private void MyTimer_tick(object sender, System.Timers.ElapsedEventArgs e)
        {

            this.Invoke(new SetControlValue(ShowTime), TimeCount);
            TimeCount++;

        }

        private void ShowTime(long t)
        {
            TimeSpan temp = new TimeSpan(0, 0, (int)t);
            textBox1.Text = string.Format("{0:00}:{1:00}", temp.Minutes, temp.Seconds);



        }
    }
}

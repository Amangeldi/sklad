﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace sklad
{
    public partial class Responsible : Form
    {
        public int Uid;
        public Responsible(int _id)
        {
            InitializeComponent();
            this.Uid = _id;
        }

        private void Responsible_Load(object sender, EventArgs e)
        {

        }
    }
}

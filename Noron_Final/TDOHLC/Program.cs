﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace TVMH
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Form1 frmNoron = new Form1();
            frmDN frmdn = new frmDN();
           
            Application.Run(frmdn);
        }
    }
}

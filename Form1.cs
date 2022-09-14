using AxMicrosoft.Office.Interop.VisOcx;
using Microsoft.Office.Interop.VisOcx;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VisioDrawingControlCrash
{
	public partial class Form1 : Form
	{
		AxDrawingControl _drawingControl;

		public Form1()
		{
			InitializeComponent();

			_drawingControl = new AxDrawingControl();
			_drawingControl.Dock = DockStyle.Fill;

			Controls.Add(_drawingControl);

			Shown += Form1_Shown;
		}

		private void Form1_Shown(object sender, EventArgs e)
		{
			// setting shutdown hehavior has impact on the faulting module
			// shutdown behavior = 0 -> raises Faulting module name: MsoAria.dll, version: 16.0.15601.20038
			// shutdown behavior = 1 -> raises 
			_drawingControl.ShutdownBehavior = 1;
		}
	}
}

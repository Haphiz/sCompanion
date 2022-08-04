using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace sCompanion
{
    public partial class frmDialogue : Form
    {
        string dialogue;
        public frmDialogue(string dialogue)
        {
            InitializeComponent();
            this.dialogue = dialogue;
        }

        private void frmDialogue_Load(object sender, EventArgs e)
        {
            lbl.Text = dialogue;
            this.Location = new Point(20, 20);
        }

        private void btnok_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }
    }
}

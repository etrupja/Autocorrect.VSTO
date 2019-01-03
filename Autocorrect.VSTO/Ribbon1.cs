using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace Autocorrect.VSTO
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

     

        private void correctselected_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void correctall_Click(object sender, RibbonControlEventArgs e)
        {
            var wordDoc = Globals.ThisAddIn.Application.ActiveDocument;
            
            for(var i = 1; i <= wordDoc.Words.Count; i++)
            {
                var range = wordDoc.Words[i];
                if (string.IsNullOrEmpty(range.Text)) continue;
                var replacementText = "blablabla";
                if (!string.IsNullOrEmpty(replacementText)) range.Text = replacementText;
            }
        }
    }
}

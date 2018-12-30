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

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Selection sel = doc.Application.Selection;
            object unit = Word.WdUnits.wdCharacter;
            object count = 1;
            object extend = Word.WdMovementType.wdMove;
            sel.MoveLeft(ref unit, ref count, ref extend);
            object unit1 = Word.WdUnits.wdWord;
            object extend1 = Word.WdMovementType.wdExtend;
            doc.Application.Selection.MoveLeft(ref unit1, ref count, ref extend1);
            doc.Application.Selection.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
        }
    }
}

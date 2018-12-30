using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;
using Autocorrect.Common;
using Autocorrect.Api.Services;

namespace Autocorrect.VSTO
{
    public partial class ThisAddIn 
    {
        private readonly AddinHelper _helper = new AddinHelper();
        private SpellChecker _spellChecker;
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _helper.RegisterEvents();
            _helper.OnKeyUp += OnKeyUp;
            _spellChecker = new SpellChecker();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _helper.UnRegisterEvents();
        }

       

        private async void OnKeyUp(object sender,KeyEventArgs args)
        {
            //return;
            if (args.KeyCode != Keys.Space) return;
            var doc = Globals.ThisAddIn.Application.ActiveDocument;

            Word.Selection sel = doc.Application.Selection;
            object unit = Word.WdUnits.wdCharacter;
            object count = 1;
            object extend = Word.WdMovementType.wdMove;
            //sel.MoveLeft(ref unit, ref count, ref extend);
            object unit1 = Word.WdUnits.wdWord;
            object extend1 = Word.WdMovementType.wdExtend;
            object collapseDirection = Word.WdCollapseDirection.wdCollapseEnd;
            doc.Application.Selection.MoveLeft(ref unit1, ref count, ref extend1);
            var text = doc.Application.Selection.Text;
            var result =await _spellChecker.CheckSpell(text);
            if (!string.IsNullOrWhiteSpace(result)) doc.Application.Selection.Text = result;
            doc.Application.Selection.Collapse(ref collapseDirection);
        }




        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }

}

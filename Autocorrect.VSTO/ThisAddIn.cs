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
using Autocorrect.Licensing;
using Autocorrect.VSTO.Properties;
using Sentry;
using System.Threading.Tasks;
using Autocorrect.VSTO.Settigs;

namespace Autocorrect.VSTO
{
    public partial class ThisAddIn 
    {
        private readonly AddinHelper _helper = new AddinHelper();
        private SpellChecker _spellChecker;
        private  void ThisAddIn_Startup(object sender, EventArgs e)
        {
            //only start our application if license is valid
          
                using (SentrySdk.Init("https://e6d85d0ca9e941bf9d6ca3a207ea31fb@sentry.io/1368700"))
                {

                try
                {
                    if (LicenseManager.IsLicenseValid())
                    {
                        _helper.RegisterEvents();
                        _helper.OnKeyUp += OnKeyUp;
                        _spellChecker = new SpellChecker();
                    }
                }
                catch (Exception ex)
                {

                    SentrySdk.CaptureException(ex);
                }
                }
           
        }
     
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _helper.UnRegisterEvents();
        }

        KeysConverter KeyConverter = new KeysConverter();

        private Keys[] DoubleKeyArray = new Keys[] { Keys.C,Keys.E };
        private Keys[] EndOfWordKeyArray = new Keys[] { Keys.Oemcomma, Keys.OemQuestion, Keys.OemSemicolon, Keys.OemQuotes, Keys.Oem7, Keys.Oem1, Keys.Space, Keys.OemPeriod };
        private  void OnKeyUp(object sender,KeyEventArgs args)
        {
            if (GlobalSettings.AutocorrectDisabled) return;
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            if (DoubleKeyArray.Contains(args.KeyCode))
            {
                // handle non word end cases
                Word.Selection sel = doc.Application.Selection;
                //string keyChar = KeyConverter.ConvertToString(args.KeyData);
                ParseDoubleLetters(sel);


            }
            else if(EndOfWordKeyArray.Contains(args.KeyCode))
            {
                //word end case

                Word.Selection sel = doc.Application.Selection;
                var text = sel.Text.ToLower();
                if (text.Contains("ç") || text.Contains("ë"))
                {
                    //if the word contains this chars already no need to check against dictionary
                }
                else
                {
                    ParseWordFromDictionary(sel);
                }
            }
           

           
        }
        private void ParseDoubleLetters(Word.Selection selection)
        {
            ExtendSelectionLeft(selection, Word.WdUnits.wdCharacter,2);
            var text = selection.Text;
            if (!string.IsNullOrEmpty(text) && text.Length>=2)
            {
                text = text.Substring(text.Length - 2);
                var isUpperCase = char.IsUpper(text.First());
                string replacementText = text;
                switch (text.ToLower())
                {
                    case "cc":
                        replacementText = "ç";
                        break;
                    case "ee":
                        replacementText = "ë";
                        break;
                    default:
                        break;
                }
                selection.Text = isUpperCase ? replacementText.ToUpper() : replacementText.ToLower();
            }
            
            CollaseSelection(selection);
        }
        private void ParseWordFromDictionary(Word.Selection selection)
        {
            ExtendSelectionLeft(selection, Word.WdUnits.wdWord);
           
            var text = selection.Text;
            var lastCharacter = text.Last().ToString();
           var result = _spellChecker.CheckSpell(text.Trim());
            if (!string.IsNullOrWhiteSpace(result)) selection.Text = result + lastCharacter;
            CollaseSelection(selection);
        }
        //Word.Selection sel = doc.Application.Selection;
        //object unit = Word.WdUnits.wdCharacter;
        //object count = 1;
        //object extend = Word.WdMovementType.wdMove;
        ////sel.MoveLeft(ref unit, ref count, ref extend);
        //object unit1 = Word.WdUnits.wdWord;
        //object extend1 = Word.WdMovementType.wdExtend;
        //object collapseDirection = Word.WdCollapseDirection.wdCollapseEnd;
        //doc.Application.Selection.MoveLeft(ref unit1, ref count, ref extend1);
        //    var text = doc.Application.Selection.Text;
        //var result = _spellChecker.CheckSpell(text);
        //    if (!string.IsNullOrWhiteSpace(result)) doc.Application.Selection.Text = result;
        //    doc.Application.Selection.Collapse(ref collapseDirection);
        private void ExtendSelectionLeft(Word.Selection sel, Word.WdUnits unit, object amount=null)
        {
            if (amount == null) amount = 1;
           
            object extend = Word.WdMovementType.wdExtend;
            object collapseDirection = Word.WdCollapseDirection.wdCollapseEnd;
            object moveUnit = unit as object;
            sel.MoveLeft(ref moveUnit, ref amount, ref extend);
           
        }
        private void MoveSelectionLeft(Word.Selection sel, Word.WdUnits unit, object amount = null)
        {
            if (amount == null) amount = 1;

            object extend = Word.WdMovementType.wdMove;
            object collapseDirection = Word.WdCollapseDirection.wdCollapseEnd;
            object moveUnit = unit as object;
            sel.MoveLeft(ref moveUnit, ref amount, ref extend);

        }

        
        //sel.MoveLeft(ref unit, ref count, ref extend);
        private void CollaseSelection(Word.Selection sel, Word.WdCollapseDirection direction = Word.WdCollapseDirection.wdCollapseEnd)
        {
            object collapseDirection = direction;
            sel.Collapse(ref collapseDirection);
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

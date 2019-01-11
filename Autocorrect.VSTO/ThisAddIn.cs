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
                         SyncData();
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
        public async Task SyncData()
        {
            try
            {
                DataProvider.SyncData().Wait();
                Settings.Default.LastSync = DateTime.Now;
                Settings.Default.Save();
            }
            catch (Exception ex)
            {
                if (Settings.Default.LastSync == DateTime.MinValue)
                {
                    MessageBox.Show("Nje problem ka ndodhur duke kontaktuar serverin. Sigurohuni qe keni nje lidhje interneti dhe mbyllenin dhe hapenin applikacionin perseri qe te marrim te dhenat e fundit", "Problem duke kontaktuar serverin");
                    SentrySdk.CaptureException(ex);
                }
                   
            }

        }
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _helper.UnRegisterEvents();
        }

       

        private  void OnKeyUp(object sender,KeyEventArgs args)
        {
            if (GlobalSettings.AutocorrectDisabled) return;
            if (!new Keys[] {Keys.Oemcomma, Keys.OemQuestion, Keys.OemSemicolon, Keys.OemQuotes,Keys.Oem7, Keys.Oem1, Keys.Space, Keys.OemPeriod }.Contains(args.KeyCode)) return;
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
            var result = _spellChecker.CheckSpell(text);
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

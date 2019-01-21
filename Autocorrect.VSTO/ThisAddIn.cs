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
using Microsoft.Office.Core;

namespace Autocorrect.VSTO
{
    public partial class ThisAddIn 
    {
        private readonly AddinHelper _helper = new AddinHelper();
        private SpellChecker _spellChecker;
        _CommandBarButtonEvents_ClickEventHandler eventHandler;

        private void ThisAddIn_Startup(object sender, EventArgs e)
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

            try
            {
                eventHandler = new _CommandBarButtonEvents_ClickEventHandler(MyButton_Click);
                Word.Application applicationObject =
            Globals.ThisAddIn.Application as Word.Application;
                applicationObject.WindowBeforeRightClick +=
        new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(App_WindowBeforeRightClick);
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error: " + exception.Message);
            }

        }

        private void MyButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            throw new NotImplementedException();
        }


        private void Add()
        {
            RemoveItem();
            return;
            Word.Application applicationObject = Globals.ThisAddIn.Application as Word.Application;

            CommandBarPopup teksSakteMenu = applicationObject.CommandBars.FindControl(MsoControlType.msoControlPopup, missing, "tekssaktemenu", missing)  as CommandBarPopup;
            if (teksSakteMenu != null) return;
            else
            {
                CommandBar popupCommandBar = applicationObject.CommandBars["Text"];
                teksSakteMenu = (CommandBarPopup)popupCommandBar.Controls.Add(MsoControlType.msoControlPopup, missing, missing, missing, true);
                teksSakteMenu.Caption = "TeksSakte";
                teksSakteMenu.Tag = "tekssaktemenu";
            }
          

        //    bool isFound = false;
        //    foreach (object _object in popupCommandBar.Controls)
        //    {
        //        CommandBarButton _commandBarButton = _object as CommandBarButton;
        //        if (_commandBarButton == null) continue;
        //        if (_commandBarButton.Tag.Equals("TekstSakteMenu"))
        //        {
        //            isFound = true;
        //            _commandBarButton.Click += eventHandler;
        //            break;
        //        }
        //    }

        //    if (!isFound)
        //    {
        //        commandBarButton = (CommandBarButton)popupCommandBar.Controls.Add
        //(MsoControlType.msoControlButton, missing, missing, missing, true);
        //        System.Diagnostics.Debug.WriteLine("Created new button, adding handler");
        //        commandBarButton.Click += eventHandler;
        //        commandBarButton.Caption = "Hello !!!";
        //        commandBarButton.FaceId = 54854;
        //        commandBarButton.Tag = "TekstSakteMenu";
        //        commandBarButton.BeginGroup = true;
        //    }

          
            Office.CommandBarButton sugjeroMenu = (Office.CommandBarButton)teksSakteMenu.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, missing);
            sugjeroMenu.Caption = "Sugjero";
            sugjeroMenu.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(suggest_click);
            sugjeroMenu.Tag = "sugjeromenu";
        }


        private void AddItem()
        {
            Word.Application applicationObject =
        Globals.ThisAddIn.Application as Word.Application;
            CommandBarButton commandBarButton =
        applicationObject.CommandBars.FindControl
        (MsoControlType.msoControlButton, missing, "HELLO_TAG1", missing)
        as CommandBarButton;

            foreach (CommandBar item in applicationObject.CommandBars)
            {
                Debug.WriteLine(item.Name);
            }

            if (commandBarButton != null)
            {
                System.Diagnostics.Debug.WriteLine("Found button, attaching handler");
                commandBarButton.Click += eventHandler;
                return;
            }
            CommandBar popupCommandBar = applicationObject.CommandBars["Spelling"];
            bool isFound = false;
            foreach (object _object in popupCommandBar.Controls)
            {
                CommandBarButton _commandBarButton = _object as CommandBarButton;
                if (_commandBarButton == null) continue;
                if (_commandBarButton.Tag.Equals("HELLO_TAG1"))
                {
                    isFound = true;
                    System.Diagnostics.Debug.WriteLine
            ("Found existing button. Will attach a handler.");
                    _commandBarButton.Click += eventHandler;
                    break;
                }
            }
            if (!isFound)
            {
                commandBarButton = (CommandBarButton)popupCommandBar.Controls.Add
        (MsoControlType.msoControlButton, missing, missing, missing, true);
                System.Diagnostics.Debug.WriteLine("Created new button, adding handler");
                commandBarButton.Click += eventHandler;
                commandBarButton.Caption = "Hello !!!";
                commandBarButton.FaceId = 353;
                commandBarButton.Tag = "HELLO_TAG1";
                commandBarButton.BeginGroup = true;
            }
            Office.CommandBarPopup cpp = (CommandBarPopup)popupCommandBar.Controls.Add(MsoControlType.msoControlButtonPopup, missing, missing, missing, true);
            cpp.Caption = "SubMenu";
            cpp.Tag = "submenu";
            //= (Office.CommandBarPopup)ContextMenu.Controls.Add(Office.MsoControlType.msoControlPopup, missing, missing, missing, missing);


            Office.CommandBarButton cbHello3 = (Office.CommandBarButton)(cpp).Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, missing);
            cbHello3.Caption = "Hello3";
            cbHello3.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(suggest_click);
            cbHello3.FaceId = 353;
        }

        private void RemoveItem()
        {
            Word.Application applicationObject =
        Globals.ThisAddIn.Application as Word.Application;
            CommandBar popupCommandBar = applicationObject.CommandBars["Text"];
            popupCommandBar.Reset();
            foreach (object _object in popupCommandBar.Controls)
            {
                CommandBarButton commandBarButton = _object as CommandBarButton;
                if (commandBarButton == null) continue;
                if (commandBarButton.Caption.Equals("TeksSakte"))
                {
                    popupCommandBar.Reset();
                }
            }
            popupCommandBar = applicationObject.CommandBars["Spelling"];
            popupCommandBar.Reset();
            foreach (object _object in popupCommandBar.Controls)
            {
                CommandBarButton commandBarButton = _object as CommandBarButton;
                if (commandBarButton == null) continue;
                if (commandBarButton.Caption.Equals("TeksSakte"))
                {
                    popupCommandBar.Reset();
                }
            }
        }
        private void AddContextMenu()
        {
            Office.CommandBar ContextMenu = this.Application.CommandBars.Add("ContextMenu", Office.MsoBarPosition.msoBarPopup, missing, true);

            if (ContextMenu != null)
            {
                Office.CommandBarButton cbHello1 = (Office.CommandBarButton)ContextMenu.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, missing);
                cbHello1.Caption = "Hello1";
                cbHello1.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(suggest_click);

                Office.CommandBarButton cbHello2 = (Office.CommandBarButton)ContextMenu.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, missing);
                cbHello2.Caption = "Hello2";
                cbHello2.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(suggest_click);

                Office.CommandBarPopup cpp = (Office.CommandBarPopup)ContextMenu.Controls.Add(Office.MsoControlType.msoControlPopup, missing, missing, missing, missing);
                cpp.Caption = "SubMenu";

                Office.CommandBarButton cbHello3 = (Office.CommandBarButton)cpp.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, missing);
                cbHello3.Caption = "Hello3";
                cbHello3.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(suggest_click);

                Office.CommandBarButton cbHello4 = (Office.CommandBarButton)cpp.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, missing);
                cbHello4.Caption = "Hello4";
                cbHello4.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(suggest_click);

                ContextMenu.ShowPopup(missing, missing);
            }
        }

        private void suggest_click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
        }

        //private void AddExampleMenuItem()
        //{
        //    Office.MsoControlType menuItem = Office.MsoControlType.msoControlButton;
        //    Office.CommandBarButton exampleMenuItem = (Office.CommandBarButton)GetCellContextMenu().Controls.Add(menuItem, missing, missing, 1, true);

        //    exampleMenuItem.Style = Office.MsoButtonStyle.msoButtonCaption;
        //    exampleMenuItem.Caption = "Example Menu Item";
        //    exampleMenuItem.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(exampleMenuItemClick);
        //}
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
            Application.WindowBeforeRightClick -=   new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(App_WindowBeforeRightClick);
        }

        private void App_WindowBeforeRightClick(Word.Selection Sel, ref bool Cancel)
        {
            try
            {
                Add();
                Add();
            }
            catch (Exception exception)
            {
                MessageBox.Show("Error: " + exception.Message);
            }
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

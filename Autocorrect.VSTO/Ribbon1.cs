﻿using Autocorrect.Api.Services;
using Autocorrect.Licensing;
using Autocorrect.VSTO.Properties;
using Autocorrect.VSTO.Settigs;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Sentry;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Autocorrect.VSTO
{
    public partial class Ribbon1
    {
        private SpellChecker _spellChecker = new SpellChecker();
        private object missing;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            var hasLicense = LicenseManager.HasLicense();
            if (hasLicense)
            {
                LicenseDetails.Visible = true;
                expirationDateValueLabel.Label = LicenseManager.ExpirationDate()?.ToLongDateString();
                if (!LicenseManager.IsLicenseValid())
                {
                    hasExpired.ShowLabel = true;
                    hasExpired.Label = "License nuk eshte e sakte";

                    license.Label = "Rregjistrohuni";
                }
                else
                {
                    license.Label = "Ndryhoni License";
                }
            }
            else
            {
                LicenseDetails.Visible = false;
                license.Label = "Rregjistrohuni";
            }
            GlobalSettings.AutocorrectDisabled = !Properties.Settings.Default.Autcorrect;
        }

        private void Ribbon1_Close(object sender, EventArgs e)
        {

        }


        private void correctselected_Click(object sender, RibbonControlEventArgs e)
        {
            var words = Globals.ThisAddIn.Application.Selection.Words;
            CorrectWords(words);
        }

        private void correctall_Click(object sender, RibbonControlEventArgs e)
        {
            var wordDoc = Globals.ThisAddIn.Application.ActiveDocument;
            CorrectWords(wordDoc.Words);
        }

        private void CorrectWords(Words words)
        {
            foreach (Range range in words)
            {
                var value = range.Text.ToString();
                var spacesFromStart = value.TakeWhile(Char.IsWhiteSpace).Count();
                var spacesFromEnd = value.Reverse().TakeWhile(c => Char.IsWhiteSpace(c) || c == '.').Count();
                range.Start += spacesFromStart;
                range.End -= spacesFromEnd;
                if (string.IsNullOrEmpty(range.Text)) continue;
                var replacementText = _spellChecker.CheckSpell(range.Text);
                if (!string.IsNullOrEmpty(replacementText)) range.Text = replacementText;
            }
        }

        private async void license_Click(object sender, RibbonControlEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                CheckFileExists = true,
                AddExtension = true,
                Filter = "License Files|*.lic"
            };
            var result = dialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                var file = dialog.FileName;
                if (File.Exists(file))
                {

                    using (var fileStream = File.OpenRead(file))
                    { 
                        var license = LicenseManager.ParseLicense(fileStream);
                        var isValid = LicenseManager.IsValid(license);
                        if (isValid)
                        {
                            try
                            {
                                var isOnlineValid = await LicenseManager.ValidateLicenseOnline(license.Id);
                                if (isOnlineValid)
                                {
                                    await LicenseManager.UpdateLicenseUtilizedCount(license.Id);
                                    await LicenseManager.SetLicense(fileStream);
                                    MessageBox.Show("Rregjistrimi u krye me sukses. Ju lutemi mbylleni applikacionin qe te shikoni ndryshimet", "Sukses");


                                    await DataProvider.SyncData();
                                    Settings.Default.LastSync = DateTime.Now;
                                    Settings.Default.Save();

                                }
                                else
                                {
                                    MessageBox.Show("Licensa ne fjale ka arritur maksimumin e perdorimit", "Licensa nuk eshte e vlefshme");
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Nje problem ka ndodhur duke kontaktuar serverin. Sigurohuni qe keni nje lidhje interneti dhe mbyllenin dhe hapenin applikacionin perseri qe te marrim te dhenat e fundit", "Problem duke kontaktuar serverin");

                                SentrySdk.CaptureException(ex);
                            }
                        }
                        else
                        {
                            MessageBox.Show("License nuk eshte e vlefshme ose ka skaduar", "Licensa nuk eshte e vlefshme");
                        }

                    }
                }
            }
        }

        private void autocorrectToggle_Click(object sender, RibbonControlEventArgs e)
        {
            GlobalSettings.AutocorrectDisabled = !autocorrectToggle.Checked;
            Properties.Settings.Default.Autcorrect = autocorrectToggle.Checked;
            Properties.Settings.Default.Save();
        }

       
        public async void SyncData()
        {
            try
            {
               await  DataProvider.SyncData();
                Settings.Default.LastSync = DateTime.Now;
                Settings.Default.Save();
                MessageBox.Show("Perditesimi u krye me sukses", "Sukses");

            }
            catch (Exception ex)
            {
                MessageBox.Show("Nje problem ka ndodhur duke kontaktuar serverin. Sigurohuni qe keni nje lidhje interneti", "Problem duke kontaktuar serverin");
                SentrySdk.CaptureException(ex);
            }

        }

        private void perditesoButton_Click(object sender, RibbonControlEventArgs e)
        {
           
          SyncData();
        }

        private void changeLangaugeToAlbanian_Click(object sender, RibbonControlEventArgs e)
        {
           var range = Globals.ThisAddIn.Application.ActiveDocument.Content.LanguageID = WdLanguageID.wdAlbanian;
        }

        private void shenjatepikesimit_Click(object sender, RibbonControlEventArgs e)
        {
            KorrigjoShenjatEPikesimit();
        }
        private void FindAndReplace(Range range, object findText, object replaceWithText)
        {
            //options
            object matchCase = false;
            object matchWholeWord = false;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = true;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;

          
            range.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replaceAll,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }
        private void KorrigjoShenjatEPikesimit()
        {
            var wordDoc = Globals.ThisAddIn.Application;
            var shenjatEPIkesimit = new List<string> { ",", "?", "!", ".", ":",";"};
            foreach (var item in shenjatEPIkesimit)
            {
                FindAndReplace(wordDoc.ActiveDocument.Content, $" {item}", item);
            }
            foreach (Range range in wordDoc.ActiveDocument.Content.Words)
            {
                var value = range.Text;
                var index = shenjatEPIkesimit.IndexOf(value);
                if (index >= 0) range.Text = $"{shenjatEPIkesimit[index]} ";


            }
        }
    }
}

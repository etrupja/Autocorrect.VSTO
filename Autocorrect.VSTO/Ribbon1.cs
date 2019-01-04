﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autocorrect.Api.Services;
using Autocorrect.Licensing;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace Autocorrect.VSTO
{
    public partial class Ribbon1
    {
        private SpellChecker _spellChecker =new SpellChecker();

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
            
                    licensing.Visible = true;
                }
                else
                {
                    licensing.Visible = false;
                }
            }
            else
            {
                LicenseDetails.Visible = false;

            }
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

        private  void CorrectWords(Words words)
        {
            foreach (Range range in words)
            {
                var value = range.Text.ToString();
                var spacesFromStart = value.TakeWhile(Char.IsWhiteSpace).Count();
                var spacesFromEnd = value.Reverse().TakeWhile( c=> Char.IsWhiteSpace(c) ||  c == '.').Count();
                range.Start += spacesFromStart;
                range.End -= spacesFromEnd;
                if (string.IsNullOrEmpty(range.Text)) continue;
                var replacementText =  _spellChecker.CheckSpell(range.Text);
                if (!string.IsNullOrEmpty(replacementText)) range.Text = replacementText;
            }
        }

        private async void license_Click(object sender, RibbonControlEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.CheckFileExists = true;
            dialog.AddExtension = true;
            dialog.Filter = "License Files|*.lic";
            var result=dialog.ShowDialog();
            if(result== DialogResult.OK)
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
                            var isOnlineValid = await LicenseManager.ValidateLicenseOnline(license.Id);
                            if (isOnlineValid)
                            {
                                await LicenseManager.UpdateLicenseUtilizedCount(license.Id);
                                await LicenseManager.SetLicense(fileStream);
                                MessageBox.Show("Rregjistrimi u krye me sukses. Ju lutemi mbylleni applikacionin qe te shikoni ndryshimet", "Sukses");
                            }
                            else
                            {
                                MessageBox.Show("Licensa ne fjale ka arritur maksimumin e perdorimit", "Licensa nuk eshte e vlefshme");
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
    }
}

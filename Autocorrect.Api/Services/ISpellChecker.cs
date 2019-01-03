using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Autocorrect.Api.Services
{
    public interface ISpellChecker
    {
        Task<string> Check(string input);
    }
    public class SpellChecker
    {
        public SpellChecker()
        {
        }
        public async Task<string> CheckSpell(string input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;

            if (DataProvider.Data.ContainsKey(input))
            {
                var correctValue = DataProvider.Data[input];
                return ReplaceAlCharacters(input, correctValue);
            }
            return string.Empty;
        }

        /// <summary>
        /// Replaces the characters of the original value with their respective al characters keeping track of the original case of the character
        /// </summary>
        /// <param name="input"></param>
        /// <param name="output"></param>
        /// <returns></returns>
        public string ReplaceAlCharacters(string input,string output)
        {
            char[] outputArray = input.ToCharArray();
            for (var i = 0; i < input.Length; i++)
            {
              if(char.ToUpperInvariant(output[i]) != char.ToUpperInvariant(input[i]))
                {
                    var isUpper = char.IsUpper(input[i]);
                    outputArray[i] = isUpper ? char.ToUpperInvariant(output[i]) : char.ToLowerInvariant(output[i]);
                }
                else
                {
                    outputArray[i] = input[i];
                }
            }
           return new string(outputArray);
        }
    }
}

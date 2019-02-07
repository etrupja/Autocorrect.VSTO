using System;
using System.Collections.Generic;
using System.Linq;
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
        public  string CheckSpell(string input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;

            return ProcessString(input);
        }

        public string ProcessString(string input)
        {
           if(input.Contains("’") && !input.EndsWith("’"))
            {
                var parts = input.Split('’');
                if (parts.Count() == 2)
                {
                    return HandleApostrophe(parts[0][0], parts[1]);
                }
            }

            return DataProvider.Data.ContainsKey(input) ? ReplaceKeepCase(input, DataProvider.Data[input]) : string.Empty;
        }
        public string HandleApostrophe(char part1, string part2)
        {
            var part2Value = CheckSpellInternal(part2);
            part2Value= part2Value.Insert(0, "'");
            part2Value = part2Value.Insert(0, HandleApostrophePrepender(part1).ToString());
            return part2Value;
        }
        public string CheckSpellInternal(string input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;

            return DataProvider.Data.ContainsKey(input) ? ReplaceKeepCase(input, DataProvider.Data[input]) : input;
        }

        public char HandleApostrophePrepender(char value)
        {
            char result;
            var isUpperCase = char.IsUpper(value);
            char.ToLowerInvariant(value);
            switch (value)
            {
                case 'c':
                    result= 'ç';
                    break;
                default:
                    result = value;
                    break;
                 
            }
            return isUpperCase ? char.ToUpperInvariant(result) : char.ToLowerInvariant(result);
        }
        public string ReplaceKeepCase(string input, string output)
        {
            if (input.Length == output.Length)
            {
               return ReplaceAlCharacters(input, output);
            }
            var outputArray = output.ToCharArray();
            var isUpperCase= char.IsUpper(input[0]);
            if (isUpperCase) outputArray[0] = char.ToUpperInvariant(output[0]);
            return string.Join("", outputArray);   
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

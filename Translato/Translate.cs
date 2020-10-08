using Translato.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Translato
{
    class Translate
    {
        public Translate() { }

        public string DeleteSymbols(out string[] symbolArr,out List<string> charArr,string str)
        {
            String[] words = str.Split(new char[] { ' ', '.', ',', '`', '"', '\\', '/', '(', ')', '+', '*', '?'}); 
            symbolArr = new string[words.Length];
            charArr = new List<string>();
            string strTest = "";
            int i = 0;

            foreach (var item in words)
            {
                int number;
                bool isNum = int.TryParse(item, out number);

                if (isNum)
                {
                    symbolArr[i] = item;
                    i++;
                }
                else
                {
                    strTest += item;
                    charArr.Add(item);
                }
            }
            return strTest;
        }
    }
}

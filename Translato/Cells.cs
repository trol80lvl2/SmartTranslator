using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Translato
{
    class Cells 
    {
        public byte[] CellInt=new byte[2];
        public byte[] CellToInt(string Cell)
        {
            char[] CharArr = Cell.ToCharArray();
            switch (CharArr[0])
            {
                case 'A':
                    CellInt[1] = 1;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'B':
                    CellInt[1] = 2;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'C':
                    CellInt[1] = 3;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'D':
                    CellInt[1] = 4;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'E':
                    CellInt[1] = 5;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'F':
                    CellInt[1] = 6;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'G':
                    CellInt[1] = 7;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'H':
                    CellInt[1] = 8;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'I':
                    CellInt[1] = 9;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'J':
                    CellInt[1] = 10;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'K':
                    CellInt[1] = 11;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'L':
                    CellInt[1] = 12;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'M':
                    CellInt[1] = 13;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'N':
                    CellInt[1] = 14;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'O':
                    CellInt[1] = 15;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'P':
                    CellInt[1] = 16;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'Q':
                    CellInt[1] = 17;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'R':
                    CellInt[1] = 18;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'S':
                    CellInt[1] = 19;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'T':
                    CellInt[1] = 20;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'U':
                    CellInt[1] = 21;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'V':
                    CellInt[1] = 22;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'W':
                    CellInt[1] = 23;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'X':
                    CellInt[1] = 24;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'Y':
                    CellInt[1] = 25;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
                case 'Z':
                    CellInt[1] = 26;
                    CellInt[0] = Convert.ToByte(Char.GetNumericValue(CharArr[1]));
                    break;
            }
            return CellInt;
        }
    }
}

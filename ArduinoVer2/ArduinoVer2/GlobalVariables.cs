using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pololu.Usc.ScopeFocus
{
    class GlobalVariables
    {
        private static string _nebpath = "";
      //  static int _serverhwnd;

        public static string NebPath
        {
            get {return _nebpath;}
            set { _nebpath = value;}           
        }

        private static int _nebSlavehwnd;
        public static int NebSlavehwnd
        {
            get { return _nebSlavehwnd; }
            set { _nebSlavehwnd = value; }
        }
        private static string _nebcamera;
        public static string Nebcamera
        {
            get { return _nebcamera; }
            set { _nebcamera = value; }
        }
        private static string _path2;
        public static string Path2
        {
            get { return _path2; }
            set { _path2 = value; }
        }
        private static string _portselected;
        public static string Portselected
        {
            get { return _portselected; }
            set { _portselected = value; }
        }
        private static string _portselected2;
        public static string Portselected2
        {
            get { return _portselected2; }
            set { _portselected2 = value; }
        }
        private static int _fineRepeatDone;
        public static int FineRepeatDone
        {
            get { return _fineRepeatDone; }
            set { _fineRepeatDone = value; }
        }
        private static bool _fineRepeatOn= false;
        public static bool FineRepeatOn
        {
            get { return _fineRepeatOn; }
            set { _fineRepeatOn = value; }
        }
        private static int _fineVRepeat;
        public static int FineVRepeat
        {
            get { return _fineVRepeat; }
            set { _fineVRepeat = value; }
        }
        private static string _equipPrefix;
        public static string EquipPrefix
        {
            get { return _equipPrefix; }
            set { _equipPrefix = value; }
        }
        private static string _solveImage;
        public static string SolveImage
        {
            get { return _solveImage; }
            set { _solveImage = value; }
        }
        private static bool _localPlateSolve;
        public static bool LocalPlateSolve
        {
            get { return _localPlateSolve; }
            set { _localPlateSolve = value; }
        }
        //public static int Serverhwnd
        //{
        //    get { return _serverhwnd; }
        //    set { _serverhwnd = value; }
        //}
    }
}

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
        private static bool _serverEnabled;
        public static bool ServerEnabled
        {
            get { return _serverEnabled; }
            set { _serverEnabled = value; }
        }
        private static bool _slaveModeEnabled;
        public static bool SlaveModeEnabled
        {
            get { return _slaveModeEnabled; }
            set { _slaveModeEnabled = value; }
        }
        private static int _capCurrent; // 11-8-16 added next 2
        public static int CapCurrent
        {
            get { return _capCurrent; }
            set { _capCurrent = value; }
        }
        private static int _capTotal;
        public static int CapTotal
        {
            get { return _capTotal; }
            set { _capTotal = value; }
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
        private static string _targetImage;
        public static string TargetImage
        {
            get { return _targetImage; }
            set { _targetImage = value; }
        }
        private static string _focusImage;
        public static string FocusImage
        {
            get { return _focusImage; }
            set { _focusImage = value; }
        }

        private static bool _localPlateSolve;
        public static bool LocalPlateSolve
        {
            get { return _localPlateSolve; }
            set { _localPlateSolve = value; }
        }
        private static Int32 _corrFileLines;
        public static Int32 CorrFileLines
        {
            get { return _corrFileLines; }
            set { _corrFileLines = value; }
        }
        private static bool ffenable = false;
        public static bool FFenable
        {
            get { return ffenable; }
            set { ffenable = value; }
        }
        //private static bool tempon = false; // remd 1-10-17
        //public static bool Tempon
        //{
        //    get { return tempon; }
        //    set { tempon = value; }
        //}
        //public static int Serverhwnd
        //{
        //    get { return _serverhwnd; }
        //    set { _serverhwnd = value; }
        //}
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pololu.Usc.ScopeFocus
{
    class time
    {
        public static DateTime CalcGSTFromUT(DateTime dDate)
        {
            double fJD;
            double fS;
            double fT;
            double fT0;
            DateTime dGST;
            double fUT;
            double fGST;

            fJD = GetJulianDay(dDate.Date, 0);
            fS = fJD - 2451545.0;
            fT = fS / 36525.0;
            fT0 = 6.697374558 + (2400.051336 * fT) + (0.000025862 * fT * fT);
            fT0 = PutIn24Hour(fT0);
            fUT = ConvTimeToDec(dDate);
            fUT = fUT * 1.002737909;
            fGST = fUT + fT0;
            while (fGST > 24)
            {
                fGST = fGST - 24;
            }
            dGST = ConvDecTUraniaTime(fGST);
            return dGST;
        }

        public static DateTime ConvDecTUraniaTime(double fTime)
        {
            DateTime dDate;

            dDate = new DateTime();
            dDate = dDate.AddHours(fTime);
            return (dDate);
        }

        public static double ConvTimeToDec(DateTime dDate)
        {
            double fHour;

            fHour = dDate.Hour + (dDate.Minute / 60.0) + (dDate.Second / 60.0 / 60.0) + (dDate.Millisecond / 60.0 / 60.0 / 1000.0);
            return fHour;
        }
        public static double GetJulianDay(DateTime dDate, int iZone)
        {
            double fJD;
            double iYear;
            double iMonth;
            double iDay;
            double iHour;
            double iMinute;
            double iSecond;
            double iGreg;
            double fA;
            double fB;
            double fC;
            double fD;
            double fFrac;

            dDate = CalcUTFromZT(dDate, iZone);

            iYear = dDate.Year;
            iMonth = dDate.Month;
            iDay = dDate.Day;
            iHour = dDate.Hour;
            iMinute = dDate.Minute;
            iSecond = dDate.Second;
            fFrac = iDay + ((iHour + (iMinute / 60) + (iSecond / 60 / 60)) / 24);
            if (iYear < 1582)
            {
                iGreg = 0;
            }
            else
            {
                iGreg = 1;
            }
            if ((iMonth == 1) || (iMonth == 2))
            {
                iYear = iYear - 1;
                iMonth = iMonth + 12;
            }

            fA = (long)Math.Floor(iYear / 100);
            fB = (2 - fA + (long)Math.Floor(fA / 4)) * iGreg;
            if (iYear < 0)
            {
                fC = (int)Math.Floor((365.25 * iYear) - 0.75);
            }
            else
            {
                fC = (int)Math.Floor(365.25 * iYear);
            }
            fD = (int)Math.Floor(30.6001 * (iMonth + 1));
            fJD = fB + fC + fD + 1720994.5;
            fJD = fJD + fFrac;
            return fJD;
        }

        public static DateTime getDateFromJD(double fJD, int iZone)
        {
            DateTime dDate;
            int iYear;
            int fMonth;
            int iDay;
            int iHour;
            int iMinute;
            int iSecond;
            double fFrac;
            double fFracDay;
            int fI;
            int fA;
            int fB;
            int fC;
            int fD;
            int fE;
            int fG;

            fJD = fJD + 0.5;
            fI = (int)Math.Floor(fJD);
            fFrac = fJD - fI;

            if (fI > 2299160)
            {
                fA = (int)Math.Floor((fI - 1867216.25) / 36524.25);
                fB = fI + 1 + fA - (int)Math.Floor((double)(fA / 4));
            }
            else
            {
                fA = 0;
                fB = fI;
            }
            fC = fB + 1524;
            fD = (int)Math.Floor((fC - 122.1) / 365.25);
            fE = (int)Math.Floor(365.25 * fD);
            fG = (int)Math.Floor((fC - fE) / 30.6001);
            fFracDay = fC - fE + fFrac - (long)Math.Floor((double)(30.6001 * fG));
            iDay = (int)Math.Floor(fFracDay);
            fFracDay = (fFracDay - iDay) * 24;
            iHour = (int)Math.Floor(fFracDay);
            fFracDay = (fFracDay - iHour) * 60;
            iMinute = (int)Math.Floor(fFracDay);
            fFracDay = (fFracDay - iMinute) * 60;
            iSecond = (int)Math.Floor(fFracDay);

            if (fG < 13.5)
            {
                fMonth = fG - 1;
            }
            else
            {
                fMonth = fG - 13;
            }

            if (fMonth > 2.5)
            {
                iYear = (int)Math.Floor((double)(fD - 4716.0));
            }
            else
            {
                iYear = (int)Math.Floor((double)(fD - 4715.0));
            }


            dDate = new DateTime(iYear, (int)Math.Floor((double)fMonth), iDay, iHour, iMinute, iSecond);
            dDate = CalcZTFromUT(dDate, iZone);
            return dDate;
        }
        public static DateTime CalcUTFromGST(DateTime dGSTDate, DateTime dCalDate)
        {
            double fJD;
            double fS;
            double fT;
            double fT0;
            double fGST;
            double fUT;
            DateTime dUT;

            bool bPrevDay;
            bool bNextDay;

            bPrevDay = false;
            bNextDay = false;

            fJD = GetJulianDay(dCalDate.Date, 0);
            fS = fJD - 2451545.0;
            fT = fS / 36525.0;
            fT0 = 6.697374558 + (2400.051336 * fT) + (0.000025862 * fT * fT);

            fT0 = PutIn24Hour(fT0);

            fGST = ConvTimeToDec(dGSTDate);
            fGST = fGST - fT0;

            while (fGST > 24)
            {
                fGST = fGST - 24;
                bNextDay = true;
            }
            while (fGST < 0)
            {
                fGST = fGST + 24;
                bPrevDay = true;
            }

            fUT = fGST * 0.9972695663;
            // fUT = fGST

            dUT = dCalDate.Date;
            dUT = dUT.AddHours(fUT);
            fUT = dUT.Millisecond;

            if (bNextDay == true)
            {
                dUT = dUT.AddDays(1);
            }

            if (bPrevDay == true)
            {
                dUT = dUT.Subtract(new TimeSpan(1, 0, 0, 0, 0));
            }
            return dUT;
        }
        public static DateTime CalcUTFromZT(DateTime dDate, int iZone)
        {
            if (iZone >= 0)
            {
                return dDate.Subtract(new TimeSpan(iZone, 0, 0));
            }
            else
            {
                return dDate.AddHours(Math.Abs(iZone));
            }
        }

        public static DateTime CalcZTFromUT(DateTime dDate, int iZone)
        {
            if (iZone >= 0)
            {
                return dDate.AddHours(iZone);
            }
            else
            {
                return dDate.Subtract(new TimeSpan(Math.Abs(iZone), 0, 0));
            }
        }
        public static DateTime CalcLMTFromUT(DateTime dDate, double fLong)
        {
            bool bAdd = false;
            DateTime dLongDate;

            dLongDate = ConvLongTUraniaTime(fLong, ref bAdd);

            if (bAdd == true)
            {
                dDate = dDate.Add(dLongDate.TimeOfDay);
            }
            else
            {
                dDate = dDate.Subtract(dLongDate.TimeOfDay);
            }

            return dDate;
        }

        public static DateTime ConvLongTUraniaTime(double fLong, ref bool bAdd)
        {
            //double fHours;
            double fMinutes;
            //double fSeconds;
            DateTime dDate;
            //DateTime dTmpDate;

            fMinutes = fLong * 4;
            if (fMinutes < 0)
            {
                bAdd = false;
            }
            else
            {
                bAdd = true;
            }
            fMinutes = Math.Abs(fMinutes);

            dDate = new DateTime();
            dDate = dDate.AddMinutes(fMinutes);
            return dDate;
        }
        public static double PutIn24Hour(double pfHour)
        {
            while (pfHour >= 24)
            {
                pfHour = pfHour - 24;
            }
            while (pfHour < 0)
            {
                pfHour = pfHour + 24;
            }
            return pfHour;
        }
    }
}

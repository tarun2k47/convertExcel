using Focus.Common.DataStructs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace convertExcel.Models
{
    public class FMYDateTime
    {
        public int GetToday()
        {
            try
            {
                return (new Date(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, CalendarType.Gregorean).Value);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public Date GetTodayDate()
        {
            try
            {
                return (new Date(CalendarType.Gregorean));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public long GetTodayDateTime()
        {
            try
            {
                return (new FDateTime(DateTime.Now, CalendarType.Gregorean).Value);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public long DateTimeToInt(DateTime dt)
        {
            try
            {
                return (new FDateTime(dt, CalendarType.Gregorean).Value);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public Date StringToDate(String sDate)
        {
            int iDay = 0;
            int iMonth = 0;
            int iYear = 0;

            try
            {
                ExtractDayMonthYear(sDate, ref iDay, ref iMonth, ref iYear);

                return (new Date(iYear, iMonth, iDay, CalendarType.Gregorean));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public Date IntToDate(int iDate)
        {
            try
            {
                return (new Date(iDate, CalendarType.Gregorean));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public int DateToInt(DateTime dt)
        {
            int iValue = 0;

            try
            {
                iValue = new Date(dt.Year, dt.Month, dt.Day, CalendarType.Gregorean).Value;
                return (iValue);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public int DateToInt(Date dt)
        {
            try
            {
                return (dt.Value);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public int StringToIntDate(String sDate)
        {
            try
            {
                return (StringToDate(sDate).Value);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static bool ExtractDayMonthYear(String sDate, ref int iDay, ref int iMonth, ref int iYear)
        {
            bool bValue = false;

            bValue = ExtractDayMonthYear(sDate, ref iDay, ref iMonth, ref iYear, '/');
            if (bValue == false)
            {
                bValue = ExtractDayMonthYear(sDate, ref iDay, ref iMonth, ref iYear, '-');
            }

            return (bValue);
        }

        public static bool ExtractDayMonthYear(String sDate, ref int iDay, ref int iMonth, ref int iYear, char cSeparator)
        {
            String[] sArr = sDate.Split(cSeparator);

            if (sArr.Length != 3)
            {
                return (false);
            }


            iDay = Convert.ToInt32(sArr[0]);
            iMonth = Convert.ToInt32(sArr[1]);
            iYear = Convert.ToInt32(sArr[2]);

            if (IsValidMonth(iMonth) == false)
            {
                return (false);
            }

            if (iYear < 1 || iYear > 2100)
            {
                return (false);
            }

            if (iDay < 1 || iDay > GetMaxDayOfMonth(iMonth, iYear))
            {
                return (false);
            }

            return (true);
        }

        public static bool IsValidMonth(int iMonth)
        {
            if (iMonth < 1 || iMonth > 12)
            {
                return (false);
            }

            return (true);
        }

        public static int GetMaxDayOfMonth(int iMonth, int iYear)
        {
            int iMaxDay = 0;

            switch (iMonth)
            {
                case 1:
                case 3:
                case 5:
                case 7:
                case 8:
                case 10:
                case 12:
                    iMaxDay = 31;
                    break;
                case 4:
                case 6:
                case 9:
                case 11:
                    iMaxDay = 30;
                    break;
                case 2:
                    iMaxDay = IsLeapYear(iYear) == true ? 29 : 28;
                    break;
            }

            return (iMaxDay);
        }

        public static bool IsLeapYear(int iYear)
        {
            if ((iYear % 400) == 0)
            {
                return (true);
            }
            else if ((iYear % 4) == 0 && (iYear % 100) != 0)
            {
                return (true);
            }

            return (false);
        }

        public int TimeToInt(string time)
        {

            int H = 0;
            int m = 0;
            int S = 0;
            //long lngTemp = 0;
            int iTemp = 0;
            string[] parts = time.Split(':');

            H = Convert.ToInt32(parts[0]);
            m = Convert.ToInt32(parts[1]);
            S = Convert.ToInt32(parts[2]);

            H = H * 256 * 256;
            m = m * 256;
            iTemp = (H + m + S);
            return iTemp;
        }

        public DateTime IntToDateTime(long iDate)
        {
            try
            {
                return (new FDateTime(iDate, CalendarType.Gregorean).DateTime);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DateTime GetDatetime(long val)
        {
            return new Focus.Common.DataStructs.FDateTime(val, CalendarType.Gregorean).DateTime;
        }
    }

}
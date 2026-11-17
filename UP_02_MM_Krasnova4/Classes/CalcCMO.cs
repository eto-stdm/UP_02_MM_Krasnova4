using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UP_02_MM_Krasnova4.Classes
{
    internal class CalcCMO
    {
        public static double Factorial(int n)
        {
            double factorial = 1;
            for (int i = 1; i <= n; i++)
                factorial *= i;

            return factorial;
        }

        public static string CalcSingleReject(double lambda, double mu)
        {
            double q = mu / (lambda + mu);
            double potk = lambda / (lambda + mu);
            double a = (lambda * mu) / (lambda + mu);

            return $"Q = {Math.Round(q, 5)}\n" +
                $"Pотк = {Math.Round(potk, 5)}\n" +
                $"А = {Math.Round(a, 5)}";
        }

        public static string CalcSingleLimitedQueue(double lambda, double mu, double m)
        {
            double ro = lambda / mu;

            double p0 = (1 - ro) / (1 - Math.Pow(ro, m + 2));
            double rotc = Math.Pow(ro, m + 1) * p0;
            double q = 1 - rotc;
            double a = lambda * q;
            double loch = Math.Pow(ro, 2) * (1.0 - Math.Pow(ro, m) * (m + 1.0 - m * ro)) / ((1.0 - Math.Pow(ro, m + 2)) * (1.0 - ro));
            double lob = 1 - p0;
            double lsist = loch + lob;
            double tsist = (1 * lsist) / lambda;
            double toch = (1 * loch) / lambda;

            return $"Q = {Math.Round(q, 5)}\n" +
                $"Pотк = {Math.Round(rotc, 5)}\n" +
                $"A = {Math.Round(a, 5)}\n" +
                $"ρ = {Math.Round(ro, 5)}\n" +
                $"P0 = {Math.Round(p0, 5)}\n" +
                $"Lсист = {Math.Round(lsist, 5)}\n" +
                $"Lоч = {Math.Round(loch, 5)}\n" +
                $"Lоб = {Math.Round(lob, 5)}\n" +
                $"Тсист = {Math.Round(tsist, 5)}\n" +
                $"Точ = {Math.Round(toch, 5)}\n";
        }

        public static string CalcSingleInfiniteQueue(double lambda, double mu)
        {
            double ro = lambda / mu;
            if (ro >= 1)
            {
                return "Предельная вероятность не существует";
            }

            double p0 = 1.0 - ro;
            double lsist = ro / (p0);
            double loch = Math.Pow(ro, 2) / (p0);
            double tsist = (1 * lsist) / lambda;
            double toch = (1 * loch) / lambda;

            return $"Q = 1\n" +
                $"Pотк = 0\n" +
                $"A = λ\n" +
                $"ρ = {Math.Round(ro, 5)}\n" +
                $"P0 = {Math.Round(p0, 5)}\n" +
                $"Lсист = {Math.Round(lsist, 5)}\n" +
                $"Lоч = {Math.Round(loch, 5)}\n" +
                $"Тсист = {Math.Round(tsist, 5)}\n" +
                $"Точ = {Math.Round(toch, 5)}\n";
        }

        public static string CalcMultiReject(double lambda, double mu, int n)
        {
            double ro = lambda / mu;

            double sum = 0;
            for (int i = 0; i <= n; i++)
            {
                sum += Math.Pow(ro, i) / Factorial(i);
            }
            double p0 = Math.Pow(sum, -1); // или 1.0 / sum

            double potk = (Math.Pow(ro, n) / Factorial(n)) * p0;

            double q = 1 - potk;
            double a = lambda * q;
            double k = a / mu;

            return
                $"ρ = {Math.Round(ro, 5)}\n" +
                $"P0 = {Math.Round(p0, 5)}\n" +
                $"Pотк = {Math.Round(potk, 5)}\n" +
                $"Q = {Math.Round(q, 5)}\n" +
                $"A = {Math.Round(a, 5)}\n" +
                $"k = {Math.Round(k, 5)}\n";
        }

        public static string CalcMultiLimitedQueue(double lambda, double mu, int n, double m)
        {
            double ro = lambda / mu;

            double sum = 1.0;
            for (int i = 1; i <= n; i++)
                sum += Math.Pow(ro, i) / Factorial(i);

            double p0 = Math.Pow(sum + (Math.Pow(ro, n + 1) * (1.0 - Math.Pow(ro / n, m))) / (n * Factorial(n) * (1.0 - ro / n)), -1);

            double rotc = (Math.Pow(ro, m + 1) / (Math.Pow(n, m) * Factorial(n))) * p0;
            double q = 1.0 - rotc;
            double a = lambda * q;
            double loch = (Math.Pow(ro, n + 1) * p0 * (1.0 - (m + 1.0 - m * (ro / n)) * Math.Pow(ro / n, m))) / (n * Factorial(n) * (1.0 - ro / n));
            double k = ro * (1.0 - ((Math.Pow(ro, n + m) / (Math.Pow(n, m) * Factorial(n))) * p0));
            double lsist = loch + k;
            double tsist = (1.0 * lsist) / lambda;
            double toch = (1.0 * loch) / lambda;

            return $"Q = {Math.Round(q, 5)}\n" +
                $"Pотк = {Math.Round(rotc, 5)}\n" +
                $"A = {Math.Round(a, 5)}\n" +
                $"ρ = {Math.Round(ro, 5)}\n" +
                $"P0 = {Math.Round(p0, 5)}\n" +
                $"Lсист = {Math.Round(lsist, 5)}\n" +
                $"Lоч = {Math.Round(loch, 5)}\n" +
                $"k = {Math.Round(k, 5)}\n" +
                $"Тсист = {Math.Round(tsist, 5)}\n" +
                $"Точ = {Math.Round(toch, 5)}\n";
        }

        public static string CalcMultiInfiniteQueue(double lambda, double mu, int n, double mAmount)
        {
            //if (mAmount != 0) { } формул с m не нашла

            double ro = lambda / mu;
            if (ro / n >= 1)
            {
                return "Предельная вероятность не существует";
            }

            double sum = 0;
            for (int i = 0; i <= n; i++)
            {
                sum += Math.Pow(ro, i) / Factorial(i);
            }
            sum += Math.Pow(ro, n + 1) / (Factorial(n) * (n - ro));
            double p0 = Math.Pow(sum, -1);

            double loch = (Math.Pow(ro, n + 1) / (n * Factorial(n) * Math.Pow(1.0 - (ro / n), 2))) * p0;
            double lsist = loch + ro;
            double tsist = (1 * lsist) / lambda;
            double toch = (1 * loch) / lambda;
            double poch = (Math.Pow(ro, n + 1) / (Factorial(n) * (n - ro))) * p0;

            return $"Q = 1\n" +
                $"Pотк = 0\n" +
                $"A = λ\n" +
                $"ρ = {Math.Round(ro, 5)}\n" +
                $"P0 = {Math.Round(p0, 5)}\n" +
                $"Lсист = {Math.Round(lsist, 5)}\n" +
                $"Lоч = {Math.Round(loch, 5)}\n" +
                $"Тсист = {Math.Round(tsist, 5)}\n" +
                $"Точ = {Math.Round(toch, 5)}\n" +
                $"Pоч = {Math.Round(poch, 5)}\n";
        }
    }
}

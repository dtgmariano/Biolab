/* File: Biolab_Statistics.cs
 * Author: Daniel Teodoro Gonçalves Mariano
 * Email: dtgmariano@gmail.com
 * Date: 2015/03/05
 * 
 * Comments: CSharp static class to provide statistical methods
 */
namespace Biolab
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    public static class Statistics
    {
        /*Calculates Population Average*/
        public static double Average_Population(List<double> nums)
        {
            return nums.Average();
        }

        /*Calculates Population Average*/
        public static int Average_Population(List<int> nums)
        {
            return Convert.ToInt32(nums.Average());
        }

        /*Calculates Population Standard Deviation*/
        public static double StandardDeviation_Population(List<double> nums)
        {
            double sum = 0.0;
            nums.ForEach(x => sum += Math.Pow((x - nums.Average()), 2));
            return Math.Sqrt(sum / (nums.Count));
        }

        public static double StandardDeviation_Population(List<int> nums)
        {
            double sum = 0.0;
            nums.ForEach(x => sum += Math.Pow((x - nums.Average()), 2));
            return Math.Sqrt(sum / (nums.Count));
        }

        /*Calculates Sample Standard Deviation*/
        public static double StandardDeviation_Sample(List<double> nums)
        {
            double sum = 0.0;
            nums.ForEach(x => sum += Math.Pow((x - nums.Average()), 2));
            return Math.Sqrt(sum / (nums.Count - 1));
        }

        /*Calculates Sample Standard Deviation*/
        public static double StandardDeviation_Sample(List<int> nums)
        {
            double sum = 0.0;
            nums.ForEach(x => sum += Math.Pow((x - nums.Average()), 2));
            return Math.Sqrt(sum / (nums.Count - 1));
        }

        /*Calculates Factorial for an integer i*/
        public static int Factorial(int i)
        {
            if (i <= 1)
                return 1;
            return i * Factorial(i - 1);
        }

    }

    public enum Hyphotesis
    {
        Ho_Is_True,
        Ho_Is_False,
        Ha_Is_True,
        Ha_Is_False
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EnergyAnalysis.Services
{
    public class Smooth
    {
        /// <summary>
        /// 计算平滑常数a
        /// </summary>
        /// <param name="size"></param>
        /// <returns></returns>
        public double a_Computed(int size)
        {
            return 2 / (size + 1.0);
        }

        /// <summary>
        /// 一次指数平滑
        /// </summary>
        /// <param name="array"></param>
        /// <param name="length"></param>
        /// <param name="n"></param>
        /// <returns></returns>
        public List<float> primary_Exponential(List<float> array, int length, int n)
        {
            float s0 = 0;
            //float s1 = 0;
            List<float> smoothed_data = new List<float>();
            for (int i = 0; i < 3; i++)
            {
                s0 = s0 + array[i];
            }
            s0 = s0 / 3;
            float a = (float)a_Computed(n);//平滑常数
            for (int j = 0; j < length; j++)
            {
                if (j == 0)
                {
                    smoothed_data.Add(a * array[j] + (1 - a) * s0);
                    //printf("%f\n", smoothed_data[j]);
                }
                else
                {
                    smoothed_data.Add(a * array[j] + (1 - a) * smoothed_data[j - 1]);
                    //printf("%f\n", smoothed_data[j]);
                }
            }
            return smoothed_data;
        }

        /// <summary>
        /// 二次指数平滑
        /// </summary>
        /// <param name="array"></param>
        /// <param name="length"></param>
        /// <param name="n"></param>
        /// <returns></returns>
        public List<float> quadratic_Exponential(List<float> array, int length, int n)
        {
            List<float> smoothed_data = new List<float>();
            float s1 = 0;
            List<float> p = new List<float>();
            p = primary_Exponential(array, length, n);
            float a = (float)a_Computed(n);
            for (int i = 0; i < 3; i++)
            {
                s1 = s1 + p[i];
            }
            s1 = s1 / 3;

            for (int j = 0; j < length; j++)
            {
                if (j == 0)
                {
                    smoothed_data.Add(a * p[j] + (1 - a) * s1);
                    //printf("%f\n", smoothed_data[j]);
                }
                else
                {
                    smoothed_data.Add(a * p[j] + (1 - a) * smoothed_data[j - 1]);
                    //printf("%f\n", smoothed_data[j]);
                }
            }
            return smoothed_data;
        }

        /// <summary>
        /// 三次指数滤波
        /// </summary>
        /// <param name="array"></param>
        /// <param name="length"></param>
        /// <param name="n"></param>
        /// <returns></returns>
        public List<float> cubic_Exponential(List<float> array, int length, int n)
        {
            float s2 = 0;
            List<float> p = new List<float>();
            p = quadratic_Exponential(array, length, n);
            List<float> smoothed_data = new List<float>();
            float a = (float)a_Computed(n);
            for (int i = 0; i < 3; i++)
            {
                s2 = s2 + p[i];
            }
            s2 = s2 / 3;

            for (int j = 0; j < length; j++)
            {
                if (j == 0)
                {
                    smoothed_data.Add(a * p[j] + (1 - a) * s2);
                    //printf("%f\n", smoothed_data[j]);
                }
                else
                {
                    smoothed_data.Add(a * p[j] + (1 - a) * smoothed_data[j - 1]);
                    //printf("%f\n", smoothed_data[j]);
                }
            }
            return smoothed_data;
        }

        /// <summary>
        /// 平滑滤波
        /// </summary>
        /// <param name="array"></param>
        /// <param name="length"></param>
        /// <param name="size"></param>
        /// <returns></returns>
        public List<float> mobile_Smoothing(List<float> array, int length, int size)
        {
            List<float> smoothed_data = new List<float>();
            if ((length / 2) == 0)
            {
                //检测输入size是否为奇数
                return null;
            }
            else
            {
                if (length >= size)
                {
                    //对每一位进行计算
                    int z = size / 2;
                    for (int i = 0; i < length; i++)
                    {
                        float sum = 0;
                        if (i < z)
                        {
                            //第一类：前z个数据
                            int k = 2 * i + 1;
                            for (int j = 0; j < k; j++)
                            {
                                sum = sum + array[j];
                            }
                            smoothed_data.Add(sum / k);
                        }
                        else if (z <= i && i < length - z)
                        {
                            //第二类：length-z与z+1之间
                            sum = array[i];
                            for (int l = 0; l < z; l++)
                            {
                                sum = sum + array[i + l + 1] + array[i - l - 1];
                            }
                            smoothed_data.Add(sum / size);
                        }
                        else if (length - z <= i && i < length)
                        {
                            //第三类：length-z与length之间
                            for (int j = 0; j <= z; j++)
                            {
                                sum = sum + array[i - j];
                            }

                            int k = length - i - 1;
                            for (int l = 0; l < k; l++)
                            {
                                sum = sum + array[i + l + 1];
                            }
                            smoothed_data.Add(sum / (size - z + k));
                        }
                    }
                    return smoothed_data;
                }
                else
                {
                    return null;
                }
            }
        }
    }
}

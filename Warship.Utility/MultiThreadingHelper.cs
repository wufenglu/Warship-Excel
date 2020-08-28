using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Utility
{
    /// <summary>
    /// 多线程循环执行
    /// </summary>
    public static class MultiThreadingHelper
    {
        /// <summary>
        /// 获取并行选项
        /// </summary>
        /// <returns></returns>
        public static ParallelOptions GetParallelOptions() {
            //这里只使用半数逻辑内核，但是考虑服务器可能有奇数内核情况
            var maxDegree = Environment.ProcessorCount % 2 == 0 ? Environment.ProcessorCount / 2 : (Environment.ProcessorCount / 2) + 1;
            var parallelOptions = new ParallelOptions() { MaxDegreeOfParallelism = maxDegree };
            return parallelOptions;
        }

        /// <summary>
        /// 多线程循环处理
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="sources"></param>
        /// <param name="body"></param>
        public static void ForEach<TSource>(this IEnumerable<TSource> sources, Action<TSource> body)
        {
            if (sources == null) {
                return;
            }
            //这里只使用半数逻辑内核，但是考虑服务器可能有奇数内核情况
            var maxDegree = Environment.ProcessorCount % 2 == 0 ? Environment.ProcessorCount / 2 : (Environment.ProcessorCount / 2) + 1;
            var parallelOptions = new ParallelOptions() { MaxDegreeOfParallelism = maxDegree };
            Parallel.ForEach(sources, parallelOptions, source => body(source));
        }
    }
}

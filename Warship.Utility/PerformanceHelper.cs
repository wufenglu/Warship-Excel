using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Collections.Concurrent;

namespace Warship.Utility
{
    /// <summary>
    /// 性能监控链路
    /// </summary>
    public class PerformanceMonitoringLink
    {
        /// <summary>
        /// 链路层级
        /// </summary>
        public static Dictionary<int, int> LinkLevel = new Dictionary<int, int>();

        /// <summary>
        /// 链路性能对象类型说明：线程ID、层级、性能监控对象
        /// </summary>
        public static ConcurrentDictionary<int, List<PerformanceMonitoring>> LinkPerformanceMonitoring = new ConcurrentDictionary<int, List<PerformanceMonitoring>>();

        #region 链路层级处理

        /// <summary>
        /// 加
        /// </summary>
        public static void LevelPlus()
        {
            int id = Thread.CurrentThread.ManagedThreadId;
            if (LinkLevel.ContainsKey(id) == false)
            {
                LinkLevel.Add(id, 0);
            }
            else
            {
                LinkLevel[id] = LinkLevel[id] + 1;
            }
        }

        /// <summary>
        /// 减
        /// </summary>
        public static void LevelReduce()
        {
            int id = Thread.CurrentThread.ManagedThreadId;
            LinkLevel[id] = LinkLevel[id] - 1;
        }

        /// <summary>
        /// 获取层级
        /// </summary>
        /// <returns></returns>
        public static int GetLevel()
        {
            int id = Thread.CurrentThread.ManagedThreadId;
            return LinkLevel[id];
        }

        /// <summary>
        /// 获取数量
        /// </summary>
        /// <returns></returns>
        public static int GetIndex()
        {
            int id = Thread.CurrentThread.ManagedThreadId;
            if (LinkPerformanceMonitoring.Keys.Contains(id) == false)
            {
                return 0;
            }
            return LinkPerformanceMonitoring[id] == null ? 0 : LinkPerformanceMonitoring[id].Count - 1;
        }

        #endregion

        #region 性能监控

        /// <summary>
        /// 获取性能监控
        /// </summary>
        /// <returns></returns>
        public static int Start(string title)
        {
            PerformanceMonitoring performanceMonitoring = new PerformanceMonitoring();
            performanceMonitoring.Start(title);

            //获取线程性能字典
            int id = Thread.CurrentThread.ManagedThreadId;
            if (LinkPerformanceMonitoring.ContainsKey(id) == false)
            {
                LinkPerformanceMonitoring.TryAdd(id, new List<PerformanceMonitoring>());
            }

            //获取线程监控层级
            int level = GetLevel();
            LinkPerformanceMonitoring[id].Add(performanceMonitoring);

            return GetIndex();
        }

        /// <summary>
        /// 停止
        /// </summary>
        /// <param name="index"></param>
        public static void Stop(int index)
        {
            //获取线程性能字典
            int id = Thread.CurrentThread.ManagedThreadId;
            List<PerformanceMonitoring> list = LinkPerformanceMonitoring[id];

            //获取线程监控层级
            list[index].Stop();

            LinkPerformanceMonitoring[id] = list;
        }

        #endregion

        /// <summary>
        /// 获取链路监控
        /// </summary>
        /// <param name="removeLinkLog">获取链路信息后移除链路日志</param>
        /// <returns></returns>
        public static List<PerformanceDtlEntity> GetLinkPerformanceMonitoring(bool removeLinkLog = true)
        {
            int id = Thread.CurrentThread.ManagedThreadId;
            var result = GetLinkLevel(LinkPerformanceMonitoring[id], 0);

            //移除链路日志
            if (removeLinkLog)
            {
                List<PerformanceMonitoring> outObj;
                LinkPerformanceMonitoring.TryRemove(id, out outObj);
            }
            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="allMonitoring"></param>
        /// <param name="level">层级</param>
        /// <returns></returns>
        public static List<PerformanceDtlEntity> GetLinkLevel(List<PerformanceMonitoring> allMonitoring, int level)
        {
            List<PerformanceDtlEntity> levelDtlList = allMonitoring.Where(w => w.DtlEntity.Level == level).Select(s => s.DtlEntity).ToList();
            foreach (PerformanceDtlEntity levelDtl in levelDtlList)
            {
                //所有的下级
                List<PerformanceMonitoring> allChild = new List<PerformanceMonitoring>();
                bool isExsit = false;
                /* 获取所有下级
                 1、序号>当前项序号，缩小数据范围
                 2、在当前范围内，只要第二次匹配上层级与当前层级一致的，则匹配层级及以后的层级都是其他的，不属于当前层级的子级
                    如：1、2、3、4、5
                        1、2、3、4
                    当从第一个1构建层级，则第一行的2、3、4、5代表这个1的子级，当再次匹配上第二行的1时，则第二行1及以后的数字都代表其他的层级数据
                 */
                foreach (PerformanceMonitoring item in allMonitoring.Where(w => w.DtlEntity.Index >= levelDtl.Index))
                {
                    if (levelDtl.Level == item.DtlEntity.Level && isExsit == true)
                    {
                        break;
                    }
                    if (levelDtl.Level == item.DtlEntity.Level)
                    {
                        isExsit = true;
                        continue;
                    }
                    allChild.Add(item);
                }

                levelDtl.Childs = GetLinkLevel(allChild, levelDtl.Level + 1);
            }
            return levelDtlList;
        }
    }

    /// <summary>
    /// 性能监控类
    /// </summary>
    public class PerformanceMonitoring
    {

        /// <summary>
        /// 秒表对象
        /// </summary>
        public System.Diagnostics.Stopwatch Stopwatch { get; set; }
        /// <summary>
        /// 性能明细列表
        /// </summary>
        public List<PerformanceDtlEntity> EntityList { get; set; }
        /// <summary>
        /// 性能明细
        /// </summary>
        public PerformanceDtlEntity DtlEntity { get; set; }

        /// <summary>
        /// 构造函数
        /// </summary>
        public PerformanceMonitoring()
        {
            //列表
            EntityList = new List<PerformanceDtlEntity>();
        }
        /// <summary>
        /// 开始计算
        /// </summary>
        /// <param name="title">标题</param>
        public void Start(string title)
        {
            PerformanceMonitoringLink.LevelPlus();

            //明细
            DtlEntity = new PerformanceDtlEntity();
            DtlEntity.Title = title;//标题
            DtlEntity.StartTime = DateTime.Now;//开始时间   
            DtlEntity.Level = PerformanceMonitoringLink.GetLevel();
            DtlEntity.Index = PerformanceMonitoringLink.GetIndex();

            //秒表
            Stopwatch = new System.Diagnostics.Stopwatch();
            Stopwatch.Start();//秒表开始
        }
        /// <summary>
        /// 结束计算
        /// </summary>
        public void Stop()
        {
            PerformanceMonitoringLink.LevelReduce();

            Stopwatch.Stop();//秒表结束            

            DtlEntity.StopTime = DateTime.Now;//结束时间
            DtlEntity.ElapsedMilliseconds = Stopwatch.ElapsedMilliseconds;//总运行时间            
            EntityList.Add(DtlEntity);//添加明细
        }
    }
    /// <summary>
    /// 性能明细
    /// </summary>
    public class PerformanceDtlEntity
    {
        /// <summary>
        /// 序号
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// 链路层级
        /// </summary>
        public int Level { get; set; }

        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// 开始时间
        /// </summary>
        public DateTime StartTime { get; set; }
        /// <summary>
        /// 结束时间
        /// </summary>
        public DateTime StopTime { get; set; }
        /// <summary>
        /// 总运行时间（一毫秒为单位）
        /// </summary>
        public long ElapsedMilliseconds { get; set; }

        /// <summary>
        /// 下级
        /// </summary>
        public List<PerformanceDtlEntity> Childs { get; set; }
    }
}

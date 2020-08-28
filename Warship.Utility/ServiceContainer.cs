using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Warship.Utility
{
    /// <summary>
    /// 服务容器
    /// </summary>
    public static class ServiceContainer
    {
        /// <summary>
        /// 容器实体列表
        /// </summary>
        private static List<ContainerEntity> ContainerList = new List<ContainerEntity>();
        /// <summary>
        /// 注册
        /// </summary>
        /// <typeparam name="Service"></typeparam>
        /// <typeparam name="Interface"></typeparam>
        /// <param name="alias"></param>
        public static void Register<Interface, Service>(string alias = null) 
            where Service : class
        {
            Type serviceType = typeof(Service);
            Type interfaceType = typeof(Interface);

            ContainerEntity entity = new ContainerEntity();
            entity.Alias = alias;
            entity.InterfaceAssemblyFullName = interfaceType.FullName;
            entity.ServiceAssembly = Assembly.GetAssembly(serviceType).FullName;
            entity.ServiceAssemblyFullName = serviceType.FullName;

            ContainerList.Add(entity);
        }

        /// <summary>
        /// 获取服务
        /// </summary>
        /// <typeparam name="Interface"></typeparam>
        /// <param name="alias"></param>
        /// <returns></returns>
        public static Interface GetService<Interface>(string alias = null) where Interface : class
        {
            Type interfaceType = typeof(Interface);
            List<ContainerEntity> list = ContainerList.Where(w => w.InterfaceAssemblyFullName == interfaceType.FullName && w.Alias == alias).ToList();
            //为空判断
            if (list.IsNullOrEmpty()) {
                return null;
            }
            return Assembly.Load(list.First().ServiceAssembly).CreateInstance(list.First().ServiceAssemblyFullName) as Interface;
        }

        /// <summary>
        /// 初始化服务：构建属性注入
        /// </summary>
        public static AppService InitService<AppService>()
        {
            AppService service = System.Activator.CreateInstance<AppService>();
            foreach (var prop in service.GetType().GetProperties())
            {
                if (prop.PropertyType.IsInterface)
                {
                    //实例化接口的实现类
                    string interfactFullName = prop.PropertyType.FullName;
                    List<ContainerEntity> list = ContainerList.Where(w => w.InterfaceAssemblyFullName == interfactFullName).ToList();
                    object appService = Assembly.Load(list.First().ServiceAssembly).CreateInstance(list.First().ServiceAssemblyFullName);

                    //初始化属性注入
                    Type serviceContainerType = typeof(ServiceContainer);
                    MethodInfo methodInfo = serviceContainerType.GetMethod("InitService");

                    //设置泛型方法的泛型
                    methodInfo = methodInfo.MakeGenericMethod(appService.GetType());

                    //执行
                    var installObject = methodInfo.Invoke(serviceContainerType, null);

                    //设置返回值
                    prop.SetValue(service, installObject);
                }
            }
            return service;
        }
    }

    /// <summary>
    /// 容器实体
    /// </summary>
    internal class ContainerEntity
    {
        /// <summary>
        /// 别名，用于一个接口多个实现类
        /// </summary>
        public string Alias { get; set; }
        /// <summary>
        /// 接口程序集完整名称
        /// </summary>
        public string InterfaceAssemblyFullName { get; set; }
        /// <summary>
        /// 服务程序集
        /// </summary>
        public string ServiceAssembly { get; set; }
        /// <summary>
        /// 服务程序集完整名称
        /// </summary>
        public string ServiceAssemblyFullName { get; set; }
    }
}

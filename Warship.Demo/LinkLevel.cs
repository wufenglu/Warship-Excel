using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Warship.Utility;
using System.Threading;

namespace Mysoft.Clgyl.Demo
{
    class LinkLevel1
    {
        public void Exec()
        {
            int index2 =  PerformanceMonitoringLink.Start("LinkLevel2");
            LinkLevel2 linkLevel2 = new LinkLevel2();
            linkLevel2.Exec();            
            PerformanceMonitoringLink.Stop(index2);

            int index3 = PerformanceMonitoringLink.Start("LinkLevel3");
            LinkLevel3 linkLevel3 = new LinkLevel3();
            linkLevel3.Exec();
            PerformanceMonitoringLink.Stop(index3);
        }
    }

    class LinkLevel2
    {
        public void Exec()
        {
            Thread.Sleep(1000);

            int index3 = PerformanceMonitoringLink.Start("LinkLevel3");
            LinkLevel3 linkLevel3 = new LinkLevel3();
            linkLevel3.Exec();
            PerformanceMonitoringLink.Stop(index3);

            int index4 = PerformanceMonitoringLink.Start("LinkLevel4");
            LinkLevel4 LinkLevel4 = new LinkLevel4();
            LinkLevel4.Exec();
            PerformanceMonitoringLink.Stop(index4);
        }
    }

    class LinkLevel3
    {
        public void Exec()
        {
            Thread.Sleep(2000);
            int index4 = PerformanceMonitoringLink.Start("LinkLevel4");
            LinkLevel4 linkLevel4 = new LinkLevel4();
            linkLevel4.Exec();
            PerformanceMonitoringLink.Stop(index4);
        }
    }

    class LinkLevel4
    {
        public void Exec()
        {
            Thread.Sleep(800);
            int index5 = PerformanceMonitoringLink.Start("LinkLevel5");
            LinkLevel5 linkLevel5 = new LinkLevel5();
            linkLevel5.Exec();
            PerformanceMonitoringLink.Stop(index5);
        }
    }

    class LinkLevel5
    {
        public void Exec()
        {
            Thread.Sleep(500);
        }
    }
}

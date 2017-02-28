using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Windows;
using Castle.ActiveRecord;
using Castle.ActiveRecord.Framework;

namespace OperateExcel
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : Application
    {


        private void onStartUp(object sender, StartupEventArgs e)
        {
            IConfigurationSource source = new Castle.ActiveRecord.Framework.Config.XmlConfigurationSource("../../TestCases/ActiveRecordConfig.xml");
            ActiveRecordStarter.Initialize(source, typeof(WordRES), typeof(WordLogic));
        }
    }
}

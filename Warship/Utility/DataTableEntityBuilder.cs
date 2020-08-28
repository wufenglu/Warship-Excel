using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Reflection;
using System.Reflection.Emit;
using Warship.Attribute.Model;
using Warship.Attribute;

namespace Warship.Utility
{
    internal class DataTableEntityBuilder<T> where T: class, new()
    {
        public static List<T> GetList(DataTable dt)
        {
            List<T> lst = new List<T>();
            if (dt == null || dt.Rows.Count == 0)
            {
                return lst;
            }

            DataTableEntityBuilder<T> eblist = CreateBuilder(dt.Rows[0]);
            foreach (DataRow dr in dt.Rows)
            {
                lst.Add(eblist.Build(dr));
            }
            dt.Dispose();
            return lst;
        }

        public static T GetEntity(DataRow dr)
        {
            DataTableEntityBuilder<T> eblist = CreateBuilder(dr);
            return eblist.Build(dr);
        }

        private static readonly MethodInfo getValueMethod = typeof(DataRow).GetMethod("get_Item", new Type[] { typeof(int) });
        private static readonly MethodInfo isDBNullMethod = typeof(DataRow).GetMethod("IsNull", new Type[] { typeof(int) });
        private delegate T Load(DataRow dr);

        private Load handler;
        private DataTableEntityBuilder() { }

        public T Build(DataRow dr)
        {
            return handler(dr);
        }

        public static DataTableEntityBuilder<T> CreateBuilder(DataRow dr)
        {
            DataTableEntityBuilder<T> dynamicBuilder = new DataTableEntityBuilder<T>();
            DynamicMethod method = new DynamicMethod("DynamicCreateEntity", typeof(T), new Type[] { typeof(DataRow) }, typeof(T), true);
            ILGenerator generator = method.GetILGenerator();
            LocalBuilder result = generator.DeclareLocal(typeof(T));
            generator.Emit(OpCodes.Newobj, typeof(T).GetConstructor(Type.EmptyTypes));
            generator.Emit(OpCodes.Stloc, result);
            var attributeList = ExcelAttributeHelper<T>.GetHeads();
            for (int i = 0; i < dr.ItemArray.Length; i++)
            {
                var list = attributeList.Where(w => w.PropertyName.ToLower() == dr.Table.Columns[i].ColumnName.ToLower());
                 if (list.Count() > 0)
                 {
                     PropertyInfo pi = typeof(T).GetProperty(list.First().PropertyName);
                     Label endIfLabel = generator.DefineLabel();
                     if (pi != null && pi.GetSetMethod() != null)
                     {
                         generator.Emit(OpCodes.Ldarg_0);
                         generator.Emit(OpCodes.Ldc_I4, i);
                         generator.Emit(OpCodes.Callvirt, isDBNullMethod);
                         generator.Emit(OpCodes.Brtrue, endIfLabel);
                         generator.Emit(OpCodes.Ldloc, result);
                         generator.Emit(OpCodes.Ldarg_0);
                         generator.Emit(OpCodes.Ldc_I4, i);
                         generator.Emit(OpCodes.Callvirt, getValueMethod);
                         generator.Emit(OpCodes.Unbox_Any, pi.PropertyType);
                         generator.Emit(OpCodes.Callvirt, pi.GetSetMethod());
                         generator.MarkLabel(endIfLabel);
                     }
                 }
            }
            generator.Emit(OpCodes.Ldloc, result);
            generator.Emit(OpCodes.Ret);
            dynamicBuilder.handler = (Load)method.CreateDelegate(typeof(Load));
            return dynamicBuilder;
        }
    }
}

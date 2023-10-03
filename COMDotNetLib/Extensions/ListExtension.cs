// https://stackoverflow.com/questions/9860387/how-do-i-create-a-dynamic-type-listt

using GCollections = global::System.Collections;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DotNetLib.Extensions
{
    public static class ListExtension
    {
        static List<T> CreateFromType<T>(T obj)
        {
            return new List<T>();
        }

        ////public static List<T> CreateFromType<T>(this T value)
        ////{
        ////    return new[] { value }.ToList();
        ////}

        //public static GCollections.IList CreateFromType<T>(T obj)
        //{
        //    Type type = obj.GetType();
        //    Type listType = typeof(List<>).MakeGenericType(new[] { type });
        //    GCollections.IList list = (GCollections.IList)Activator.CreateInstance(listType);
        //    return list;
        //}

    }
}

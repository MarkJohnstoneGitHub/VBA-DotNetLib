using System;
using System.Linq.Expressions;
using System.Reflection;

namespace DotNetLib.Extensions
{
    // https://stackoverflow.com/questions/9382216/get-methodinfo-from-a-method-reference-c-sharp
    // https://stackoverflow.com/a/56476316/10759363
    public static class GetMethodInfoExtension
    {

        // Direct helpers
        // Given the above remarks, we can provide the following set of helper methods to alleviate somewhat
        // the tedious task of casting every method to reach its informations:

        // No cast necessary
        public static MethodInfo GetMethodInfo(Action action) => action.Method;
        public static MethodInfo GetMethodInfo<T>(Action<T> action) => action.Method;
        public static MethodInfo GetMethodInfo<T, U>(Action<T, U> action) => action.Method;
        public static MethodInfo GetMethodInfo<TResult>(Func<TResult> fun) => fun.Method;
        public static MethodInfo GetMethodInfo<T, TResult>(Func<T, TResult> fun) => fun.Method;
        public static MethodInfo GetMethodInfo<T, U, TResult>(Func<T, U, TResult> fun) => fun.Method;

        // Cast necessary
        public static MethodInfo GetMethodInfo(Delegate del) => del.Method;

        //Indirect helpers
        // Now for some corner cases the above helpers won't work. Say the method uses out parameters for example.
        // In those special cases, extracting method informations from lambda expressions becomes handy, and we get
        // back to the solution provided by other posters (code inspiration from here):

        // Get MethodInfo from Lambda expressions
        public static MethodInfo GetIndirectMethodInfo(Expression<Action> expression)
            => GetIndirectMethodInfo((LambdaExpression)expression);
        public static MethodInfo GetIndirectMethodInfo<T>(Expression<Action<T>> expression)
            => GetIndirectMethodInfo((LambdaExpression)expression);
        public static MethodInfo GetIndirectMethodInfo<T, TResult>(Expression<Func<TResult>> expression)
            => GetIndirectMethodInfo((LambdaExpression)expression);
        public static MethodInfo GetIndirectMethodInfo<T, TResult>(Expression<Func<T, TResult>> expression)
            => GetIndirectMethodInfo((LambdaExpression)expression);

        // Used by the above
        private static MethodInfo GetIndirectMethodInfo(LambdaExpression expression)
        {
            if (!(expression.Body is MethodCallExpression methodCall))
            {
                throw new ArgumentException(
                    $"Invalid Expression ({expression.Body}). Expression should consist of a method call only.");
            }
            return methodCall.Method;
        }

    }
}

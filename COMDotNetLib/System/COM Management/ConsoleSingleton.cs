// https://stackoverflow.com/questions/4362111/how-do-i-show-a-console-output-window-in-a-forms-application
// https://saezndaree.wordpress.com/2009/03/29/how-to-redirect-the-consoles-output-to-a-textbox-in-c/
// https://github.com/kellyethridge/VBCorLib/blob/master/Source/CorLib/System/Console.cls
// https://stackoverflow.com/questions/15604014/no-console-output-when-using-allocconsole-and-target-architecture-x86
// https://developercommunity.visualstudio.com/t/console-output-is-gone-in-vs2017-works-fine-when-d/12166
// https://learn.microsoft.com/en-us/dotnet/api/system.diagnostics.processstartinfo?view=netframework-4.8
// https://github.com/Tyrrrz/CliWrap
// https://www.youtube.com/watch?v=Pt-0KM5SxmI&ab_channel=NickChapsas

using DotNetLib.Extensions;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Description("Represents the standard input, output, and error streams for console applications.")]
    [Guid("88AEC64B-B9AF-4360-A654-12CD4EA11BD6")]
    [ProgId("DotNetLib.System.Console")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IConsoleSingleton))]
    public class ConsoleSingleton : IConsoleSingleton
    {
        public ConsoleSingleton() { }

        public void Clear()
        {
            Console.Clear();
        }

        public void WriteLine()
        {  
            Console.WriteLine();
        }

        public void WriteLine2(string value)
        { 
            Console.WriteLine(value); 
        }

        public void WriteLine3(string format, object arg0)
        {
            object argument0 = arg0.Unwrap();
            Console.WriteLine(format, argument0);
        }

        public void WriteLine4(string format, object arg0, object arg1)
        {
            object argument0 = arg0.Unwrap();
            object argument1 = arg1.Unwrap();
            Console.WriteLine(format, argument0, argument1);
        }

        public void WriteLine5(string format, object arg0, object arg1, object arg2)
        {
            object argument0 = arg0.Unwrap();
            object argument1 = arg1.Unwrap();
            object argument2 = arg2.Unwrap();
            Console.WriteLine(format, argument0, argument1,argument2);
        }

        public void WriteLine6(string format, object arg0, object arg1, object arg2, object arg3)
        {
            object argument0 = arg0.Unwrap();
            object argument1 = arg1.Unwrap();
            object argument2 = arg2.Unwrap();
            object argument3 = arg3.Unwrap();
            Console.WriteLine(format, argument0, argument1, argument2, argument3);
        }

        public void WriteLine7(string format, [In] ref object[] arg)
        {
            object[] argument;
            if (arg == null)
            {
                argument = null;
            }
            else
            {
                argument = arg.Unwrap();
            }
            Console.WriteLine(format, argument);
        }

    }
}

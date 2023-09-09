using System.Runtime.InteropServices;

namespace DotNetLib.Extensions
{
    [ComVisible(false)]
    public interface IWrappedObject
    {
        object WrappedObject { get; }
    }
}

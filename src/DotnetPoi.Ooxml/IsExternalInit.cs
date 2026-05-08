// Polyfill for netstandard2.0 compatibility.
// Enables C# 9 record types compiled in this assembly.
#if NETSTANDARD2_0
namespace System.Runtime.CompilerServices
{
    internal static class IsExternalInit { }
}
#endif

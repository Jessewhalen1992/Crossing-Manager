#if ENABLE_MAP_OD
using Autodesk.Gis.Map;
#endif

namespace XingManager.Interop
{
    /// <summary>
    /// Optional shim for Map Object Data related helpers. In the base
    /// implementation this is a no-op to keep the core plug-in free from
    /// Map 3D dependencies.
    /// </summary>
    public static class MapOdShim
    {
#if ENABLE_MAP_OD
        // Placeholder for future Map Object Data integration.
#endif
    }
}
